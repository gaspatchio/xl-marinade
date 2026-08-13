# ABOUTME: Sprint 7 pipeline integration - triage scoring + enrichment + mutation replay
# ABOUTME: Implements user-runnable, mutation-compliant Phase 2 workflow with --no-llm support

"""
Sprint 7 Pipeline Integration

This module integrates triage scoring, enrichment service, and inline mutation
management into the Phase 2 documentation pipeline.

## Stage Ordering (Hard Contract)

Sprint 7 enforces strict stage ordering to ensure upstream enrichment is consumed downstream:

1. **Label Enrichment** (triage-gated)
   - Triage scoring selects low-confidence bindings
   - LLM/fixture enrichment proposes label improvements
   - Append mutations → replay → overlay refreshed

2. **Structural Evaluation** (triage-gated)
   - Evaluate binding structure quality
   - LLM/fixture proposes merge/split/disable mutations
   - Append mutations → replay → overlay refreshed

3. **Segmentation Tagging** (NOT triage-gated)
   - Execute for EVERY active binding
   - Uses refreshed overlay labels from steps 1-2
   - Persist to segmentation.db (not mutations.json)
   - 3-retry policy; failures recorded, run continues

## CLI Options

- `--no-llm`: Hard disable LLM (replay-only mode)
- `--llm-provider`: `fixture` (deterministic) or `openai` (real LLM)
- `--llm-threshold`: Triage confidence threshold (default 0.60)
- `--llm-audit-log`: Enable enrichment_audit.jsonl logging

## Fast Extraction Requirement

Sprint 7 runners MUST invoke extractor in fast mode (`--fast`) and MUST fail
if `ir_metadata.build_mode != "fast"`.

See: docs/phase2_documentation_agent/backlog/sprint7/SPRINT7_PLAN.md
"""

import json
import logging
import re
import sqlite3
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from xl_marinade.core.db_uri import connect_read_only
from xl_marinade.core.labelling.mutation_engine import MutationLogger, replay_mutations
from xl_marinade.core.labelling.overlay_database import write_overlay_to_db
from xl_marinade.llm.enrichment_service import (
    EnrichmentProvider,
    EnrichmentService,
    LabelEnrichmentRequest,
    SegmentationTagRequest,
    StructuralFixCandidate,
    StructuralFixRequest,
)
from xl_marinade.llm.inline_mutation_manager import ConservationError, InlineMutationManager
from xl_marinade.llm.triage_confidence import (
    assess_all_triage_confidence,
    should_enrich,
)

logger = logging.getLogger(__name__)


@dataclass
class Sprint7PipelineConfig:
    """Configuration for Sprint 7 pipeline."""

    ir_db_path: str
    output_dir: Path

    # LLM configuration
    no_llm: bool = False
    llm_provider: EnrichmentProvider | None = None
    llm_threshold: float = 0.60
    llm_audit_log: bool = False

    # Stage control
    enable_label_enrichment: bool = True
    enable_structural_evaluation: bool = True
    enable_segmentation_tagging: bool = True

    # Safety valve for large corpora (optional)
    # If set, cap the number of selected bindings processed per stage (lowest score first).
    max_bindings_label: int | None = None
    max_bindings_structural: int | None = None
    max_bindings_segmentation: int | None = None

    # Optional explicit binding-id filters (for review packs / targeted runs).
    # When provided, the stage will only process binding_ids in these lists.
    label_binding_ids: list[str] | None = None
    structural_binding_ids: list[str] | None = None
    segmentation_binding_ids: list[str] | None = None


@dataclass
class Sprint7PipelineResult:
    """Result summary from Sprint 7 pipeline."""

    # Selection counts
    selected_binding_count_label: int = 0
    selected_binding_count_structural: int = 0

    # Enrichment counts
    enriched_binding_count_label: int = 0
    enriched_binding_count_structural: int = 0

    # Processed counts (after optional caps)
    processed_binding_count_label: int = 0
    processed_binding_count_structural: int = 0

    # Mutation counts
    mutations_appended_count_label: int = 0
    mutations_appended_count_structural: int = 0
    structural_mutations_appended_count: int = 0

    # Segmentation counts
    segmentation_tagged_count: int = 0
    segmentation_failed_count: int = 0

    # Provider info
    provider: str = "none"
    model: str | None = None

    # Conservation check
    cell_conservation_ok: bool = True

    # Failed bindings
    failed_bindings_count_label: int = 0
    failed_bindings_count_structural: int = 0


def validate_fast_extraction(ir_db_path: str) -> None:
    """
    Validate that IR was built in fast mode.

    Args:
        ir_db_path: Path to IR database

    Raises:
        ValueError: If build_mode != "fast"
    """
    conn = connect_read_only(ir_db_path)
    try:
        cursor = conn.execute("SELECT value FROM ir_metadata WHERE key='build_mode'")
        row = cursor.fetchone()

        if not row:
            raise ValueError("Missing ir_metadata.build_mode (IR database may be corrupted)")

        build_mode = row[0]
        if build_mode != "fast":
            raise ValueError(
                f"Sprint 7 requires fast extraction (build_mode='fast'), "
                f"but IR has build_mode='{build_mode}'. "
                f"Please re-extract with --fast flag."
            )

        logger.info(f"✅ IR build_mode validated: {build_mode}")
    finally:
        conn.close()


def _check_overlay_reflects_mutations(overlay_db_path: str, mutations_appended_count: int) -> bool:
    """
    Check if overlay database reflects applied mutations.

    This is an integration guardrail to detect "silent ignore" where mutations
    were appended but downstream outputs don't reflect them.

    Args:
        overlay_db_path: Path to semantic_overlay.db
        mutations_appended_count: Number of mutations appended in this run

    Returns:
        True if overlay reflects mutations, False otherwise
    """
    if mutations_appended_count == 0:
        return True  # No mutations to check

    if not Path(overlay_db_path).exists():
        logger.error(f"Overlay database not found: {overlay_db_path}")
        return False

    try:
        conn = sqlite3.connect(str(overlay_db_path))

        # Check if overlay has mutation_log entries
        cursor = conn.execute("SELECT COUNT(*) FROM mutation_log")
        mutation_count = cursor.fetchone()[0]

        # Check if semantic_variables has labels from mutations
        cursor = conn.execute("""
            SELECT COUNT(*) FROM semantic_variables 
            WHERE label_source IS NOT NULL
        """)
        labeled_count = cursor.fetchone()[0]

        conn.close()

        if mutation_count == 0:
            logger.error("Overlay has no mutation_log entries despite mutations being appended")
            return False

        if labeled_count == 0:
            logger.error("Overlay has no labeled bindings despite mutations being appended")
            return False

        logger.info(
            f"✅ Overlay reflects mutations: {mutation_count} mutations, {labeled_count} labeled bindings"
        )
        return True

    except Exception as e:
        logger.error(f"Failed to check overlay state: {e}")
        return False


_NUMERIC_LITERAL_RE = re.compile(r"(?<![A-Za-z_])(\d+(\.\d+)?)")


def _redact_numeric_literals(formula: str) -> str:
    return _NUMERIC_LITERAL_RE.sub("<REDACTED>", formula)


def _fetch_formula_evidence(
    ir_db_path: str,
    binding_id: str,
    limit: int = 10,
) -> tuple[str, str, dict[str, int], list[str]]:
    """
    Best-effort extraction of formula evidence for prompts.

    Returns:
        (semantic_formula, formula_r1c1, function_tokens, representative_formulas)
    """
    semantic_formula = ""
    formula_r1c1 = ""
    representative_formulas: list[str] = []
    function_tokens: dict[str, int] = {}

    conn = connect_read_only(ir_db_path)
    try:
        objects = {
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type IN ('table', 'view')"
            ).fetchall()
        }
        formulas: list[str] = []
        r1c1s: list[str] = []

        if "cell_to_binding" not in objects:
            return ("", "", {}, [])

        # Prefer agent_cells when present (legacy/agent schema)
        if "agent_cells" in objects:
            rows = conn.execute(
                """
                SELECT ac.formula, ac.formula_r1c1
                FROM cell_to_binding ctb
                JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
                WHERE ctb.binding_id = ? AND ac.formula IS NOT NULL
                LIMIT ?
                """,
                (binding_id, limit),
            ).fetchall()
            formulas = [row[0] for row in rows if row and row[0]]
            r1c1s = [row[1] for row in rows if row and len(row) > 1 and row[1]]
        # Fast schema: cells(formula_id) -> formulas(formula_r1c1, formula_a1_example)
        elif "cells" in objects and "formulas" in objects:
            rows = conn.execute(
                """
                SELECT f.formula_a1_example, f.formula_r1c1
                FROM cell_to_binding ctb
                JOIN cells c ON ctb.cell_id = c.cell_id
                JOIN formulas f ON c.formula_id = f.formula_id
                WHERE ctb.binding_id = ? AND f.formula_a1_example IS NOT NULL
                LIMIT ?
                """,
                (binding_id, limit),
            ).fetchall()
            formulas = [row[0] for row in rows if row and row[0]]
            r1c1s = [row[1] for row in rows if row and len(row) > 1 and row[1]]
        else:
            return ("", "", {}, [])

        if formulas:
            representative_formulas = [
                _redact_numeric_literals(f) for f in formulas[: min(5, len(formulas))]
            ]
            semantic_formula = representative_formulas[0]

        if r1c1s:
            formula_r1c1 = r1c1s[0]

        # Count common functions (cheap + deterministic)
        common_funcs = [
            "SUM",
            "AVERAGE",
            "IF",
            "IFS",
            "MIN",
            "MAX",
            "ROUND",
            "COUNT",
            "COUNTA",
            "COUNTIF",
            "COUNTIFS",
            "VLOOKUP",
            "HLOOKUP",
            "XLOOKUP",
            "INDEX",
            "MATCH",
        ]
        for formula in formulas:
            upper = formula.upper()
            for func in common_funcs:
                if func in upper:
                    function_tokens[func] = function_tokens.get(func, 0) + 1

        return (semantic_formula, formula_r1c1, function_tokens, representative_formulas)
    finally:
        conn.close()


def _seed_baseline_mutation_if_needed(
    mutation_logger: MutationLogger,
    ir_db_path: str,
) -> None:
    """
    Ensure mutations.json is non-empty so we can materialize a semantic_overlay.db.

    write_overlay_to_db() requires replay_mutations() output, which requires at least
    one mutation. In the real pipeline a baseline labeling step would produce those.
    For harness runs we seed a deterministic **no-op** mutation (set_orphan_status=false)
    so we don't pollute labels.
    """
    if mutation_logger.mutations:
        return

    conn = connect_read_only(ir_db_path)
    try:
        objects = {
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type IN ('table', 'view')"
            ).fetchall()
        }
        if "agent_bindings" in objects:
            row = conn.execute(
                "SELECT binding_id FROM agent_bindings ORDER BY binding_id LIMIT 1"
            ).fetchone()
        else:
            row = conn.execute(
                "SELECT binding_id FROM bindings ORDER BY binding_id LIMIT 1"
            ).fetchone()
    finally:
        conn.close()

    if not row:
        return

    binding_id = row[0]
    mutation_logger.append_mutation(
        "set_orphan_status",
        {"binding_id": binding_id, "is_orphan": False},
        {
            "reasoning": "Sprint 7 harness seed mutation (no-op) to enable overlay materialization",
            "knowledge_source": "test_harness",
            "sprint": 7,
        },
    )


def _replay_and_write_overlay(
    ir_db_path: str,
    mutations_path: Path,
    overlay_db_path: Path,
) -> None:
    # Enrichment resilience: LLM-proposed mutations can conflict (e.g. two merges
    # consuming the same binding); skip the offenders instead of discarding the
    # whole enrichment and falling back to deterministic docs.
    overlay = replay_mutations(ir_db_path, str(mutations_path), skip_conflicts=True)
    write_overlay_to_db(overlay, str(mutations_path), ir_db_path, str(overlay_db_path))


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8", newline="\n"
    )


def run_sprint7_pipeline(config: Sprint7PipelineConfig) -> Sprint7PipelineResult:
    """
    Run Sprint 7 enrichment pipeline with stage ordering.

    Args:
        config: Pipeline configuration

    Returns:
        Pipeline result summary

    Raises:
        ValueError: If fast extraction validation fails
        ConservationError: If structural mutation violates conservation
        RuntimeError: If overlay doesn't reflect applied mutations (AC3 guardrail)
    """
    # Validate fast extraction
    validate_fast_extraction(config.ir_db_path)

    # Setup paths
    mutations_path = config.output_dir / "mutations.json"
    overlay_db_path = config.output_dir / "semantic_overlay.db"
    segmentation_db_path = config.output_dir / "segmentation.db"
    rejected_proposals_path = config.output_dir / "rejected_proposals.json"
    audit_log_path = config.output_dir / "enrichment_audit.jsonl" if config.llm_audit_log else None

    # Initialize result
    result = Sprint7PipelineResult()
    rejected_proposals = []

    # Determine provider name
    if config.no_llm:
        result.provider = "none"
    elif config.llm_provider:
        result.provider = config.llm_provider.get_provider_name()
        # Get model name if OpenAI provider
        if hasattr(config.llm_provider, "get_model_name"):
            result.model = config.llm_provider.get_model_name()

    # Initialize mutation logger and load existing mutations (append-only contract)
    mutation_logger = MutationLogger()
    mutation_logger.load_from_file(str(mutations_path))

    # Initialize inline mutation manager
    mutation_manager = InlineMutationManager(config.ir_db_path, str(mutations_path))

    # Replay-only mode: materialize overlay from existing mutations, then exit.
    if config.no_llm:
        if mutation_logger.mutations:
            logger.info("Replay-only mode: re-materializing overlay from existing mutations.json")
            _replay_and_write_overlay(config.ir_db_path, mutations_path, overlay_db_path)
        else:
            logger.info("Replay-only mode: no existing mutations.json found; nothing to replay")

        with open(rejected_proposals_path, "w", encoding="utf-8", newline="\n") as f:
            json.dump([], f, indent=2)
        logger.info(f"Rejected proposals written to: {rejected_proposals_path}")
        return result

    # Ensure we can materialize an overlay projection for downstream stages.
    needs_overlay = (
        config.enable_label_enrichment
        or config.enable_structural_evaluation
        or config.enable_segmentation_tagging
    )
    if needs_overlay:
        _seed_baseline_mutation_if_needed(mutation_logger, config.ir_db_path)
        mutation_logger.save(str(mutations_path))
        _replay_and_write_overlay(config.ir_db_path, mutations_path, overlay_db_path)

    # ========================================================================
    # STAGE 1: Label Enrichment (triage-gated)
    # ========================================================================

    if config.enable_label_enrichment and not config.no_llm and config.llm_provider:
        logger.info("=" * 60)
        logger.info("STAGE 1: Label Enrichment (triage-gated)")
        logger.info("=" * 60)

        # Create enrichment service
        service = EnrichmentService(config.llm_provider, audit_log_path=audit_log_path)

        # Assess triage confidence
        logger.info("Assessing triage confidence for all bindings...")
        triage_results = assess_all_triage_confidence(config.ir_db_path, str(overlay_db_path))

        # Select bindings for enrichment
        selected_bindings = [
            tc for tc in triage_results if should_enrich(tc, threshold=config.llm_threshold)
        ]

        eligible_count = len(selected_bindings)
        logger.info(
            f"Selected {eligible_count} bindings for label enrichment (threshold={config.llm_threshold})"
        )

        if config.label_binding_ids is not None:
            requested = {str(bid) for bid in config.label_binding_ids}
            selected_bindings = [tc for tc in selected_bindings if tc.binding_id in requested]
            logger.info(
                f"Filtered label-enrichment selection to {len(selected_bindings)} bindings (explicit id filter)"
            )
        result.selected_binding_count_label = len(selected_bindings)

        if config.max_bindings_label is not None:
            selected_bindings = sorted(selected_bindings, key=lambda tc: tc.score)[
                : config.max_bindings_label
            ]
            logger.info(
                f"Capped label-enrichment selection to {len(selected_bindings)} bindings (max_bindings_label)"
            )
        result.processed_binding_count_label = len(selected_bindings)

        _write_json(
            config.output_dir / "processed_label_binding_ids.json",
            [tc.binding_id for tc in selected_bindings],
        )

        _write_json(
            config.output_dir / "triage_label.json",
            [
                {
                    "binding_id": tc.binding_id,
                    "score": tc.score,
                    "label_score": tc.label_score,
                    "structural_score": tc.structural_score,
                    "reasons": tc.reasons,
                }
                for tc in triage_results
            ],
        )

        # Enrich each selected binding
        for tc in selected_bindings:
            logger.info(f"Enriching label for binding {tc.binding_id} (score={tc.score:.2f})...")

            # Create enrichment request
            # (In real pipeline, this would load formula evidence from IR)
            # Load current label from overlay projection
            conn = sqlite3.connect(str(overlay_db_path))
            try:
                row = conn.execute(
                    "SELECT label FROM semantic_variables WHERE binding_id = ?",
                    (tc.binding_id,),
                ).fetchone()
            finally:
                conn.close()
            current_label = row[0] if row else None

            semantic_formula, formula_r1c1, function_tokens, representative_formulas = (
                _fetch_formula_evidence(
                    config.ir_db_path,
                    tc.binding_id,
                )
            )

            request = LabelEnrichmentRequest(
                binding_id=tc.binding_id,
                current_label=current_label,
                semantic_formula=semantic_formula or "SUM(<REDACTED>)",
                formula_r1c1=formula_r1c1 or "SUM(R[-1]C:RC)",
                function_tokens=function_tokens or {"SUM": 1},
                representative_formulas=representative_formulas or ["=SUM(A1:A10)"],
            )

            # Enrich with retry
            mutation_id, errors = service.enrich_label_with_retry(request, mutation_logger)

            if mutation_id is not None:
                result.enriched_binding_count_label += 1
                result.mutations_appended_count_label += 1
                logger.info(f"  ✓ Enriched (mutation_id={mutation_id})")
            elif errors:
                result.failed_bindings_count_label += 1
                rejected_proposals.append(
                    {"binding_id": tc.binding_id, "stage": "label_enrichment", "errors": errors}
                )
                logger.warning(f"  ✗ Failed after {len(errors)} attempts")
            else:
                logger.info("  - No change")

        # Save mutations and replay
        logger.info("Saving label enrichment mutations...")
        mutation_logger.save(str(mutations_path))
        logger.info("Replaying mutations to refresh overlay...")
        _replay_and_write_overlay(config.ir_db_path, mutations_path, overlay_db_path)
        logger.info(f"✅ Stage 1 complete: {result.enriched_binding_count_label} bindings enriched")

    # ========================================================================
    # STAGE 2: Structural Evaluation (triage-gated)
    # ========================================================================

    if config.enable_structural_evaluation and not config.no_llm and config.llm_provider:
        logger.info("=" * 60)
        logger.info("STAGE 2: Structural Evaluation (triage-gated)")
        logger.info("=" * 60)

        # Re-assess triage confidence (using refreshed overlay from Stage 1)
        logger.info("Re-assessing triage confidence after label enrichment...")
        triage_results = assess_all_triage_confidence(config.ir_db_path, str(overlay_db_path))

        # Select bindings for structural evaluation (structural_score only)
        selected_bindings = [
            tc for tc in triage_results if tc.structural_score < config.llm_threshold
        ]

        logger.info(f"Selected {len(selected_bindings)} bindings for structural evaluation")

        if config.structural_binding_ids is not None:
            requested = {str(bid) for bid in config.structural_binding_ids}
            selected_bindings = [tc for tc in selected_bindings if tc.binding_id in requested]
            logger.info(
                f"Filtered structural-evaluation selection to {len(selected_bindings)} bindings (explicit id filter)"
            )
        result.selected_binding_count_structural = len(selected_bindings)

        if config.max_bindings_structural is not None:
            selected_bindings = sorted(selected_bindings, key=lambda tc: tc.structural_score)[
                : config.max_bindings_structural
            ]
            logger.info(
                f"Capped structural-evaluation selection to {len(selected_bindings)} bindings (max_bindings_structural)"
            )
        result.processed_binding_count_structural = len(selected_bindings)

        _write_json(
            config.output_dir / "processed_structural_binding_ids.json",
            [tc.binding_id for tc in selected_bindings],
        )

        _write_json(
            config.output_dir / "triage_structural.json",
            [
                {
                    "binding_id": tc.binding_id,
                    "score": tc.score,
                    "label_score": tc.label_score,
                    "structural_score": tc.structural_score,
                    "reasons": tc.reasons,
                }
                for tc in triage_results
            ],
        )

        # Create enrichment service
        service = EnrichmentService(config.llm_provider, audit_log_path=audit_log_path)

        # Generate a minimal deterministic candidate set: merge with a neighbor binding.
        all_binding_ids = sorted(mutation_manager.binding_metadata.keys())
        binding_to_neighbor: dict[str, str] = {}
        for idx, bid in enumerate(all_binding_ids):
            if idx + 1 < len(all_binding_ids):
                binding_to_neighbor[bid] = all_binding_ids[idx + 1]

        for tc in selected_bindings:
            neighbor_id = binding_to_neighbor.get(tc.binding_id)
            if not neighbor_id:
                continue

            md = mutation_manager.binding_metadata.get(tc.binding_id, {})
            range_str = md.get("range_str", "")
            cell_count = int(md.get("cell_count", 0) or 0)

            conn = sqlite3.connect(str(overlay_db_path))
            try:
                row = conn.execute(
                    "SELECT label FROM semantic_variables WHERE binding_id = ?",
                    (tc.binding_id,),
                ).fetchone()
                neighbor_label_row = conn.execute(
                    "SELECT label FROM semantic_variables WHERE binding_id = ?",
                    (neighbor_id,),
                ).fetchone()
            finally:
                conn.close()
            current_label = row[0] if row else None
            neighbor_label = neighbor_label_row[0] if neighbor_label_row else None

            new_binding_id = f"merge_{tc.binding_id[:8]}_{neighbor_id[:8]}"
            candidates = [
                StructuralFixCandidate(
                    mutation_type="merge_bindings",
                    parameters={
                        "source_binding_ids": [tc.binding_id, neighbor_id],
                        "new_binding_id": new_binding_id,
                        "label": current_label or neighbor_label,
                    },
                    rationale="Merge a likely-fragmented adjacent binding pair to improve auditability.",
                )
            ]

            request = StructuralFixRequest(
                binding_id=tc.binding_id,
                current_label=current_label,
                range_str=range_str,
                cell_count=cell_count,
                formula_pattern="",
                neighbor_bindings=[
                    {
                        "binding_id": neighbor_id,
                        "range_str": mutation_manager.binding_metadata.get(neighbor_id, {}).get(
                            "range_str"
                        ),
                        "label": neighbor_label,
                    }
                ],
                candidates=candidates,
            )

            mutation_id, errors = service.propose_structural_fix_with_retry(
                request, mutation_logger
            )
            if mutation_id is not None:
                try:
                    mutation_manager.validate_and_save_mutations(mutation_logger, validate_last_n=1)
                    mutation_manager.replay_current_mutations()
                    _replay_and_write_overlay(config.ir_db_path, mutations_path, overlay_db_path)
                    result.enriched_binding_count_structural += 1
                    result.mutations_appended_count_structural += 1
                    result.structural_mutations_appended_count += 1
                    logger.info(f"  ✓ Structural fix applied (mutation_id={mutation_id})")
                except ConservationError as e:
                    mutation_logger.mutations.pop()
                    mutation_logger.next_id -= 1
                    result.cell_conservation_ok = False
                    rejected_proposals.append(
                        {
                            "binding_id": tc.binding_id,
                            "stage": "structural_fix",
                            "errors": [f"Conservation violation: {str(e)}"],
                        }
                    )
                    logger.error(f"  ✗ Conservation violation: {e}")
            elif errors:
                result.failed_bindings_count_structural += 1
                rejected_proposals.append(
                    {
                        "binding_id": tc.binding_id,
                        "stage": "structural_fix",
                        "errors": errors,
                    }
                )
                logger.warning(f"  ✗ Failed after {len(errors)} attempts")

        logger.info("✅ Stage 2 complete")

    # ========================================================================
    # STAGE 3: Segmentation Tagging (NOT triage-gated, every active binding)
    # ========================================================================

    if config.enable_segmentation_tagging and not config.no_llm and config.llm_provider:
        logger.info("=" * 60)
        logger.info("STAGE 3: Segmentation Tagging (all active bindings)")
        logger.info("=" * 60)

        # Create enrichment service
        service = EnrichmentService(config.llm_provider, audit_log_path=audit_log_path)

        # Get all active bindings from overlay
        conn = sqlite3.connect(str(overlay_db_path))
        cursor = conn.execute("""
            SELECT binding_id, label
            FROM semantic_variables
            WHERE is_active = 1
        """)
        active_bindings = [(row[0], row[1]) for row in cursor.fetchall()]
        conn.close()

        if config.segmentation_binding_ids is not None:
            requested = {str(bid) for bid in config.segmentation_binding_ids}
            active_bindings = [(bid, label) for (bid, label) in active_bindings if bid in requested]
            logger.info(
                f"Filtered segmentation selection to {len(active_bindings)} bindings (explicit id filter)"
            )

        if config.max_bindings_segmentation is not None:
            active_bindings = sorted(active_bindings, key=lambda x: x[0])[
                : config.max_bindings_segmentation
            ]
            logger.info(
                f"Capped segmentation selection to {len(active_bindings)} bindings (max_bindings_segmentation)"
            )

        _write_json(
            config.output_dir / "processed_segmentation_binding_ids.json",
            [bid for bid, _ in active_bindings],
        )

        logger.info(f"Tagging {len(active_bindings)} active bindings...")

        # Create segmentation database
        _create_segmentation_db(str(segmentation_db_path))

        # Tag each binding
        for binding_id, label in active_bindings:
            logger.info(f"Tagging binding {binding_id}...")

            semantic_formula, formula_r1c1, function_tokens, _ = _fetch_formula_evidence(
                config.ir_db_path,
                binding_id,
            )

            # Create request
            request = SegmentationTagRequest(
                binding_id=binding_id,
                current_label=label,
                semantic_formula=semantic_formula or "SUM(<REDACTED>)",
                formula_r1c1=formula_r1c1 or "SUM(R[-1]C:RC)",
                function_tokens=function_tokens or {"SUM": 1},
                allowed_tags=["Input", "Calculation", "Output", "Lookup"],
            )

            # Propose tag with retry
            tag_response, errors = service.propose_segmentation_tag_with_retry(request)

            if tag_response:
                result.segmentation_tagged_count += 1

                # Persist to segmentation.db
                _persist_segmentation_tag(
                    str(segmentation_db_path),
                    binding_id,
                    tag_response.tag,
                    tag_response.reasoning,
                    tag_response.confidence,
                    result.provider,
                    result.model,
                )

                logger.info(f"  ✓ Tagged: {tag_response.tag}")
            elif errors:
                result.segmentation_failed_count += 1
                rejected_proposals.append(
                    {"binding_id": binding_id, "stage": "segmentation_tagging", "errors": errors}
                )
                logger.warning(f"  ✗ Failed after {len(errors)} attempts")

        logger.info(f"✅ Stage 3 complete: {result.segmentation_tagged_count} bindings tagged")

    # ========================================================================
    # Write Summary and Rejected Proposals
    # ========================================================================

    # Write rejected proposals
    with open(rejected_proposals_path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(rejected_proposals, f, indent=2)

    logger.info(f"Rejected proposals written to: {rejected_proposals_path}")

    # ========================================================================
    # AC3 Guardrail: Check overlay reflects mutations
    # ========================================================================

    total_mutations_appended = (
        result.mutations_appended_count_label + result.mutations_appended_count_structural
    )

    if total_mutations_appended > 0:
        logger.info("=" * 60)
        logger.info("AC3 GUARDRAIL: Checking overlay reflects mutations")
        logger.info("=" * 60)

        if not _check_overlay_reflects_mutations(str(overlay_db_path), total_mutations_appended):
            raise RuntimeError(
                f"AC3 GUARDRAIL FAILED: {total_mutations_appended} mutations were appended "
                f"but overlay database does not reflect them. This indicates a 'silent ignore' "
                f"failure where enrichment was performed but downstream outputs won't show it. "
                f"Check that replay_mutations() was called after appending mutations."
            )

        logger.info(
            f"✅ AC3 guardrail passed: Overlay reflects {total_mutations_appended} mutations"
        )

    return result


def _create_baseline_overlay_with_dummy_mutation(
    ir_db_path: str, mutations_path: str, overlay_db_path: str
) -> None:
    """
    Create baseline overlay with a dummy mutation.

    This is needed because write_overlay_to_db requires at least one mutation
    to have been applied (to ensure it came from replay_mutations, not hand-crafted).

    In a real pipeline, the baseline overlay would be created by TwoPassLabellingEngine
    which generates real label mutations. For testing, we create a minimal valid overlay.
    """

    # Get first binding from IR to create a dummy mutation
    conn = connect_read_only(ir_db_path)
    cursor = conn.execute("SELECT name FROM sqlite_master WHERE type IN ('table', 'view')")
    objects = {row[0] for row in cursor.fetchall()}

    if "agent_bindings" in objects:
        cursor = conn.execute("SELECT binding_id FROM agent_bindings LIMIT 1")
    else:
        cursor = conn.execute("SELECT binding_id FROM bindings LIMIT 1")

    row = cursor.fetchone()
    conn.close()

    if not row:
        # No bindings in IR - create empty mutations file
        # (This will fail write_overlay_to_db but that's expected)
        with open(mutations_path, "w", encoding="utf-8", newline="\n") as f:
            json.dump([], f, indent=2)
        return

    binding_id = row[0]

    # Create a dummy mutation (set_label with placeholder)
    dummy_mutation = {
        "mutation_id": 1,
        "timestamp": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "action": "set_label",
        "parameters": {"binding_id": binding_id, "old": None, "new": "_baseline_placeholder"},
        "metadata": {
            "reasoning": "Baseline placeholder for Sprint 7 testing",
            "knowledge_source": "test_harness",
            "sprint": 7,
        },
    }

    # Write mutations file
    with open(mutations_path, "w", encoding="utf-8", newline="\n") as f:
        json.dump([dummy_mutation], f, indent=2)

    # Replay to create overlay (enrichment path: skip conflicting mutations)
    overlay = replay_mutations(ir_db_path, mutations_path, skip_conflicts=True)
    write_overlay_to_db(overlay, mutations_path, ir_db_path, overlay_db_path)


def _create_segmentation_db(db_path: str) -> None:
    """Create segmentation database schema."""
    conn = sqlite3.connect(db_path)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS segmentation_tags (
            binding_id TEXT PRIMARY KEY,
            tag TEXT NOT NULL,
            reasoning TEXT,
            confidence REAL,
            provider TEXT,
            model TEXT,
            timestamp TEXT
        )
    """)

    conn.commit()
    conn.close()


def _persist_segmentation_tag(
    db_path: str,
    binding_id: str,
    tag: str,
    reasoning: str,
    confidence: float,
    provider: str,
    model: str | None,
) -> None:
    """Persist segmentation tag to database."""

    conn = sqlite3.connect(db_path)

    conn.execute(
        """
        INSERT OR REPLACE INTO segmentation_tags
        (binding_id, tag, reasoning, confidence, provider, model, timestamp)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """,
        (
            binding_id,
            tag,
            reasoning,
            confidence,
            provider,
            model,
            datetime.now(UTC).isoformat(),
        ),
    )

    conn.commit()
    conn.close()
