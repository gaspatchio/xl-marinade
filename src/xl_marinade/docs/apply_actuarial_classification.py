# ABOUTME: Apply actuarial classification to all labelled bindings in overlay
# ABOUTME: Updates overlay database with classification results from ActuarialClassifier

import logging
import sqlite3
import sys
import time
from pathlib import Path
from typing import TYPE_CHECKING

from xl_marinade.core.labelling.mutation_engine import (
    MutationLogger,
    handle_set_reconciliation_flag,
)
from xl_marinade.core.labelling.overlay_database import (
    attach_ir_to_overlay,
    load_overlay_from_db,
    write_overlay_to_db,
)
from xl_marinade.core.parser import FormulaParser

from .classifiers.actuarial import classify_all_bindings
from .classifiers.hard_rules import validate_classification
from .classifiers.reconciliation import (
    classify_reconciliation_requirement,
    load_reconciliation_overrides,
)

if TYPE_CHECKING:
    pass  # RAG removed (NEW 2)

logger = logging.getLogger(__name__)


def _log_timing(step: str, elapsed: float, extra: str = "") -> None:
    suffix = f" ({extra})" if extra else ""
    logger.info(f"TIMING apply_actuarial_classification.{step}: {elapsed:.2f}s{suffix}")


def apply_actuarial_classification(
    overlay_db_path: str,
    ir_db_path: str,
    reconciliation_overrides_path: str | None = None,
    mutations_path: str | None = None,
) -> int:
    """
    Apply actuarial classification to all labelled bindings.

    Updates overlay database in-place with actuarial_class, reasoning, and confidence.
    Also appends generated mutations to mutations_path if provided.

    Args:
        overlay_db_path: Path to semantic_overlay.db
        ir_db_path: Path to Phase 1 IR database
        reconciliation_overrides_path: Optional path to overrides JSON
        mutations_path: Optional path to mutations.json to persist changes

    Returns:
        Number of bindings classified

    Raises:
        FileNotFoundError: If overlay or IR database doesn't exist
    """
    if not Path(overlay_db_path).exists():
        raise FileNotFoundError(f"Overlay database not found: {overlay_db_path}")

    if not Path(ir_db_path).exists():
        raise FileNotFoundError(f"IR database not found: {ir_db_path}")

    start_total = time.perf_counter()
    # Open overlay database. uri=True is required so attach_ir_to_overlay()'s
    # `ATTACH DATABASE 'file:...?mode=ro'` is parsed as a URI on this connection.
    overlay_conn = sqlite3.connect(overlay_db_path, uri=True)

    try:
        # Attach IR database
        attach_start = time.perf_counter()
        attach_ir_to_overlay(overlay_conn, ir_db_path)
        _log_timing("attach_ir", time.perf_counter() - attach_start)

        # Run classification
        logger.info("Running actuarial classification...")
        classify_start = time.perf_counter()
        results = classify_all_bindings(overlay_conn, ir_db_path)
        _log_timing(
            "classify_all_bindings",
            time.perf_counter() - classify_start,
            f"bindings={len(results)}",
        )

        # Load current overlay state to apply mutations
        try:
            load_start = time.perf_counter()
            overlay_state = load_overlay_from_db(overlay_db_path, ir_db_path)
            _log_timing("load_overlay_state", time.perf_counter() - load_start)
        except Exception as e:
            logger.warning(f"Failed to load overlay state: {e}")
            raise

        # Initialize mutation logger
        mutation_logger = MutationLogger()
        initial_mutation_count = 0

        if mutations_path and Path(mutations_path).exists():
            mutation_logger.load_from_file(mutations_path)
            initial_mutation_count = len(mutation_logger.mutations)
        elif overlay_state.mutations_applied:
            # Fallback to DB state for ID if no file provided
            max_id = max(m["mutation_id"] for m in overlay_state.mutations_applied)
            mutation_logger.next_id = max_id + 1

        # Generate override_binding mutations for each classification
        classification_count = 0

        try:
            sheet_rows = overlay_conn.execute("""
                SELECT sv.binding_id, b.sheet
                FROM semantic_variables sv
                LEFT JOIN ir.agent_bindings b ON sv.binding_id = b.binding_id
                WHERE sv.label IS NOT NULL AND sv.is_active = 1
            """).fetchall()
        except sqlite3.OperationalError:
            sheet_rows = overlay_conn.execute("""
                SELECT sv.binding_id, b.sheet
                FROM semantic_variables sv
                LEFT JOIN ir.bindings b ON sv.binding_id = b.binding_id
                WHERE sv.label IS NOT NULL AND sv.is_active = 1
            """).fetchall()
        sheet_map = {binding_id: sheet for binding_id, sheet in sheet_rows if sheet is not None}

        try:
            composite_rows = overlay_conn.execute("""
                SELECT cb.composite_id, b.sheet
                FROM composite_bindings cb
                JOIN semantic_variables sv ON sv.binding_id = cb.composite_id
                LEFT JOIN ir.agent_bindings b ON cb.ir_binding_id = b.binding_id
                WHERE sv.label IS NOT NULL AND sv.is_active = 1
                ORDER BY cb.composite_id, cb.ordinal
            """).fetchall()
        except sqlite3.OperationalError:
            composite_rows = overlay_conn.execute("""
                SELECT cb.composite_id, b.sheet
                FROM composite_bindings cb
                JOIN semantic_variables sv ON sv.binding_id = cb.composite_id
                LEFT JOIN ir.bindings b ON cb.ir_binding_id = b.binding_id
                WHERE sv.label IS NOT NULL AND sv.is_active = 1
                ORDER BY cb.composite_id, cb.ordinal
            """).fetchall()
        for composite_id, sheet in composite_rows:
            if composite_id not in sheet_map and sheet is not None:
                sheet_map[composite_id] = sheet

        cells_columns: list[str]
        try:
            cells_columns = [row[1] for row in overlay_conn.execute("PRAGMA ir.table_info(cells)")]
        except sqlite3.OperationalError:
            cells_columns = []
        has_agent_cells_view = False
        if "binding_id" not in cells_columns:
            try:
                overlay_conn.execute("SELECT 1 FROM ir.agent_cells LIMIT 1")
                has_agent_cells_view = True
            except sqlite3.OperationalError:
                has_agent_cells_view = False

        def _load_formula_binding_ids() -> set[str]:
            if "binding_id" in cells_columns:
                rows = overlay_conn.execute("""
                    SELECT DISTINCT c.binding_id
                    FROM ir.cells c
                    JOIN semantic_variables sv ON sv.binding_id = c.binding_id
                    WHERE sv.label IS NOT NULL AND sv.is_active = 1
                      AND c.formula_a1 IS NOT NULL AND c.formula_a1 != ''
                """).fetchall()
                formula_ids = {row[0] for row in rows}
                composite_rows = overlay_conn.execute("""
                    SELECT DISTINCT cb.composite_id
                    FROM composite_bindings cb
                    JOIN semantic_variables sv ON sv.binding_id = cb.composite_id
                    JOIN ir.cells c ON c.binding_id = cb.ir_binding_id
                    WHERE sv.label IS NOT NULL AND sv.is_active = 1
                      AND c.formula_a1 IS NOT NULL AND c.formula_a1 != ''
                """).fetchall()
                formula_ids.update(row[0] for row in composite_rows)
                return formula_ids
            if has_agent_cells_view:
                rows = overlay_conn.execute("""
                    SELECT DISTINCT ctb.binding_id
                    FROM ir.agent_cells ac
                    JOIN ir.cell_to_binding ctb ON ac.cell_id = ctb.cell_id
                    JOIN semantic_variables sv ON sv.binding_id = ctb.binding_id
                    WHERE sv.label IS NOT NULL AND sv.is_active = 1
                      AND ac.formula IS NOT NULL AND ac.formula != ''
                """).fetchall()
                formula_ids = {row[0] for row in rows}
                composite_rows = overlay_conn.execute("""
                    SELECT DISTINCT cb.composite_id
                    FROM composite_bindings cb
                    JOIN semantic_variables sv ON sv.binding_id = cb.composite_id
                    JOIN ir.cell_to_binding ctb ON ctb.binding_id = cb.ir_binding_id
                    JOIN ir.agent_cells ac ON ac.cell_id = ctb.cell_id
                    WHERE sv.label IS NOT NULL AND sv.is_active = 1
                      AND ac.formula IS NOT NULL AND ac.formula != ''
                """).fetchall()
                formula_ids.update(row[0] for row in composite_rows)
                return formula_ids
            return set()

        formula_binding_ids = _load_formula_binding_ids()

        loop_start = time.perf_counter()
        for binding_id, result in results.items():
            # Check if already classified by manual override
            binding = overlay_state.bindings.get(binding_id)
            if binding and binding.actuarial_class and binding.actuarial_class_confidence == 1.0:
                if (
                    binding.actuarial_class_reasoning
                    and "Manual" in binding.actuarial_class_reasoning
                ):
                    logger.info(
                        f"Skipping {binding_id}: Manually classified as {binding.actuarial_class}"
                    )
                    continue

            # Validate classification with hard rules
            # Get sheet name and formula status for validation
            sheet_name = sheet_map.get(binding_id) or "Unknown"
            has_formula = binding_id in formula_binding_ids

            # Apply hard rules validation
            final_class, final_confidence, rejection = validate_classification(
                binding_id=binding_id,
                sheet_name=sheet_name,
                has_formula=has_formula,
                proposed_class=result.actuarial_type.value,
                confidence=result.confidence,
            )

            # Build reasoning with rejection info if applicable
            reasoning = result.reasoning
            if rejection:
                reasoning = f"OVERRIDDEN: {rejection.rejected_reason}. Original: {reasoning}"
                logger.warning(f"Hard rule applied to {binding_id}: {rejection.rejected_reason}")

            # Generate mutation with potentially overridden class and confidence
            mutation_logger.override_binding(
                binding_id=binding_id,
                old_label=binding.label if binding else None,
                actuarial_class=final_class,
                reasoning=f"{reasoning} (confidence: {final_confidence:.2f})",
                classification_confidence=final_confidence,
            )
            classification_count += 1

        logger.info(f"Generated {classification_count} classification mutations")
        _log_timing(
            "generate_classification_mutations",
            time.perf_counter() - loop_start,
            f"mutations={classification_count}",
        )

        # Apply NEW mutations to overlay state
        # Classification confidence is now included in the mutation parameters
        new_mutations = mutation_logger.mutations[initial_mutation_count:]
        apply_start = time.perf_counter()
        for mutation in new_mutations:
            # We only have override_binding here so far
            if mutation["action"] == "override_binding":
                from xl_marinade.core.labelling.structural_mutations import handle_override_binding

                overlay_state = handle_override_binding(overlay_state, mutation, ir_db_path)
        _log_timing(
            "apply_classification_mutations",
            time.perf_counter() - apply_start,
            f"mutations={len(new_mutations)}",
        )

        # Log summary by type
        type_counts: dict[str, int] = {}
        for result in results.values():
            type_name = result.actuarial_type.value
            type_counts[type_name] = type_counts.get(type_name, 0) + 1

        for type_name, count in sorted(type_counts.items()):
            logger.info(f"  {type_name}: {count}")

        # Apply reconciliation classification
        logger.info("Classifying reconciliation requirements...")

        # Load overrides if provided
        overrides = {}
        if reconciliation_overrides_path:
            overrides = load_reconciliation_overrides(reconciliation_overrides_path)

        parser = FormulaParser()
        has_binding_id_column = "binding_id" in cells_columns
        ast_cache: dict[str, object | None] = {}
        _missing = object()

        def _load_binding_formulas(candidate_ids: list[str]) -> dict[str, str]:
            if not candidate_ids:
                return {}

            overlay_conn.execute("DROP TABLE IF EXISTS temp_recon_candidates")
            overlay_conn.execute("""
                CREATE TEMP TABLE temp_recon_candidates (
                    binding_id TEXT PRIMARY KEY
                )
            """)
            batch = []
            for binding_id in candidate_ids:
                batch.append((binding_id,))
                if len(batch) >= 10000:
                    overlay_conn.executemany(
                        "INSERT OR IGNORE INTO temp_recon_candidates (binding_id) VALUES (?)", batch
                    )
                    batch = []
            if batch:
                overlay_conn.executemany(
                    "INSERT OR IGNORE INTO temp_recon_candidates (binding_id) VALUES (?)", batch
                )

            try:
                if has_binding_id_column:
                    rows = overlay_conn.execute("""
                        WITH first_formula AS (
                            SELECT c.binding_id, MIN(c.cell_address_a1) AS min_addr
                            FROM ir.cells c
                            JOIN temp_recon_candidates t ON t.binding_id = c.binding_id
                            WHERE c.formula_a1 IS NOT NULL AND c.formula_a1 != ''
                            GROUP BY c.binding_id
                        )
                        SELECT c.binding_id, c.formula_a1
                        FROM ir.cells c
                        JOIN first_formula f
                          ON f.binding_id = c.binding_id
                         AND f.min_addr = c.cell_address_a1
                    """).fetchall()
                    return {row[0]: row[1] for row in rows if row[1]}
                if has_agent_cells_view:
                    rows = overlay_conn.execute("""
                        WITH first_formula AS (
                            SELECT ctb.binding_id, MIN(ac.cell_address) AS min_addr
                            FROM ir.agent_cells ac
                            JOIN ir.cell_to_binding ctb ON ac.cell_id = ctb.cell_id
                            JOIN temp_recon_candidates t ON t.binding_id = ctb.binding_id
                            WHERE ac.formula IS NOT NULL AND ac.formula != ''
                            GROUP BY ctb.binding_id
                        )
                        SELECT ctb.binding_id, ac.formula
                        FROM ir.agent_cells ac
                        JOIN ir.cell_to_binding ctb ON ac.cell_id = ctb.cell_id
                        JOIN first_formula f
                          ON f.binding_id = ctb.binding_id
                         AND f.min_addr = ac.cell_address
                    """).fetchall()
                    return {row[0]: row[1] for row in rows if row[1]}
                return {}
            finally:
                overlay_conn.execute("DROP TABLE IF EXISTS temp_recon_candidates")

        reconciliation_count = 0
        recon_candidates = []
        for binding_id, binding_overlay in overlay_state.bindings.items():
            if not binding_overlay.is_active:
                continue
            if not binding_overlay.label:
                continue
            if binding_overlay.actuarial_class not in ("Calculation", "Result"):
                continue
            recon_candidates.append(binding_id)
        formula_map = _load_binding_formulas(recon_candidates)

        # Iterate over in-memory overlay_state instead of querying DB
        # This ensures we see the classification changes we just applied (which are not yet in DB)
        recon_start = time.perf_counter()
        for binding_id in recon_candidates:
            binding_overlay = overlay_state.bindings.get(binding_id)
            if not binding_overlay or not binding_overlay.label:
                continue

            role = binding_overlay.actuarial_class
            label = binding_overlay.label

            # Get formula for this binding (first cell)
            formula = formula_map.get(binding_id)

            # Parse formula
            ast = None
            if formula:
                cached = ast_cache.get(formula, _missing)
                if cached is _missing:
                    try:
                        cached = parser.parse(formula)
                    except Exception as e:
                        logger.warning(f"Failed to parse formula for {binding_id}: {e}")
                        cached = None
                    ast_cache[formula] = cached
                ast = cached

            # Classify
            variable = {"variable_id": binding_id, "label": label}
            required, rationale = classify_reconciliation_requirement(
                variable, ast, role or "Unknown", overrides
            )  # noqa: E501

            # Generate mutation
            source = "override" if binding_id in overrides else "heuristic"
            mutation_logger.set_reconciliation_flag(binding_id, required, rationale, source)

            # Get the mutation object we just added
            mutation = mutation_logger.mutations[-1]

            # Apply to in-memory state using engine handler
            # This updates BindingOverlay and appends to state.mutations_applied
            handle_set_reconciliation_flag(overlay_state, mutation, ir_db_path)

            if required:
                reconciliation_count += 1

        _log_timing(
            "reconciliation_classification",
            time.perf_counter() - recon_start,
            f"required={reconciliation_count} candidates={len(recon_candidates)}",
        )

        # Persist mutations to file if path provided
        if mutations_path:
            mutation_logger.save(mutations_path)
            logger.info(f"Saved mutations to {mutations_path}")

        # Persist updated state to DB (includes both classification and reconciliation mutations)
        # Note: mutations_path parameter is not the actual mutations file since
        # classification generates mutations inline that are saved separately by the caller
        write_start = time.perf_counter()
        write_overlay_to_db(
            overlay_state, "<inline_classification_mutations>", ir_db_path, overlay_db_path
        )
        _log_timing("write_overlay_db", time.perf_counter() - write_start)

        msg = (
            f"Reconciliation classification complete: "
            f"{reconciliation_count}/{len(overlay_state.bindings)} require reconciliation"
        )
        logger.info(msg)
        _log_timing("total", time.perf_counter() - start_total)

        return classification_count

    finally:
        overlay_conn.close()


if __name__ == "__main__":
    # CLI usage
    import argparse

    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    parser = argparse.ArgumentParser(description="Apply actuarial classification to overlay.")
    parser.add_argument("overlay_db", help="Path to overlay database")
    parser.add_argument("ir_db", help="Path to IR database")
    parser.add_argument("--reconciliation-overrides", help="Path to reconciliation overrides JSON")

    args = parser.parse_args()

    count = apply_actuarial_classification(
        args.overlay_db,
        args.ir_db,
        reconciliation_overrides_path=args.reconciliation_overrides,
        mutations_path=None,  # CLI default, user can extend if needed
    )

    # stderr + ASCII: stdout is strict-encoded with the platform code page when
    # redirected or captured (cp1252 on Windows cannot encode U+2705), and this
    # banner runs AFTER the overlay DB is written — a crash here would make a
    # completed run look like a failure to any caller reading the exit code.
    print(f"\nClassified {count} bindings", file=sys.stderr)
