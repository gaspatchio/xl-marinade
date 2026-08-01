# ABOUTME: Two-pass context-aware labelling engine for Sprint 2
# ABOUTME: Extends simple labeller with top-down and bottom-up refinement using dependency graph

import json
import logging
import re
import sqlite3
import time
from pathlib import Path
from typing import Any

from xl_marinade.core.labelling.mutation_engine import (
    MutationLogger,
    OverlayState,
    replay_mutations,
)
from xl_marinade.core.labelling.overlay_database import write_overlay_to_db
from xl_marinade.core.labelling.simple_labeller import Binding, is_numeric, simple_label_selection
from xl_marinade.core.ref_converter import col_num_to_letter, parse_cell_address

from .dependency_traversal import (
    build_dependency_graph,
    get_children,
    get_parents,
    get_topological_order,
)
from .label_generator import generate_label_from_formula, should_generate_label
from .refiners.sheet_context import SheetContextRefiner

logger = logging.getLogger(__name__)

# A redundant trailing location suffix (Row N / Column X) that the first
# disambiguation pass may have appended but which failed to separate a group of
# siblings sharing that row/column. Stripped before re-suffixing with the
# distinguishing cell coordinate.
_LOC_SUFFIX_RE = re.compile(r"\s*\((?:Row \d+|Column [A-Z]+)\)\s*$")


class TwoPassLabellingEngine:
    """
    Two-pass labelling engine with context-aware refinement.

    Pass 1 (Top-Down): Process bindings from roots to leaves, using parent context
    Pass 2 (Bottom-Up): Process bindings from leaves to roots, using child context

    Inherits core scoring logic from simple_labeller, adds context injection.

    ARCHITECTURAL NOTE: This engine is now STATELESS - it does not maintain a
    private overlay dict. All state is managed through:
    1. MutationLogger (write path)
    2. OverlayState (read-only projection, rebuilt via replay_mutations)

    This enforces Single Source of Truth: mutations.json is authoritative.
    """

    def __init__(self, ir_db_path: str):
        """
        Initialize two-pass labelling engine.

        Args:
            ir_db_path: Path to Phase 1 IR database
        """
        if not Path(ir_db_path).exists():
            raise FileNotFoundError(f"IR database not found: {ir_db_path}")

        self.ir_db_path = ir_db_path
        # Type note: graph is either nx.DiGraph or dict[str, list] (fallback)
        self.graph: Any = None  # nx.DiGraph or dict-based graph structure
        self.bindings: dict[str, Binding] = {}  # binding_id -> Binding
        self.mutation_logger = MutationLogger()
        # Read-only projection of current state (rebuilt via replay after each pass)
        self.overlay_state: OverlayState = OverlayState()
        self._last_synced_mutation_count = 0
        self._parent_cache: dict[str, list[str]] | None = None
        self._child_cache: dict[str, list[str]] | None = None
        # Formula family groupings (family_id -> [binding_ids])
        self.formula_families: dict[str, list[str]] = {}
        self._family_id_map: dict[str, str] = {}  # binding_id -> family_id
        self._family_representatives: dict[str, str] = {}  # family_id -> representative_binding_id

    def _log_timing(self, step: str, elapsed: float, extra: str = "") -> None:
        suffix = f" ({extra})" if extra else ""
        logger.info(f"TIMING two_pass_labeller.{step}: {elapsed:.2f}s{suffix}")

    def load_bindings(self) -> None:
        """Load all bindings from IR database.

        Note: This loads ALL bindings initially. Orphaned bindings (those not in
        the dependency graph) will be filtered out in build_graph().
        """
        start_time = time.perf_counter()
        conn = sqlite3.connect(f"file:{self.ir_db_path}?mode=ro", uri=True)

        # Use agent_bindings view for compatibility with both legacy and fast schemas
        try:
            rows = conn.execute("""
                SELECT binding_id, sheet, address, spatial_candidates
                FROM agent_bindings
            """).fetchall()
        except sqlite3.OperationalError as exc:
            conn.close()
            raise RuntimeError(
                "agent_bindings view not found; bindings persistence is required for "
                "documentation agent compatibility."
            ) from exc

        if not rows:
            conn.close()
            raise ValueError(
                "No bindings found in agent_bindings; check fast pipeline persistence."
            )

        self.bindings = {
            row[0]: Binding(
                binding_id=row[0],
                sheet=row[1],
                address_a1=row[2],
                debug_label=None,  # Not available in agent view
                label_candidates_json=row[3] if row[3] else "{}",
            )
            for row in rows
        }

        self._load_formula_families(conn)
        self._hydrate_missing_label_candidates(conn)
        conn.close()
        logger.info(f"Loaded {len(self.bindings)} total bindings from IR")
        self._log_timing(
            "load_bindings", time.perf_counter() - start_time, f"bindings={len(self.bindings)}"
        )

    def _load_formula_families(self, conn: sqlite3.Connection) -> None:
        """Load formula family mappings from IR database (if available)."""
        try:
            tables = {
                r[0]
                for r in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                ).fetchall()
            }
            if "formula_families" not in tables:
                return

            rows = conn.execute("""
                SELECT ffm.binding_id, ff.family_id, ff.representative_binding_id
                FROM formula_family_members ffm
                JOIN formula_families ff ON ffm.family_id = ff.family_id
                ORDER BY ff.family_id, ffm.ordinal
            """).fetchall()

            for binding_id, family_id, representative_id in rows:
                self._family_id_map[binding_id] = family_id
                self._family_representatives[family_id] = representative_id
                if family_id not in self.formula_families:
                    self.formula_families[family_id] = []
                self.formula_families[family_id].append(binding_id)

            if self.formula_families:
                logger.info(
                    f"Loaded {len(self.formula_families)} formula families "
                    f"covering {len(self._family_id_map)} bindings"
                )
        except sqlite3.OperationalError:
            pass

    def _hydrate_missing_label_candidates(self, conn: sqlite3.Connection) -> None:
        """Populate basic scan candidates when label candidates are missing."""
        missing = []
        scan_specs: list[tuple[Binding, str, int, int, list[str]]] = []
        needed_positions: dict[str, set[tuple[int, int]]] = {}

        for binding in self.bindings.values():
            try:
                data = json.loads(binding.label_candidates_json)
            except (json.JSONDecodeError, TypeError):
                data = {}
            if not isinstance(data, dict) or not data.get("label_candidates"):
                missing.append(binding)

                parsed = parse_cell_address(binding.address_a1)
                row = int(parsed.get("row", 0)) if parsed else 0
                col = int(parsed.get("col", 0)) if parsed else 0
                height = int(parsed.get("height", 1)) if parsed else 1
                width = int(parsed.get("width", 1)) if parsed else 1
                sheet = binding.sheet

                if row <= 0 or col <= 0 or not sheet:
                    continue

                directions = []
                if height == 1 and width > 1:
                    directions.append("left")
                elif width == 1 and height > 1:
                    directions.append("above")
                else:
                    directions.extend(["left", "above"])

                scan_specs.append((binding, sheet, row, col, directions))
                for direction in directions:
                    for offset in range(1, 8):
                        if direction == "left":
                            r = row
                            c = col - offset
                        else:
                            r = row - offset
                            c = col

                        if r <= 0 or c <= 0:
                            break

                        needed_positions.setdefault(sheet, set()).add((r, c))

        if not missing:
            return

        start_time = time.perf_counter()
        cell_rows = []
        if needed_positions:
            conn.execute("DROP TABLE IF EXISTS temp_label_scan_positions")
            conn.execute("""
                CREATE TEMP TABLE temp_label_scan_positions (
                    sheet TEXT NOT NULL,
                    row INTEGER NOT NULL,
                    col INTEGER NOT NULL,
                    PRIMARY KEY (sheet, row, col)
                )
            """)

            batch = []
            for sheet, positions in needed_positions.items():
                for row, col in positions:
                    batch.append((sheet, row, col))
                    if len(batch) >= 10000:
                        conn.executemany(
                            """
                            INSERT OR IGNORE INTO temp_label_scan_positions (sheet, row, col)
                            VALUES (?, ?, ?)
                        """,
                            batch,
                        )
                        batch = []

            if batch:
                conn.executemany(
                    """
                    INSERT OR IGNORE INTO temp_label_scan_positions (sheet, row, col)
                    VALUES (?, ?, ?)
                """,
                    batch,
                )

            cell_rows = conn.execute("""
                SELECT c.sheet, c.row, c.col, c.value, c.formula, c.data_type
                FROM agent_cells c
                JOIN temp_label_scan_positions p
                  ON p.sheet = c.sheet
                 AND p.row = c.row
                 AND p.col = c.col
                WHERE (c.value IS NOT NULL AND c.value != '')
                   OR (c.formula IS NOT NULL AND c.formula != '')
            """).fetchall()
            conn.execute("DROP TABLE IF EXISTS temp_label_scan_positions")

        cell_index: dict[str, dict[tuple[int, int], tuple[str | None, str | None, str | None]]] = {}
        for sheet, row, col, value, formula, data_type in cell_rows:
            cell_index.setdefault(sheet, {})[(row, col)] = (value, formula, data_type)

        for binding, sheet, row, col, directions in scan_specs:
            candidates = []
            for direction in directions:
                candidate = self._scan_label_candidate(sheet, row, col, direction, cell_index)
                if candidate:
                    candidates.append(candidate)

            if candidates:
                binding.label_candidates_json = json.dumps(
                    {"label_candidates": candidates, "axis_labels": []}, sort_keys=True
                )

        elapsed = time.perf_counter() - start_time
        self._log_timing(
            "hydrate_label_candidates", elapsed, f"missing={len(missing)} cells={len(cell_rows)}"
        )
        logger.info(f"Generated scan candidates for {len(missing)} bindings")

    def _scan_label_candidate(
        self,
        sheet: str,
        row: int,
        col: int,
        direction: str,
        cell_index: dict[str, dict[tuple[int, int], tuple[str | None, str | None, str | None]]],
    ) -> dict[str, Any] | None:
        """Build a scan_left/scan_above candidate from adjacent cells."""
        cells = []
        literals = []

        for offset in range(1, 8):
            if direction == "left":
                r = row
                c = col - offset
            else:
                r = row - offset
                c = col

            if r <= 0 or c <= 0:
                break

            cell_data = cell_index.get(sheet, {}).get((r, c))
            if not cell_data:
                continue
            value, formula, data_type = cell_data
            literal = value if value not in (None, "") else formula
            if literal not in (None, ""):
                literals.append(str(literal))

            cells.append(
                {
                    "address": f"{sheet}!{col_num_to_letter(c)}{r}",
                    "value": value,
                    "formula": formula,
                    "dtype": data_type,
                }
            )

        if not literals:
            return None

        return {
            "type": f"scan_{direction}",
            "address": cells[0]["address"] if cells else "",
            "literals": literals,
            "cells": cells,
            "format_tokens": {},
            "merged_span": None,
        }

    def build_graph(self) -> None:
        """Build dependency graph and identify orphans.

        After building the dependency graph, this:
        1. Identifies bindings that are NOT reachable from any root (orphans)
        2. Generates mutations for any changes in orphan status
        3. Syncs overlay state to persist orphan flags
        4. Filters self.bindings to only semantically referenced bindings

        Orphaned bindings (structurally detected but not semantically referenced)
        are excluded from labelling but preserved in the IR for future processing
        (e.g., key-lookup association via Stories 13/14).
        """
        start_time = time.perf_counter()
        self.graph = build_dependency_graph(self.ir_db_path)

        # Get bindings that are in the dependency graph
        try:
            import networkx as nx

            if isinstance(self.graph, nx.DiGraph):
                bindings_in_graph = set(self.graph.nodes())
            else:
                # Dict-based graph: include both sources and targets
                bindings_in_graph = set(self.graph.keys()) if self.graph else set()
                # Also include target-only bindings from reverse graph
                if "__reverse__" in self.graph:
                    bindings_in_graph.update(self.graph["__reverse__"].keys())
                    # Remove the special key from the set
                    bindings_in_graph.discard("__reverse__")
        except ImportError:
            # Dict-based graph: include both sources and targets
            bindings_in_graph = set(self.graph.keys()) if self.graph else set()
            # Also include target-only bindings from reverse graph
            if "__reverse__" in self.graph:
                bindings_in_graph.update(self.graph["__reverse__"].keys())
                # Remove the special key from the set
                bindings_in_graph.discard("__reverse__")

        # Calculate current orphan set
        current_orphans = set(self.bindings.keys()) - bindings_in_graph

        # Generate orphan status mutations (Story 15)
        # Compare with previously-stored orphan status (if any)
        for binding_id in self.bindings:
            is_orphan = binding_id in current_orphans
            previously_orphan = (
                binding_id in self.overlay_state.bindings
                and self.overlay_state.bindings[binding_id].is_orphan
            )

            # Only log changes (idempotent: avoid duplicate mutations)
            if is_orphan != previously_orphan:
                self.mutation_logger.set_orphan_status(
                    binding_id=binding_id,
                    is_orphan=is_orphan,
                    reasoning=f"{'Not reachable from' if is_orphan else 'Reachable from'} root cells",
                )

        # Log orphan count (Story 15 AC)
        logger.info(
            f"Identified {len(current_orphans)} orphan bindings "
            f"(out of {len(self.bindings)} total bindings)"
        )

        # Sync overlay state to apply orphan mutations before labelling
        if self.mutation_logger.mutations:
            self._sync_overlay_state()

        # Count before filtering
        total_bindings = len(self.bindings)

        # Filter to only semantically referenced bindings
        self.bindings = {
            bid: binding for bid, binding in self.bindings.items() if bid in bindings_in_graph
        }

        # Log filtering results
        orphaned_count = total_bindings - len(self.bindings)
        logger.info(
            f"Built dependency graph: {len(self.bindings)} semantically referenced bindings, "
            f"{orphaned_count} orphans marked (excluded from labelling)"
        )
        self._build_graph_context_cache()
        self._log_timing("build_graph", time.perf_counter() - start_time)

    def _build_graph_context_cache(self) -> None:
        """Precompute parent/child lists for faster context lookups."""
        if not self.graph:
            self._parent_cache = {}
            self._child_cache = {}
            return

        try:
            import networkx as nx

            if isinstance(self.graph, nx.DiGraph):
                self._parent_cache = {
                    node: sorted(self.graph.predecessors(node)) for node in self.graph.nodes()
                }
                self._child_cache = {
                    node: sorted(self.graph.successors(node)) for node in self.graph.nodes()
                }
                return
        except ImportError:
            pass

        reverse_graph: dict[str, set[str]] = self.graph.get("__reverse__", {})
        self._child_cache = {}
        for node, children in self.graph.items():
            if node == "__reverse__":
                continue
            self._child_cache[node] = sorted(children)

        self._parent_cache = {node: sorted(parents) for node, parents in reverse_graph.items()}

        all_nodes = set(self._child_cache.keys()) | set(self._parent_cache.keys())
        for node in all_nodes:
            self._child_cache.setdefault(node, [])
            self._parent_cache.setdefault(node, [])

    def _get_parent_context(self, binding_id: str) -> str:
        """
        Get context string from parent labels (for Pass 1).

        Queries the current OverlayState (read-only projection) for parent labels.

        Args:
            binding_id: Binding to get parent context for

        Returns:
            Context string to inject into prompt (empty if no parents)
        """
        if not self.graph:
            return ""

        if self._parent_cache is not None:
            parents = self._parent_cache.get(binding_id, [])
        else:
            parents = get_parents(self.graph, binding_id)
        if not parents:
            return ""

        # Get labels from overlay_state (read-only projection)
        parent_labels = []
        for parent_id in parents:
            if parent_id in self.overlay_state.bindings:
                binding_overlay = self.overlay_state.bindings[parent_id]
                if binding_overlay.label:
                    # We don't have score in OverlayState, use label_source as proxy
                    label = binding_overlay.label
                    parent_labels.append(f"{label}")

        if not parent_labels:
            return ""

        # Format as context string
        context = "Parent variables: " + ", ".join(parent_labels)
        return context

    def _get_child_context(self, binding_id: str) -> str:
        """
        Get context string from child labels (for Pass 2).

        Queries the current OverlayState (read-only projection) for child labels.

        Args:
            binding_id: Binding to get child context for

        Returns:
            Context string to inject into prompt (empty if no children)
        """
        if not self.graph:
            return ""

        if self._child_cache is not None:
            children = self._child_cache.get(binding_id, [])
        else:
            children = get_children(self.graph, binding_id)
        if not children:
            return ""

        # Get labels from overlay_state (read-only projection)
        child_labels = []
        for child_id in children:
            if child_id in self.overlay_state.bindings:
                binding_overlay = self.overlay_state.bindings[child_id]
                if binding_overlay.label:
                    label = binding_overlay.label
                    child_labels.append(f"{label}")

        if not child_labels:
            return ""

        # Format as context string
        context = "Child variables (dependencies): " + ", ".join(child_labels)
        return context

    def _process_binding_pass_1(self, binding_id: str) -> None:
        """
        Process binding in Pass 1 (top-down with parent context).

        Logs mutations only - does NOT maintain local state.
        State is synced via replay_mutations after pass completes.

        Args:
            binding_id: Binding to process
        """
        binding = self.bindings[binding_id]

        # Get parent context (from replayed state)
        parent_context = self._get_parent_context(binding_id)

        # Select label using simple_label_selection (Sprint 1 logic)
        # Note: In real implementation with LLM, we would inject parent_context into prompt
        # For now, we use the same logic as simple labeller
        label, score, candidate_type = simple_label_selection(binding)

        # Log mutation (write path - Single Source of Truth)
        reasoning = f"Pass 1 (top-down): Selected from IR candidates (type: {candidate_type})"
        if parent_context:
            reasoning += f". Context: {parent_context}"

        self.mutation_logger.set_label(
            binding_id=binding_id,
            old_label=None,
            new_label=label,
            reasoning=reasoning,
            knowledge_source="ir_candidates",
            confidence=score,
        )

        # Propagate label to formula family siblings
        family_id = self._family_id_map.get(binding_id)
        if family_id and self._family_representatives.get(family_id) == binding_id:
            siblings = [
                bid for bid in self.formula_families.get(family_id, []) if bid != binding_id
            ]
            if siblings:
                self.mutation_logger.propagate_family_label(
                    formula_family_id=family_id,
                    representative_binding_id=binding_id,
                    label=label,
                    reasoning=f"Family label propagation from representative {binding_id}",
                    family_sibling_ids=siblings,
                )
                logger.info(
                    f"Pass 1: Propagated label '{label}' from {binding_id} "
                    f"to {len(siblings)} family siblings (family={family_id})"
                )

        logger.debug(f"Pass 1: {binding_id} -> {label} (score: {score:.2f})")

    def _refine_with_rag(
        self, binding_id: str, current_label: str, current_score: float
    ) -> tuple[str, float, str] | None:
        """
        Refine label using RAG knowledge base.

        Args:
            binding_id: Binding ID
            current_label: Current label candidate
            current_score: Current confidence score

        Returns:
            Tuple (new_label, new_score, reasoning_suffix) if refined, else None
        """
        # RAG refinement removed (NEW 2). No-op: the deterministic labels stand;
        # low-confidence cases are covered by LLM enrichment (sprint7) when a key is set.
        return None

    def _process_binding_pass_2(self, binding_id: str) -> None:
        """
        Process binding in Pass 2 (bottom-up with child context).

        Queries overlay_state for current label, logs mutations if refinement occurs.

        Args:
            binding_id: Binding to process
        """
        binding = self.bindings[binding_id]

        # Get current label from overlay_state (replayed from Pass 1)
        if binding_id not in self.overlay_state.bindings:
            logger.warning(f"Pass 2: Binding {binding_id} not in overlay_state, skipping")
            return

        binding_overlay = self.overlay_state.bindings[binding_id]
        if not binding_overlay.label:
            logger.warning(f"Pass 2: Binding {binding_id} has no label, skipping")
            return

        old_label = binding_overlay.label
        # We don't have score stored in OverlayState - use heuristic from candidate type
        # This is acceptable since we're just comparing relative improvement
        old_score, old_type = simple_label_selection(binding)[1:3]

        # Get child context (from replayed state)
        child_context = self._get_child_context(binding_id)

        # Re-evaluate label with child context
        # Note: In real implementation with LLM, we would inject child_context into prompt
        # For now, we use the same logic as simple labeller
        new_label, new_score, new_type = simple_label_selection(binding)

        # --- RAG Refinement Start ---
        rag_result = self._refine_with_rag(binding_id, new_label, new_score)
        rag_reasoning = ""

        if rag_result:
            rag_label, rag_score, rag_suffix = rag_result
            # If RAG provides a better score than what we have
            if rag_score > new_score:
                new_label = rag_label
                new_score = rag_score
                new_type = "rag_knowledge"
                rag_reasoning = f". {rag_suffix}"
        # --- RAG Refinement End ---

        # Only update if new score is significantly better (threshold: 0.1)
        # This prevents unnecessary churn
        score_improvement = new_score - old_score

        if score_improvement > 0.1:
            # Log mutation (write path)
            reasoning = (
                f"Pass 2 (bottom-up): Refined label (score improved by {score_improvement:.2f})"
            )
            if child_context:
                reasoning += f". Context: {child_context}"

            if rag_reasoning:
                reasoning += rag_reasoning

            self.mutation_logger.set_label(
                binding_id=binding_id,
                old_label=old_label,
                new_label=new_label,
                reasoning=reasoning,
                knowledge_source=(
                    "ir_candidates" if new_type != "rag_knowledge" else "rag_knowledge"
                ),
                confidence=new_score,
            )

            logger.debug(
                f"Pass 2: {binding_id} -> {new_label} "
                f"(score: {new_score:.2f}, improved from {old_score:.2f})"
            )
        else:
            logger.debug(f"Pass 2: {binding_id} -> kept {old_label} (no significant improvement)")

    def run_pass_1_top_down(self) -> None:
        """
        Run Pass 1: Top-down traversal with parent context.

        Processes bindings in topological order (roots to leaves).
        After processing, syncs overlay_state by replaying all mutations.
        """
        # `is None` (not falsy): an all-orphans model yields an *empty* graph,
        # and an empty nx.DiGraph is falsy while the dict fallback
        # ({'__reverse__': {}}) is truthy. Only None means build_graph() never ran.
        if self.graph is None:
            raise RuntimeError("Must call build_graph() before running passes")

        start_time = time.perf_counter()
        # Get topological order (roots to leaves)
        order = get_topological_order(self.graph)

        logger.info(f"Pass 1: Processing {len(order)} bindings in topological order")

        # Track family siblings whose representative has already been processed
        propagated_siblings: set[str] = set()

        for binding_id in order:
            if binding_id in self.bindings:
                # Skip non-representative family siblings (label already propagated)
                if binding_id in propagated_siblings:
                    continue
                self._process_binding_pass_1(binding_id)
                # Mark siblings as propagated if this was a representative
                family_id = self._family_id_map.get(binding_id)
                if family_id and self._family_representatives.get(family_id) == binding_id:
                    for sid in self.formula_families.get(family_id, []):
                        if sid != binding_id:
                            propagated_siblings.add(sid)

        # CRITICAL: Sync overlay_state by replaying mutations
        # This rebuilds the read-only projection from the authoritative mutation log
        self._sync_overlay_state()

        logger.info(
            f"Pass 1 complete: Logged {len(self.mutation_logger.mutations)} mutations, "
            f"synced {len(self.overlay_state.bindings)} bindings"
        )
        self._log_timing("pass_1_top_down", time.perf_counter() - start_time)

    def run_pass_2_bottom_up(self) -> None:
        """
        Run Pass 2: Bottom-up traversal with child context.

        Processes bindings in reverse topological order (leaves to roots).
        After processing, syncs overlay_state by replaying all mutations.
        """
        # `is None` (not falsy): see run_pass_1_top_down — an empty graph is a
        # valid built state (all-orphans model) and must not be treated as unbuilt.
        if self.graph is None:
            raise RuntimeError("Must call build_graph() before running passes")

        start_time = time.perf_counter()
        # Get topological order and reverse it (leaves to roots)
        order = get_topological_order(self.graph)
        reverse_order = list(reversed(order))

        logger.info(
            f"Pass 2: Processing {len(reverse_order)} bindings in reverse topological order"
        )

        for binding_id in reverse_order:
            if binding_id in self.bindings:
                # Skip non-representative family siblings (label already propagated)
                family_id = self._family_id_map.get(binding_id)
                if family_id:
                    rep = self._family_representatives.get(family_id)
                    if rep and binding_id != rep:
                        continue
                self._process_binding_pass_2(binding_id)

        # CRITICAL: Sync overlay_state by replaying mutations
        self._sync_overlay_state()

        logger.info("Pass 2 complete")
        self._log_timing("pass_2_bottom_up", time.perf_counter() - start_time)

    def _sync_overlay_state(self) -> None:
        """
        Sync overlay_state by replaying all mutations from mutation_logger.

        This rebuilds the read-only projection from the authoritative mutation log.
        Called after each pass to ensure consistent state.

        Performance Note: This replays the entire mutation log from scratch (O(n) mutations).
        Becomes bottleneck at >500 mutations. Future optimization: implement incremental
        replay or streaming projection updates.

        Raises:
            Exception: If replay_mutations fails, overlay_state remains in previous state
        """
        import json
        import tempfile

        current_count = len(self.mutation_logger.mutations)
        if current_count == self._last_synced_mutation_count:
            logger.debug(
                "Skipping sync_overlay_state: no new mutations (count=%d)",
                current_count,
            )
            return

        start_time = time.perf_counter()
        # Write mutations to temporary file for replay
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
            json.dump(self.mutation_logger.mutations, f)
            temp_path = f.name

        try:
            # Replay mutations to rebuild overlay_state
            # If this fails, overlay_state remains in its previous state
            new_state = replay_mutations(self.ir_db_path, temp_path)
            # Only update overlay_state if replay succeeded
            self.overlay_state = new_state
            self._last_synced_mutation_count = current_count
        except Exception as e:
            # Log error but preserve current overlay_state
            logger.error(f"Failed to sync overlay_state: {e}")
            logger.warning("Continuing with previous overlay_state (may be stale)")
            raise
        finally:
            # Clean up temporary file
            Path(temp_path).unlink(missing_ok=True)
            self._log_timing(
                "sync_overlay_state",
                time.perf_counter() - start_time,
                f"mutations={len(self.mutation_logger.mutations)}",
            )

    def run_labelling(
        self,
        output_mutations_path: str,
        output_overlay_path: str,
        enable_sheet_refinement: bool = True,
        run_confidence_assessment: bool = False,
    ) -> int:
        """
        Run full two-pass labelling workflow.

        Args:
            output_mutations_path: Where to write mutations.json
            output_overlay_path: Where to write semantic_overlay.db
            enable_sheet_refinement: Whether to run sheet-level refinement (default: True)
            run_confidence_assessment: Whether to run Pass 4 confidence assessment (default: False)
                NOTE: Pass 4 should run AFTER classification to assess classification_confidence.
                Set to False during labelling, then call run_pass_4_confidence_assessment()
                after apply_actuarial_classification() completes.

        Returns:
            Number of bindings labelled
        """
        start_time = time.perf_counter()
        logger.info("Starting two-pass labelling")

        # Load data
        self.load_bindings()
        self.build_graph()

        # Run Pass 1 (top-down) - includes sync
        self.run_pass_1_top_down()

        # Run Pass 2 (bottom-up) - includes sync
        self.run_pass_2_bottom_up()

        # Disambiguate duplicate labels (Story 20)
        disambiguate_start = time.perf_counter()
        self._disambiguate_duplicate_labels()
        self._log_timing("disambiguate_duplicates", time.perf_counter() - disambiguate_start)

        # Run Pass 3 (label generation for low-confidence bindings) - includes sync
        pass3_start = time.perf_counter()
        self.run_pass_3_label_generation()
        self._log_timing("pass_3_label_generation", time.perf_counter() - pass3_start)

        # Save mutations
        self.mutation_logger.save(output_mutations_path)
        logger.info(f"Saved mutations to {output_mutations_path}")

        # Use already-synced overlay_state (no need to replay again)
        # Pass 3 already called _sync_overlay_state(), so overlay_state is current
        overlay = self.overlay_state

        # Log fallback label metrics
        self._log_fallback_metrics()

        # Write overlay database (architectural boundary enforced by runtime check)
        overlay_start = time.perf_counter()
        write_overlay_to_db(overlay, output_mutations_path, self.ir_db_path, output_overlay_path)
        logger.info(f"Saved overlay to {output_overlay_path}")
        self._log_timing("write_overlay_db", time.perf_counter() - overlay_start)

        # Run Pass 4 (confidence assessment and review report generation) if enabled
        # CRITICAL: Pass 4 should run AFTER classification to assess classification_confidence
        # Set run_confidence_assessment=False during labelling, then call this method
        # after apply_actuarial_classification() completes
        if run_confidence_assessment:
            logger.info("Running Pass 4 (confidence assessment)")
            self.run_pass_4_confidence_assessment(output_overlay_path, output_mutations_path)
        else:
            logger.info("Skipping Pass 4 (will run after classification)")

        # Run sheet-level refinement (Post-Pass)
        if enable_sheet_refinement:
            sheet_start = time.perf_counter()
            logger.info("Running sheet-level refinement (Post-Pass)")
            refiner = SheetContextRefiner(self.ir_db_path)
            refinement_count = refiner.refine(output_overlay_path, output_mutations_path)
            logger.info(
                f"Sheet-level refinement: {refinement_count} potential refinements identified"
            )
            self._log_timing("sheet_refinement", time.perf_counter() - sheet_start)

        # Return count of labelled bindings from final overlay_state
        labelled_count = sum(1 for b in overlay.bindings.values() if b.label is not None)
        self._log_timing("total_labelling", time.perf_counter() - start_time)
        return labelled_count

    def _log_fallback_metrics(self) -> None:
        """
        Log metrics about fallback label usage.

        Analyzes mutations to count:
        - Total fallback labels generated
        - Binding IDs rejected (debug_label contained ::)
        - Meaningful debug_labels used
        - A1 address fallbacks
        """
        fallback_count = 0
        binding_id_rejected_count = 0
        meaningful_debug_label_count = 0
        a1_fallback_count = 0
        total_labels = 0

        # Analyze set_label mutations
        for mutation in self.mutation_logger.mutations:
            if mutation.get("action") == "set_label":
                total_labels += 1
                knowledge_source = mutation.get("knowledge_source", "")

                if knowledge_source == "fallback":
                    fallback_count += 1

                    # Try to determine fallback sub-type from the label itself
                    # This is approximate since we don't have access to the original binding here
                    label = mutation.get("new_label", "")

                    # If label contains ::, it was a binding ID that got rejected
                    # (shouldn't happen after Story 23 fix, but track for monitoring)
                    if "::" in label:
                        binding_id_rejected_count += 1
                    # If label contains !, it's likely an A1 fallback (Sheet!Address)
                    elif "!" in label:
                        a1_fallback_count += 1
                    # Otherwise, assume it's a meaningful debug_label
                    else:
                        meaningful_debug_label_count += 1

        # Log summary
        if fallback_count > 0:
            logger.info("Fallback label metrics:")
            logger.info(
                f"  Total fallback labels: {fallback_count}/{total_labels} ({100 * fallback_count / total_labels:.1f}%)"
            )
            logger.info(f"  Binding IDs in labels (should be 0): {binding_id_rejected_count}")
            logger.info(f"  A1 address fallbacks: {a1_fallback_count}")
            logger.info(f"  Meaningful debug_labels: {meaningful_debug_label_count}")
        else:
            logger.info("No fallback labels generated - all bindings had label candidates")

    def _disambiguate_duplicate_labels(self) -> None:
        """
        Detect and fix duplicate labels after Pass 2.

        When multiple bindings share the same label, append a disambiguating suffix based on:
        - Pattern 1 (Header vs Data): If one is a header row and one is data, append [Header]
        - Pattern 2 (Source vs Lookup): If different sheets, append sheet name in parentheses
        - Pattern 3 (Generic): Append location suffix (Row N or Column X)

        Story 22 Fix: If duplicates remain after initial disambiguation (e.g., both have [Header]),
        apply additional location-based disambiguation as a fallback.

        This ensures all labels are unique and meaningful for actuaries.
        """
        from collections import defaultdict

        logger.info("Disambiguating duplicate labels")

        # 1. Group bindings by label
        label_groups: dict[str, list[str]] = defaultdict(list)
        for binding_id, binding_overlay in self.overlay_state.bindings.items():
            if binding_overlay.label:
                label_groups[binding_overlay.label].append(binding_id)

        duplicate_ids = [
            binding_id
            for binding_ids in label_groups.values()
            if len(binding_ids) > 1
            for binding_id in binding_ids
        ]

        if not duplicate_ids:
            logger.info("No duplicate labels found")
            return

        # Pre-load binding shapes and addresses for duplicates only (one connection for efficiency)
        conn = sqlite3.connect(self.ir_db_path)
        conn.execute("DROP TABLE IF EXISTS temp_duplicate_bindings")
        conn.execute("""
            CREATE TEMP TABLE temp_duplicate_bindings (
                binding_id TEXT PRIMARY KEY
            )
        """)
        batch = []
        for binding_id in duplicate_ids:
            batch.append((binding_id,))
            if len(batch) >= 10000:
                conn.executemany(
                    "INSERT OR IGNORE INTO temp_duplicate_bindings (binding_id) VALUES (?)", batch
                )
                batch = []
        if batch:
            conn.executemany(
                "INSERT OR IGNORE INTO temp_duplicate_bindings (binding_id) VALUES (?)", batch
            )

        shape_rows = conn.execute("""
            SELECT b.binding_id, b.shape_rows, b.shape_cols
            FROM agent_bindings b
            JOIN temp_duplicate_bindings t ON t.binding_id = b.binding_id
        """).fetchall()
        address_rows = conn.execute("""
            SELECT b.binding_id, b.address
            FROM agent_bindings b
            JOIN temp_duplicate_bindings t ON t.binding_id = b.binding_id
        """).fetchall()
        conn.execute("DROP TABLE IF EXISTS temp_duplicate_bindings")
        conn.close()

        shape_cache: dict[str, tuple[int, int]] = {row[0]: (row[1], row[2]) for row in shape_rows}
        address_cache: dict[str, str] = {row[0]: row[1] for row in address_rows}

        # 2. For each duplicate group, apply disambiguation
        disambiguation_count = 0
        label_map: dict[str, str] = {
            binding_id: binding_overlay.label
            for binding_id, binding_overlay in self.overlay_state.bindings.items()
            if binding_overlay.label
        }
        location_suffix_cache: dict[str, str] = {}

        def _get_cached_location_suffix(binding_id: str) -> str:
            if binding_id in location_suffix_cache:
                return location_suffix_cache[binding_id]
            suffix = self._get_location_suffix(binding_id, shape_cache, address_cache)
            location_suffix_cache[binding_id] = suffix
            return suffix

        for label, binding_ids in label_groups.items():
            if len(binding_ids) <= 1:
                continue  # Not a duplicate

            group_has_data_rows = any(
                shape_cache.get(binding_id, (0, 0))[0] > 10 for binding_id in binding_ids
            )
            group_sheets = {
                self.bindings[binding_id].sheet
                for binding_id in binding_ids
                if binding_id in self.bindings
            }
            group_has_multiple_sheets = len(group_sheets) > 1

            # Determine disambiguation strategy for each binding
            for binding_id in binding_ids:
                if binding_id not in self.bindings or binding_id not in shape_cache:
                    continue
                shape_rows, _ = shape_cache[binding_id]
                if shape_rows == 1 and group_has_data_rows:
                    label_text = label.strip().strip('"')
                    is_generic_label = False
                    if label_text and is_numeric(label_text):
                        is_generic_label = True
                    binding_sheet = self.bindings[binding_id].sheet
                    if binding_sheet and label_text.lower() == binding_sheet.lower():
                        is_generic_label = True
                    if is_generic_label:
                        suffix = _get_cached_location_suffix(binding_id)
                    else:
                        suffix = "[Header]"
                elif group_has_multiple_sheets:
                    suffix = f"({self.bindings[binding_id].sheet})"
                else:
                    suffix = _get_cached_location_suffix(binding_id)
                if suffix:
                    new_label = f"{label} {suffix}"
                    self.mutation_logger.set_label(
                        binding_id=binding_id,
                        old_label=label,
                        new_label=new_label,
                        reasoning=f"Disambiguate duplicate label '{label}'",
                        knowledge_source="disambiguation",
                        confidence=1.0,  # High confidence - deterministic disambiguation
                    )
                    disambiguation_count += 1
                    label_map[binding_id] = new_label

        # 3. Story 22: Check for remaining duplicates (e.g., both have [Header] suffix)
        # Re-group bindings by label after first pass
        label_groups_after = defaultdict(list)
        for binding_id, label in label_map.items():
            if label:
                label_groups_after[label].append(binding_id)

        # Find labels that are still duplicated
        still_duplicated = {label for label, ids in label_groups_after.items() if len(ids) > 1}

        fallback_count = 0
        if still_duplicated:
            logger.info(
                f"Found {len(still_duplicated)} labels still duplicated after first pass - applying fallback disambiguation"
            )

            for label in still_duplicated:
                duplicates = label_groups_after[label]
                # The first-pass suffix (often a shared "(Row N)") failed to
                # separate this group. Strip that redundant generic suffix and
                # re-suffix each sibling with its distinguishing top-left cell
                # coordinate (e.g. "(C55)"/"(D55)"), which separates siblings
                # that share a row or a column.
                base = _LOC_SUFFIX_RE.sub("", label)
                for binding_id in duplicates:
                    suffix = self._get_cell_coord_suffix(binding_id) or _get_cached_location_suffix(
                        binding_id
                    )
                    if not suffix:
                        continue
                    if base.rstrip().endswith(suffix):
                        # Already carries this exact coordinate; avoid doubling.
                        new_label = base
                    else:
                        new_label = f"{base} {suffix}"
                    if new_label == label_map.get(binding_id):
                        continue
                    self.mutation_logger.set_label(
                        binding_id=binding_id,
                        old_label=label,
                        new_label=new_label,
                        reasoning="Fallback disambiguation via distinguishing cell coordinate",
                        knowledge_source="disambiguation",
                        confidence=1.0,
                    )
                    fallback_count += 1
                    label_map[binding_id] = new_label

            if fallback_count > 0:
                logger.info(f"Applied {fallback_count} fallback disambiguations")

        if disambiguation_count > 0 or fallback_count > 0:
            self._sync_overlay_state()
            logger.info(f"Disambiguated {disambiguation_count} duplicate labels (first pass)")

        if disambiguation_count == 0 and not still_duplicated:
            logger.info("No duplicate labels found")

    def _get_disambiguation_suffix(
        self,
        binding_id: str,
        label: str,
        duplicates: list[str],
        shape_cache: dict[str, tuple[int, int]],
        address_cache: dict[str, str],
    ) -> str:
        """
        Determine disambiguation suffix based on patterns.

        Args:
            binding_id: Binding to disambiguate
            label: Current label
            duplicates: List of all binding_ids with this label
            shape_cache: Pre-loaded shape data (binding_id -> (rows, cols))

        Returns:
            Suffix to append (empty string if no disambiguation needed)
        """
        if binding_id not in self.bindings:
            return ""

        binding = self.bindings[binding_id]

        # Get binding shape from cache
        if binding_id not in shape_cache:
            return ""

        shape_rows, shape_cols = shape_cache[binding_id]

        # Pattern 1: Header vs Data (check if one is a single-row header)
        # A binding with 1 row and another with >10 rows suggests header vs data
        if shape_rows == 1:
            # Check if any duplicate has many rows (likely data)
            for dup_id in duplicates:
                if dup_id == binding_id or dup_id not in shape_cache:
                    continue
                dup_rows, _ = shape_cache[dup_id]
                if dup_rows > 10:
                    return "[Header]"

        # Pattern 2: Source vs Lookup (check sheet names)
        sheets = set()
        for dup_id in duplicates:
            if dup_id in self.bindings:
                sheets.add(self.bindings[dup_id].sheet)

        if len(sheets) > 1:
            # Multiple sheets - append sheet name
            return f"({binding.sheet})"

        # Pattern 3: Generic (use location)
        return self._get_location_suffix(binding_id, shape_cache, address_cache)

    def _get_location_suffix(
        self,
        binding_id: str,
        shape_cache: dict[str, tuple[int, int]],
        address_cache: dict[str, str] | None = None,
    ) -> str:
        """
        Generate location-based suffix for disambiguation (Story 22).

        Used as fallback when other disambiguation patterns don't apply
        (e.g., both duplicates already have [Header] suffix).

        Args:
            binding_id: Binding to generate suffix for
            shape_cache: Pre-loaded shape data (binding_id -> (rows, cols))

        Returns:
            Location suffix like "(Row 14)" or "(Column X)" or "(Sheet!Address)"
        """
        address_a1 = ""
        if address_cache is not None:
            address_a1 = address_cache.get(binding_id, "")
        else:
            # Query IR database for address (binding might not be in self.bindings after filtering)
            conn = sqlite3.connect(self.ir_db_path)
            row = conn.execute(
                """
                SELECT address FROM agent_bindings WHERE binding_id = ?
            """,
                (binding_id,),
            ).fetchone()
            conn.close()

            if not row:
                return ""

            address_a1 = row[0]

        if not address_a1:
            return ""

        # Get binding shape from cache
        if binding_id not in shape_cache:
            return ""

        shape_rows, shape_cols = shape_cache[binding_id]

        # Extract row/column from address
        # address_a1 format: "Sheet!A1:B10" or "Sheet!A1"
        address = address_a1
        if "!" in address:
            address = address.split("!")[-1]

        # Remove $ signs
        address = address.replace("$", "")

        # Determine if primarily row-oriented or column-oriented
        if shape_rows > shape_cols:
            # Column-oriented - extract column letter
            # Extract first column letter(s)
            col_match = re.match(r"^([A-Z]+)", address)
            if col_match:
                return f"(Column {col_match.group(1)})"
        else:
            # Row-oriented - extract row number
            # Extract first row number
            row_match = re.search(r"(\d+)", address)
            if row_match:
                return f"(Row {row_match.group(1)})"

        # Fallback: use full address
        return f"({address})"

    def _get_cell_coord_suffix(self, binding_id: str) -> str:
        """Return a "(<col><row>)" suffix from the binding's top-left cell.

        Unlike the row/column-only location suffix, the full cell coordinate
        (e.g. "(C55)") distinguishes siblings that share a row (or a column) —
        the exact case the first-pass disambiguation suffix fails on.
        """
        address = ""
        binding = self.bindings.get(binding_id)
        if binding is not None:
            address = binding.address_a1 or ""
        if not address:
            conn = sqlite3.connect(self.ir_db_path)
            row = conn.execute(
                "SELECT address FROM agent_bindings WHERE binding_id = ?",
                (binding_id,),
            ).fetchone()
            conn.close()
            if not row:
                return ""
            address = row[0] or ""
        if not address:
            return ""
        if "!" in address:
            address = address.split("!")[-1]
        address = address.replace("$", "")
        top_left = address.split(":")[0].strip()
        if not re.match(r"^[A-Z]+\d+$", top_left):
            return ""
        return f"({top_left})"

    def run_pass_3_label_generation(self) -> None:
        """
        Pass 3: Generate labels for bindings with low confidence or generic labels.

        This pass attempts to generate meaningful labels from:
        - Formula patterns (SUM, AVERAGE, NPV, etc.)
        - Dependency labels
        - Actuarial context

        Only applies to Result and Calculation variables with poor labels.
        After processing, syncs overlay_state by replaying all mutations.
        """
        start_time = time.perf_counter()
        logger.info("Pass 3: Label generation for low-confidence bindings")

        # We need formula and dependency information from IR
        import sqlite3

        conn = sqlite3.connect(self.ir_db_path)

        generation_count = 0
        label_dict = {
            bid: (b.label, 1.0, "stateless")
            for bid, b in self.overlay_state.bindings.items()
            if b.label
        }

        labeled_binding_ids = [
            binding_id
            for binding_id, binding_overlay in self.overlay_state.bindings.items()
            if binding_overlay.label and binding_id in self.bindings
        ]
        formula_map: dict[str, str] = {}
        dependency_map: dict[str, list[str]] = {}
        if labeled_binding_ids:
            conn.execute("DROP TABLE IF EXISTS temp_pass3_bindings")
            conn.execute("""
                CREATE TEMP TABLE temp_pass3_bindings (
                    binding_id TEXT PRIMARY KEY
                )
            """)
            batch = []
            for binding_id in labeled_binding_ids:
                batch.append((binding_id,))
                if len(batch) >= 10000:
                    conn.executemany(
                        "INSERT OR IGNORE INTO temp_pass3_bindings (binding_id) VALUES (?)", batch
                    )
                    batch = []
            if batch:
                conn.executemany(
                    "INSERT OR IGNORE INTO temp_pass3_bindings (binding_id) VALUES (?)", batch
                )

            formula_rows = conn.execute("""
                WITH first_formula AS (
                    SELECT ctb.binding_id, MIN(ac.cell_address) AS min_addr
                    FROM agent_cells ac
                    JOIN cell_to_binding ctb ON ac.cell_id = ctb.cell_id
                    JOIN temp_pass3_bindings t ON t.binding_id = ctb.binding_id
                    WHERE ac.formula IS NOT NULL AND ac.formula != ''
                    GROUP BY ctb.binding_id
                )
                SELECT ctb.binding_id, ac.formula
                FROM agent_cells ac
                JOIN cell_to_binding ctb ON ac.cell_id = ctb.cell_id
                JOIN first_formula f
                  ON f.binding_id = ctb.binding_id
                 AND f.min_addr = ac.cell_address
            """).fetchall()
            formula_map = {row[0]: row[1] for row in formula_rows if row[1]}

        candidate_ids = []
        for binding_id in labeled_binding_ids:
            binding_overlay = self.overlay_state.bindings.get(binding_id)
            if not binding_overlay or not binding_overlay.label:
                continue

            # We don't have score stored - use heuristic from binding candidates
            _, current_score, candidate_type = simple_label_selection(self.bindings[binding_id])

            formula = formula_map.get(binding_id)
            if not formula:
                continue

            # Check if we should generate a label
            # Note: We don't have actuarial_class yet (it's applied after labelling)
            # So we use heuristics: Result-like patterns are bindings with formula but no children
            if self._child_cache is not None:
                children = self._child_cache.get(binding_id, [])
            else:
                children = get_children(self.graph, binding_id) if self.graph else []
            is_likely_result = formula and not children
            is_likely_calculation = formula and children

            pseudo_class = (
                "Result" if is_likely_result else ("Calculation" if is_likely_calculation else None)
            )

            if should_generate_label(
                binding_overlay.label, current_score, pseudo_class, candidate_type
            ):
                candidate_ids.append(binding_id)

        if candidate_ids:
            conn.execute("DROP TABLE IF EXISTS temp_pass3_candidates")
            conn.execute("""
                CREATE TEMP TABLE temp_pass3_candidates (
                    binding_id TEXT PRIMARY KEY
                )
            """)
            batch = []
            for binding_id in candidate_ids:
                batch.append((binding_id,))
                if len(batch) >= 10000:
                    conn.executemany(
                        "INSERT OR IGNORE INTO temp_pass3_candidates (binding_id) VALUES (?)", batch
                    )
                    batch = []
            if batch:
                conn.executemany(
                    "INSERT OR IGNORE INTO temp_pass3_candidates (binding_id) VALUES (?)", batch
                )

            dep_rows = conn.execute("""
                SELECT d.from_binding, d.to_binding
                FROM agent_binding_dependencies d
                JOIN temp_pass3_candidates t ON t.binding_id = d.from_binding
            """).fetchall()
            for from_binding, to_binding in dep_rows:
                dependency_map.setdefault(from_binding, []).append(to_binding)

            conn.execute("DROP TABLE IF EXISTS temp_pass3_candidates")

        conn.execute("DROP TABLE IF EXISTS temp_pass3_bindings")

        # Iterate over bindings in overlay_state (not a dict anymore)
        for binding_id in candidate_ids:
            binding_overlay = self.overlay_state.bindings.get(binding_id)
            if not binding_overlay or not binding_overlay.label:
                continue

            current_label = binding_overlay.label
            _, current_score, candidate_type = simple_label_selection(self.bindings[binding_id])
            formula = formula_map.get(binding_id)
            if not formula:
                continue

            # Get dependencies (bindings this one depends on)
            # Edge from_binding -> to_binding means "from_binding depends on to_binding"
            # So we query WHERE from_binding = binding_id
            parent_bindings = dependency_map.get(binding_id, [])

            # Attempt to generate a better label
            # Pass overlay_state.bindings as a dict-like structure for label lookup
            # Fix: label_generator expects (label, score, type) tuple, but BindingOverlay only has label
            # We assume labels in OverlayState are high confidence (filtered during admission)
            generated_label = generate_label_from_formula(
                formula, parent_bindings, self.ir_db_path, label_dict
            )

            if generated_label:
                old_label = current_label

                # Log mutation (write path)
                self.mutation_logger.set_label(
                    binding_id=binding_id,
                    old_label=old_label,
                    new_label=generated_label,
                    reasoning=(
                        f"Pass 3: Generated from formula pattern. "
                        f"Original: '{old_label}' (score: {current_score:.2f}, "
                        f"type: {candidate_type})"
                    ),
                    knowledge_source="formula_generation",
                    confidence=0.85,
                )

                generation_count += 1
                logger.debug(f"Pass 3: {binding_id} -> {generated_label} (generated from formula)")

        conn.close()

        # CRITICAL: Sync overlay_state by replaying mutations
        self._sync_overlay_state()

        logger.info(f"Pass 3 complete: Generated {generation_count} labels")
        self._log_timing(
            "pass_3_generation_internal",
            time.perf_counter() - start_time,
            f"generated={generation_count}",
        )

    def run_pass_4_confidence_assessment(
        self,
        overlay_db_path: str,
        mutations_path: str,
        confidence_threshold: float = 0.7,
        output_report_dir: str | None = None,
    ) -> None:
        """
        Pass 4: Confidence assessment and review report generation.

        Assesses confidence for all labeling and classification decisions.
        Updates confidence scores in overlay database.
        Generates Markdown review report for Cursor agent to evaluate low-confidence cases.

        Args:
            overlay_db_path: Path to semantic overlay database
            mutations_path: Path to mutations.json (for audit trail)
            confidence_threshold: Threshold below which to flag for review (default: 0.7)
            output_report_dir: Directory to save review report (default: same dir as overlay)
        """
        from pathlib import Path

        from xl_marinade.docs.confidence_scorer import assess_all_confidence
        from xl_marinade.docs.low_confidence_reporter import (
            LowConfidenceCase,
            filter_low_confidence_cases,
            generate_review_report,
        )

        logger.info(f"Pass 4: Confidence assessment (threshold={confidence_threshold})")

        # Assess confidence for all bindings
        all_confidence = assess_all_confidence(
            self.ir_db_path,
            overlay_db_path,
        )

        logger.info(f"Assessed confidence for {len(all_confidence)} bindings")

        # Filter low-confidence cases
        low_conf_ids = filter_low_confidence_cases(all_confidence, confidence_threshold)

        if not low_conf_ids:
            logger.info("No low-confidence bindings found - skipping review report")
            return

        # Build detailed cases for report
        import sqlite3

        ir_conn = sqlite3.connect(self.ir_db_path)
        overlay_conn = sqlite3.connect(overlay_db_path)

        cases = []
        for binding_id in low_conf_ids:
            # Find confidence scores for this binding
            conf_scores = next((c for c in all_confidence if c.binding_id == binding_id), None)
            if not conf_scores:
                continue

            # Get binding details from IR
            binding_row = ir_conn.execute(
                """
                SELECT sheet, address, spatial_candidates
                FROM agent_bindings
                WHERE binding_id = ?
            """,
                (binding_id,),
            ).fetchone()

            if not binding_row:
                continue

            sheet, address, candidates_json = binding_row

            # Get formula from binding (fast schema uses formula_pattern from agent_bindings)
            # For fast schema, we need to get formula via cell_to_binding join
            formula_row = ir_conn.execute(
                """
                SELECT ac.formula
                FROM cell_to_binding ctb
                JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
                WHERE ctb.binding_id = ? AND ac.formula IS NOT NULL
                LIMIT 1
            """,
                (binding_id,),
            ).fetchone()

            formula = formula_row[0] if formula_row else None

            # Get current label and classification from overlay
            overlay_row = overlay_conn.execute(
                """
                SELECT label, actuarial_class
                FROM semantic_variables
                WHERE binding_id = ?
            """,
                (binding_id,),
            ).fetchone()

            current_label, current_class = overlay_row if overlay_row else (None, None)

            # Parse candidates JSON
            import json

            try:
                if candidates_json:
                    candidates_data = json.loads(candidates_json)
                    # Handle case where JSON is a dict with 'label_candidates' key
                    if isinstance(candidates_data, dict):
                        candidates = candidates_data.get("label_candidates", [])
                    elif isinstance(candidates_data, list):
                        # Backwards compatibility: if it's already a list
                        candidates = candidates_data
                    else:
                        candidates = []
                else:
                    candidates = []
            except (json.JSONDecodeError, TypeError, AttributeError) as e:
                logger.warning(f"Failed to parse candidates JSON for {binding_id}: {e}")
                candidates = []

            # Get parent/child labels (placeholder - requires graph traversal)
            parent_labels = []
            child_labels = []

            # RAG matches removed (NEW 2): empty placeholder
            rag_matches = []

            case = LowConfidenceCase(
                binding_id=binding_id,
                sheet=sheet,
                address=address,
                formula=formula,
                current_label=current_label,
                current_classification=current_class,
                label_confidence=conf_scores.label_confidence,
                classification_confidence=conf_scores.classification_confidence,
                label_components=conf_scores.label_components,
                classification_components=conf_scores.classification_components,
                candidates=candidates,
                parent_labels=parent_labels,
                child_labels=child_labels,
                rag_matches=rag_matches,
            )

            cases.append(case)

        ir_conn.close()

        # Update confidence scores in overlay database
        logger.info("Updating confidence scores in overlay database")
        for conf_scores in all_confidence:
            overlay_conn.execute(
                """
                UPDATE semantic_variables
                SET label_confidence = ?,
                    classification_confidence = ?
                WHERE binding_id = ?
            """,
                (
                    conf_scores.label_confidence,
                    conf_scores.classification_confidence,
                    conf_scores.binding_id,
                ),
            )

        overlay_conn.commit()
        overlay_conn.close()

        logger.info(f"Updated confidence scores for {len(all_confidence)} bindings")

        # Generate review report
        if output_report_dir is None:
            output_report_dir = str(Path(overlay_db_path).parent)

        report_path = str(Path(output_report_dir) / "low_confidence_review.md")
        generate_review_report(cases, report_path, confidence_threshold)

        logger.info(
            f"Pass 4 complete: Generated review report with {len(cases)} cases at {report_path}"
        )
