# ABOUTME: Deterministic time-dependent binding annotation for Sprint 6.
# ABOUTME: Annotates bindings based on spatial alignment and graph dependency relative to time index.
from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from pathlib import Path

from xl_marinade.core.ref_converter import parse_cell_address

_MIN_ALIGNMENT_OVERLAP = 3
_MAX_SEPARATOR_GAP = 1
_MIN_LENGTH_SIMILARITY = 0.8

# Single source of truth for the time-dependence confidence ladder. A binding is
# annotated is_time_dependent with one of four tiers relative to the time index:
#   0.95 — aligned AND graph-dependent (spatial + calc-graph evidence)
#   0.70 — graph-dependent but not spatially aligned  (TIME_GRAPH_CONFIDENCE_MIN)
#   0.50 — spatially aligned only, no graph dependency (a per-item value that merely
#          sits beside a long integer column the heuristic mistook for a time index)
#   0.00 — neither aligned nor dependent
# TIME_GRAPH_CONFIDENCE_MIN is BOTH the graph-only tier value here AND the floor the
# semantic builder uses to keep an axis as axis_kind='time' (builder.py
# _extract_axes_from_tables): the two are the same number *by design* — a series is
# emitted as a time axis iff its dependence is graph-backed (>= this floor), which
# drops the 0.50 spatial-only false positives. Coupling the tier value and the filter
# threshold to one constant keeps them from drifting apart (they have silently
# diverged before). Change here => the builder filter moves in lockstep.
TIME_GRAPH_CONFIDENCE_MIN = 0.7


@dataclass(frozen=True)
class TimeIndex:
    sheet_key: str
    binding_id: str
    axis: str  # "row" or "col" (orientation of the vector)
    top_left_row: int
    top_left_col: int
    shape_rows: int
    shape_cols: int


@dataclass(frozen=True)
class BindingInfo:
    binding_id: str
    sheet_key: str
    top_left_row: int
    top_left_col: int
    shape_rows: int
    shape_cols: int


@dataclass(frozen=True)
class TimeDependenceAnnotation:
    binding_id: str
    time_index_binding_id: str
    is_time_dependent: bool
    confidence: float
    reasons_top3: list[str]
    evidence_flags: dict[str, bool]


def infer_time_dependence(db_path: Path) -> list[dict[str, object]]:
    """
    Infer time dependence annotations from an existing ir.db file.

    Returns:
        List of dicts matching binding_time_annotations schema.
    """
    conn = sqlite3.connect(str(db_path))
    try:
        return infer_time_dependence_from_conn(conn)
    finally:
        conn.close()


def infer_time_dependence_from_conn(conn: sqlite3.Connection) -> list[dict[str, object]]:
    """
    Infer time dependence annotations from an open SQLite connection.

    This is required for fast-pipeline post-processing where the build connection
    may use exclusive locking (a second connection would fail with 'database is locked').
    """
    time_indices = _load_primary_time_indices(conn)
    if not time_indices:
        return []

    bindings = _load_bindings(conn)
    dependencies = _load_binding_dependencies(conn)

    annotations = []
    for sheet_key, time_index in time_indices.items():
        sheet_bindings = [
            b
            for b in bindings
            if b.sheet_key == sheet_key and b.binding_id != time_index.binding_id
        ]
        sheet_annotations = _annotate_sheet_bindings(time_index, sheet_bindings, dependencies)
        annotations.extend(sheet_annotations)

    return [
        {
            "binding_id": a.binding_id,
            "time_index_binding_id": a.time_index_binding_id,
            "is_time_dependent": a.is_time_dependent,
            "confidence": a.confidence,
            "reasons_top3_json": json.dumps(a.reasons_top3),
            "evidence_flags_json": json.dumps(a.evidence_flags),
        }
        for a in annotations
    ]


def persist_binding_time_annotations(db_path: Path, annotations: list[dict[str, object]]) -> None:
    """Persist binding time annotations into ir.db."""
    from xl_marinade.core import schema as ir_schema

    conn = ir_schema.open_existing_database(db_path)
    try:
        ir_schema.insert_binding_time_annotations(conn, annotations)
    finally:
        conn.close()


def _load_primary_time_indices(conn: sqlite3.Connection) -> dict[str, TimeIndex]:
    """Load rank-1 time index candidates per sheet."""
    cursor = conn.cursor()
    # Join with bindings to get shape/location
    if _table_has_column(conn, "time_index_candidates", "sheet"):
        cursor.execute(
            """
            SELECT t.sheet, t.binding_id, b.address_a1, b.shape_rows, b.shape_cols
            FROM time_index_candidates t
            JOIN bindings b ON t.binding_id = b.binding_id
            WHERE t.rank = 1
            """
        )
    else:
        cursor.execute(
            """
            SELECT t.sheet_id, t.binding_id, b.address_a1, b.shape_rows, b.shape_cols
            FROM time_index_candidates t
            JOIN bindings b ON t.binding_id = b.binding_id
            WHERE t.rank = 1
            """
        )
    indices = {}
    for row in cursor.fetchall():
        sheet_value, binding_id, address, shape_rows, shape_cols = row
        parsed = parse_cell_address(address)

        # Determine axis (orientation of the vector)
        # If shape_rows=1, it's a row vector (extends horizontally), so axis is "row"
        # If shape_cols=1, it's a col vector (extends vertically), so axis is "col"
        if shape_rows == 1:
            axis = "row"
        elif shape_cols == 1:
            axis = "col"
        else:
            # Fallback for non-1D time indices (unlikely given inference logic, but safe default)
            axis = "unknown"

        indices[str(sheet_value)] = TimeIndex(
            sheet_key=str(sheet_value),
            binding_id=binding_id,
            axis=axis,
            top_left_row=int(parsed["row"]),
            top_left_col=int(parsed["col"]),
            shape_rows=shape_rows,
            shape_cols=shape_cols,
        )
    return indices


def _load_bindings(conn: sqlite3.Connection) -> list[BindingInfo]:
    cursor = conn.cursor()
    address_column = _bindings_address_column(conn)
    if _table_has_column(conn, "bindings", "sheet_id"):
        cursor.execute(
            f"""
            SELECT binding_id, sheet_id, {address_column}, shape_rows, shape_cols
            FROM bindings
            """
        )
    elif _table_has_column(conn, "bindings", "sheet"):
        cursor.execute(
            f"""
            SELECT binding_id, sheet, {address_column}, shape_rows, shape_cols
            FROM bindings
            """
        )
    else:
        raise RuntimeError("bindings table missing sheet or sheet_id column")
    bindings = []
    for row in cursor.fetchall():
        binding_id, sheet_value, address, shape_rows, shape_cols = row
        parsed = parse_cell_address(address)
        bindings.append(
            BindingInfo(
                binding_id=binding_id,
                sheet_key=str(sheet_value),
                top_left_row=int(parsed["row"]),
                top_left_col=int(parsed["col"]),
                shape_rows=shape_rows,
                shape_cols=shape_cols,
            )
        )
    return bindings


def _load_binding_dependencies(conn: sqlite3.Connection) -> set[tuple[str, str]]:
    """Load direct dependencies between bindings."""
    cursor = conn.cursor()
    if _table_exists(conn, "binding_edges"):
        cursor.execute("SELECT from_binding_id, to_binding_id FROM binding_edges")
    elif _table_exists(conn, "binding_level_edges"):
        cursor.execute("SELECT from_binding_id, to_binding_id FROM binding_level_edges")
    else:
        raise RuntimeError("No binding edge table found (binding_edges or binding_level_edges)")
    return {(row[0], row[1]) for row in cursor.fetchall()}


def _annotate_sheet_bindings(
    time_index: TimeIndex,
    bindings: list[BindingInfo],
    dependencies: set[tuple[str, str]],
) -> list[TimeDependenceAnnotation]:
    annotations = []

    # Build dependency graph for this sheet (or global? dependencies are global IDs)
    # We need to check if binding depends on time_index.
    # Simple BFS for reachability? Or just direct?
    # Design says: "Graph dependency evidence (direct/indirect dependency on time index)."
    # For v1, let's stick to direct or 1-hop? Or full reachability?
    # Full reachability might be expensive if graph is huge.
    # But usually we care if *binding* depends on *time_index*.
    # i.e. binding -> ... -> time_index.
    # Wait, usually formulas refer TO the time index.
    # So `binding` depends on `time_index`.
    # Edge direction: `from_binding` depends on `to_binding`.
    # So we look for path `binding` -> `time_index`.

    # Let's pre-calculate reachability to time_index for all bindings on this sheet?
    # Or just check direct for now and expand if needed.
    # Given "indirect", we should probably do a traversal.
    # Since we process per sheet, and time index is per sheet, we can filter relevant edges?
    # No, dependencies can cross sheets.

    # Optimization: Reverse graph from time_index?
    # Find all bindings that depend on time_index.
    # i.e. traverse backwards from time_index in the `binding_edges` (where edge is from->to).
    # Wait, `binding_edges` is `from` depends on `to`.
    # So if A depends on Time, there is edge A->Time.
    # So we want to find all A such that A -> ... -> Time.
    # So we traverse the graph backwards from Time along incoming edges?
    # Yes. `SELECT from_binding_id FROM binding_edges WHERE to_binding_id = ?`

    # Since we loaded all dependencies into memory (might be large?), let's optimize.
    # If `dependencies` is set of (from, to), we want ancestors of time_index.
    # Let's build a reverse map: to -> list[from]
    reverse_deps = {}
    for u, v in dependencies:
        reverse_deps.setdefault(v, []).append(u)

    dependent_on_time = _find_ancestors(time_index.binding_id, reverse_deps)

    for binding in bindings:
        is_aligned = _check_alignment(time_index, binding)
        depends_on_time = binding.binding_id in dependent_on_time

        reasons = []
        confidence = 0.0
        is_time_dependent = False
        evidence_flags = {
            "aligned": is_aligned,
            "depends_on_time_index": depends_on_time,
            "alignment_only_penalty_applied": False,
        }

        if is_aligned and depends_on_time:
            is_time_dependent = True
            confidence = 0.95
            reasons.append("Aligned with time index and depends on it")
        elif is_aligned and not depends_on_time:
            is_time_dependent = True
            confidence = 0.5  # Penalty
            evidence_flags["alignment_only_penalty_applied"] = True
            reasons.append("Aligned with time index (spatial only)")
            reasons.append("No graph dependency found (penalty applied)")
        elif not is_aligned and depends_on_time:
            # Depends but not aligned. Could be a summary or a scalar lookup.
            # Usually time-dependent variables are arrays aligned with time.
            # If it's a scalar depending on time (e.g. "Year 1 Discount Factor"), is it time dependent?
            # Design says: "Bindings aligned but not dependent (should have penalty). Bindings dependent but not aligned (should be time-dependent)."
            # Wait, "Bindings dependent but not aligned (should be time-dependent)" is listed under "Negative + Edge Cases".
            # Actually, if it depends on time but isn't aligned, it might be a specific point in time.
            # But usually "time dependent" implies "varies with time" -> vector.
            # If it's a single cell, it's a scalar.
            # Let's mark it as time-dependent but maybe lower confidence?
            # Or maybe check if it's a vector?
            # If it's a scalar (1x1) and depends on time, it's technically time dependent.
            # But for "projection columns", alignment is key.
            # Let's follow the design implication: "Bindings dependent but not aligned (should be time-dependent)".
            is_time_dependent = True
            confidence = TIME_GRAPH_CONFIDENCE_MIN
            reasons.append("Depends on time index but not spatially aligned")
        else:
            is_time_dependent = False
            confidence = 0.0
            reasons.append("Neither aligned nor dependent")

        annotations.append(
            TimeDependenceAnnotation(
                binding_id=binding.binding_id,
                time_index_binding_id=time_index.binding_id,
                is_time_dependent=is_time_dependent,
                confidence=confidence,
                reasons_top3=reasons[:3],
                evidence_flags=evidence_flags,
            )
        )

    return annotations


def _find_ancestors(target_id: str, reverse_deps: dict[str, list[str]]) -> set[str]:
    """Find all bindings that depend on target_id (ancestors in dependency graph)."""
    visited = set()
    queue = [target_id]
    while queue:
        current = queue.pop(0)
        if current in visited:
            continue
        visited.add(current)
        for ancestor in reverse_deps.get(current, []):
            if ancestor not in visited:
                queue.append(ancestor)
    return visited


def _check_alignment(time_index: TimeIndex, binding: BindingInfo) -> bool:
    """
    Check if binding is spatially aligned with time index.

    If time index is ROW (horizontal):
        Aligned if binding occupies same columns (overlap).
        Ideally start/end match, or binding is subset/superset.
        Allow blank separators?
        Design says: "Spatial alignment logic (projection-style layouts; tolerate blank separators)."

    If time index is COL (vertical):
        Aligned if binding occupies same rows.
    """
    if time_index.axis == "row":
        # Time index is horizontal (e.g. A1:Z1).
        # Binding should be horizontal and share columns.
        # Check column overlap.
        ti_start = time_index.top_left_col
        ti_end = time_index.top_left_col + time_index.shape_cols - 1

        b_start = binding.top_left_col
        b_end = binding.top_left_col + binding.shape_cols - 1

        return _ranges_aligned_with_gap(ti_start, ti_end, b_start, b_end)

    elif time_index.axis == "col":
        # Time index is vertical.
        # Check row overlap.
        ti_start = time_index.top_left_row
        ti_end = time_index.top_left_row + time_index.shape_rows - 1

        b_start = binding.top_left_row
        b_end = binding.top_left_row + binding.shape_rows - 1

        return _ranges_aligned_with_gap(ti_start, ti_end, b_start, b_end)

    return False


def _table_has_column(conn: sqlite3.Connection, table: str, column: str) -> bool:
    cursor = conn.cursor()
    cursor.execute(f"PRAGMA table_info({table})")
    return any(row[1] == column for row in cursor.fetchall())


def _table_exists(conn: sqlite3.Connection, table: str) -> bool:
    cursor = conn.cursor()
    cursor.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name = ?",
        (table,),
    )
    return cursor.fetchone() is not None


def _bindings_address_column(conn: sqlite3.Connection) -> str:
    if _table_has_column(conn, "bindings", "address_a1"):
        return "address_a1"
    if _table_has_column(conn, "bindings", "top_left_a1"):
        return "top_left_a1"
    raise RuntimeError("bindings table missing address column (address_a1/top_left_a1)")


def _ranges_aligned_with_gap(
    ti_start: int,
    ti_end: int,
    b_start: int,
    b_end: int,
) -> bool:
    overlap_start = max(ti_start, b_start)
    overlap_end = min(ti_end, b_end)
    ti_len = ti_end - ti_start + 1
    b_len = b_end - b_start + 1

    if overlap_start <= overlap_end:
        overlap_len = overlap_end - overlap_start + 1
        if overlap_len >= _MIN_ALIGNMENT_OVERLAP:
            return True
        if overlap_len == ti_len or overlap_len == b_len:
            return True
        return False

    gap = _range_gap(ti_start, ti_end, b_start, b_end)
    if gap is None:
        return False
    if gap > _MAX_SEPARATOR_GAP:
        return False
    if min(ti_len, b_len) < _MIN_ALIGNMENT_OVERLAP:
        return False
    if _length_similarity(ti_len, b_len) < _MIN_LENGTH_SIMILARITY:
        return False
    return True


def _range_gap(ti_start: int, ti_end: int, b_start: int, b_end: int) -> int | None:
    if b_start > ti_end:
        return b_start - ti_end - 1
    if ti_start > b_end:
        return ti_start - b_end - 1
    return None


def _length_similarity(a_len: int, b_len: int) -> float:
    if a_len <= 0 or b_len <= 0:
        return 0.0
    return min(a_len, b_len) / max(a_len, b_len)
