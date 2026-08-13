# ABOUTME: Basic JSON model specification generator for Sprint 1
# ABOUTME: Extracts labelled variables from overlay and generates machine-readable spec

import json
import logging
import re
import sqlite3
import time
from collections.abc import Callable, Iterable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from xl_marinade.core.db_uri import connect_read_only
from xl_marinade.core.grouping.geometry import BoundingBox, parse_a1_address
from xl_marinade.core.parser import (
    ASTNode,
    BinaryNode,
    FormulaParser,
    FunctionNode,
    RefNode,
    UnaryNode,
)
from xl_marinade.docs.semantic_index.axis_extraction import TIME_AXIS_NAME
from xl_marinade.docs.utils.formula_explainer import generate_explanation
from xl_marinade.docs.utils.ir_schema import detect_dependency_edges
from xl_marinade.docs.utils.range_matching import find_binding_for_range

logger = logging.getLogger(__name__)


def _log_timing(step: str, elapsed: float, extra: str = "") -> None:
    suffix = f" ({extra})" if extra else ""
    logger.info(f"TIMING json_spec_generator.{step}: {elapsed:.2f}s{suffix}")


# Optional dependency for validation
try:
    import jsonschema
except ImportError:
    jsonschema = None

# Schema path - loaded lazily in validate_json_spec
SCHEMA_PATH = Path(__file__).parent / "json_spec_schema_v0.5.json"


def _parse_a1_cell(cell_ref: str) -> tuple[int, int] | None:
    """
    Parse a single A1 cell reference to (column_index, row_index).

    Args:
        cell_ref: Cell reference like "A1", "Z99", "AA100"

    Returns:
        Tuple of (col, row) as 0-based indices, or None if invalid

    Example:
        >>> _parse_a1_cell("A1")
        (0, 0)
        >>> _parse_a1_cell("B5")
        (1, 4)
        >>> _parse_a1_cell("AA10")
        (26, 9)
    """
    match = re.match(r"^([A-Z]+)(\d+)$", cell_ref.upper())
    if not match:
        return None

    col_letters, row_str = match.groups()

    # Convert column letters to 0-based index (A=0, Z=25, AA=26, etc.)
    col_idx = 0
    for char in col_letters:
        col_idx = col_idx * 26 + (ord(char) - ord("A") + 1)
    col_idx -= 1

    row_idx = int(row_str) - 1  # Convert to 0-based

    return (col_idx, row_idx)


def _col_index_to_letters(col_idx: int) -> str:
    """
    Convert 0-based column index to Excel column letters.

    Args:
        col_idx: 0-based column index (0=A, 25=Z, 26=AA, etc.)

    Returns:
        Column letters (e.g., "A", "Z", "AA")

    Example:
        >>> _col_index_to_letters(0)
        "A"
        >>> _col_index_to_letters(25)
        "Z"
        >>> _col_index_to_letters(26)
        "AA"
    """
    result = []
    col_idx += 1  # Convert to 1-based for the algorithm

    while col_idx > 0:
        col_idx -= 1  # Adjust for 0-based modulo
        result.append(chr(ord("A") + (col_idx % 26)))
        col_idx //= 26

    return "".join(reversed(result))


def _expand_a1_range(range_str: str, max_cells: int = 10000) -> list[str]:
    """
    Expand an A1 range notation into individual cell addresses.

    Handles:
    - Single cells: "A1" -> ["A1"]
    - Ranges: "A1:A10" -> ["A1", "A2", ..., "A10"]
    - 2D ranges: "A1:B2" -> ["A1", "B1", "A2", "B2"]

    Args:
        range_str: A1 notation (single cell or range)
        max_cells: Maximum number of cells to expand (safety limit)

    Returns:
        List of individual cell addresses

    Raises:
        ValueError: If range is invalid or exceeds max_cells

    Example:
        >>> _expand_a1_range("A1")
        ["A1"]
        >>> _expand_a1_range("A1:A3")
        ["A1", "A2", "A3"]
        >>> _expand_a1_range("A1:B2")
        ["A1", "B1", "A2", "B2"]
    """
    # Strip any $ signs (absolute references)
    range_str = range_str.replace("$", "")

    # Check if it's a range or single cell
    if ":" not in range_str:
        # Single cell
        if _parse_a1_cell(range_str):
            return [range_str]
        else:
            raise ValueError(f"Invalid cell reference: {range_str}")

    # Parse range
    parts = range_str.split(":", 1)
    if len(parts) != 2:
        raise ValueError(f"Invalid range format: {range_str}")

    start_cell, end_cell = parts
    start_pos = _parse_a1_cell(start_cell)
    end_pos = _parse_a1_cell(end_cell)

    if not start_pos or not end_pos:
        raise ValueError(f"Invalid range references: {range_str}")

    start_col, start_row = start_pos
    end_col, end_row = end_pos

    # Calculate number of cells
    num_cols = abs(end_col - start_col) + 1
    num_rows = abs(end_row - start_row) + 1
    total_cells = num_cols * num_rows

    if total_cells > max_cells:
        logger.debug(
            f"Range {range_str} contains {total_cells} cells, exceeding max_cells={max_cells}. "
            f"Skipping expansion to avoid performance issues."
        )
        # For large ranges, return empty list to skip expansion
        # This prevents O(n²) explosion - caller must handle gracefully
        return []

    # Expand range
    cells = []
    for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
        for col in range(min(start_col, end_col), max(start_col, end_col) + 1):
            col_letters = _col_index_to_letters(col)
            cell = f"{col_letters}{row + 1}"  # Convert back to 1-based row
            cells.append(cell)

    return cells


def _serialize_value(val: str | None) -> Any:
    """
    Serialize a cell value for JSON output.

    Handles:
    - None -> null (not "null" string)
    - Numeric strings -> numbers
    - JSON strings -> parsed objects/arrays
    - Plain strings -> unchanged

    Args:
        val: Cell value_snapshot from IR (string or None)

    Returns:
        Properly typed value for JSON serialization
    """
    # Handle None explicitly - this becomes JSON null
    if val is None:
        return None

    # Empty string -> empty string (not null)
    if val == "":
        return ""

    # Try to parse as JSON (for arrays/objects stored as strings)
    if val.startswith("[") or val.startswith("{"):
        try:
            return json.loads(val)
        except json.JSONDecodeError:
            return val  # Return as-is if invalid JSON

    # Try to parse as number
    cleaned = val.strip()
    if cleaned:
        # Check if it's a number (proper validation for negative numbers and decimals)
        try:
            # Try parsing as float first (handles both int and float formats)
            num = float(cleaned)
            # Return as int if it's a whole number, else float
            return int(num) if num.is_integer() else num
        except ValueError:
            pass

    # Return as string
    return val


def _extract_range_values(cells: list[tuple[Any, ...]]) -> Any:
    """
    Extract evaluated values from cells, handling both single cells and ranges.

    For single cells: Returns typed value (number, string, null)
    For multi-cell ranges: Returns array (simple for now - just list of values)

    Uses evaluated_value (from data_only=True load) instead of value_snapshot.
    This ensures formula cells show calculated results, not formula text.

    Args:
        cells: List of (evaluated_value, formula_a1) tuples from cells table

    Returns:
        Typed value for single cell, or array of values for ranges
    """
    if not cells:
        return None

    # Single cell - return typed value
    if len(cells) == 1:
        return _serialize_value(cells[0][0])

    # Multi-cell range - return array of values
    # Use evaluated_value which contains actual calculated results
    values = [_serialize_value(cell[0]) for cell in cells]

    return values


def _chunked_list(values: list[Any], size: int) -> Iterable[list[Any]]:
    for start in range(0, len(values), size):
        yield values[start : start + size]


def _load_composite_members(
    overlay_conn: sqlite3.Connection,
) -> dict[str, list[tuple[str, str | None, str | None]]]:
    try:
        rows = overlay_conn.execute("""
            SELECT cb.composite_id, cb.ir_binding_id, b.sheet, b.address
            FROM composite_bindings cb
            JOIN ir.agent_bindings b ON cb.ir_binding_id = b.binding_id
            ORDER BY cb.composite_id, cb.ordinal
        """).fetchall()
    except sqlite3.OperationalError:
        rows = overlay_conn.execute("""
            SELECT cb.composite_id, cb.ir_binding_id, b.sheet, b.address_a1
            FROM composite_bindings cb
            JOIN ir.bindings b ON cb.ir_binding_id = b.binding_id
            ORDER BY cb.composite_id, cb.ordinal
        """).fetchall()

    members: dict[str, list[tuple[str, str | None, str | None]]] = {}
    for composite_id, ir_binding_id, sheet, address in rows:
        members.setdefault(composite_id, []).append((ir_binding_id, sheet, address))

    return members


def _load_cells_by_binding(
    overlay_conn: sqlite3.Connection,
    binding_ids: list[str],
) -> dict[str, list[tuple[str, Any, Any]]]:
    cells_by_binding: dict[str, list[tuple[str, Any, Any]]] = {}
    if not binding_ids:
        return cells_by_binding

    chunk_size = 900
    for chunk in _chunked_list(binding_ids, chunk_size):
        placeholders = ",".join("?" * len(chunk))
        try:
            rows = overlay_conn.execute(
                f"""
                SELECT ctb.binding_id, ac.cell_address, ac.value, ac.formula
                FROM ir.cell_to_binding ctb
                JOIN ir.agent_cells ac ON ctb.cell_id = ac.cell_id
                WHERE ctb.binding_id IN ({placeholders})
                ORDER BY ctb.binding_id, ac.cell_address
            """,
                chunk,
            ).fetchall()
        except sqlite3.OperationalError:
            rows = overlay_conn.execute(
                f"""
                SELECT c.binding_id, c.cell_address_a1, c.evaluated_value, c.formula_a1
                FROM ir.cells c
                WHERE c.binding_id IN ({placeholders})
                ORDER BY c.binding_id, c.cell_address_a1
            """,
                chunk,
            ).fetchall()
        for binding_id, cell_address, value, formula in rows:
            cells_by_binding.setdefault(binding_id, []).append((cell_address, value, formula))

    return cells_by_binding


def _load_binding_shapes(
    overlay_conn: sqlite3.Connection,
    binding_ids: list[str],
) -> dict[str, tuple[int, int, str | None]]:
    shapes: dict[str, tuple[int, int, str | None]] = {}
    if not binding_ids:
        return shapes

    chunk_size = 900
    for chunk in _chunked_list(binding_ids, chunk_size):
        placeholders = ",".join("?" * len(chunk))
        try:
            rows = overlay_conn.execute(
                f"""
                SELECT binding_id, shape_rows, shape_cols, address
                FROM ir.agent_bindings
                WHERE binding_id IN ({placeholders})
            """,
                chunk,
            ).fetchall()
        except sqlite3.OperationalError:
            rows = overlay_conn.execute(
                f"""
                SELECT binding_id, shape_rows, shape_cols, address_a1
                FROM ir.bindings
                WHERE binding_id IN ({placeholders})
            """,
                chunk,
            ).fetchall()
        for binding_id, shape_rows, shape_cols, address in rows:
            shapes[binding_id] = (shape_rows, shape_cols, address)

    return shapes


def _build_dependency_maps(
    overlay_conn: sqlite3.Connection,
) -> tuple[dict[str, set[str]], dict[str, set[str]]]:
    parents: dict[str, set[str]] = {}
    children: dict[str, set[str]] = {}

    edges = detect_dependency_edges(overlay_conn, schema="ir")
    if edges is None:
        return parents, children

    rows = overlay_conn.execute(
        f"SELECT {edges.from_col}, {edges.to_col} FROM ir.{edges.table}"
    ).fetchall()
    for from_binding, to_binding in rows:
        children.setdefault(from_binding, set()).add(to_binding)
        parents.setdefault(to_binding, set()).add(from_binding)

    return parents, children


def _load_axis_period_labels(
    overlay_conn: sqlite3.Connection,
    axis_binding_id: str,
) -> list[Any]:
    """Load the actual cells of a time-axis binding, in spatial order, as period labels.

    These are the real axis cells (e.g. the projection-period numbers in A7:A607),
    never invented. Ordered by (row, col) so t=1,2,3,... maps to the right cell.
    Returns [] if the axis cells cannot be read.
    """
    cells_by_binding = _load_cells_by_binding(overlay_conn, [axis_binding_id])
    cells = cells_by_binding.get(axis_binding_id, [])
    if not cells:
        return []

    def _sort_key(cell: tuple[str, Any, Any]) -> tuple[int, int]:
        address = cell[0]
        pure = address.split("!")[-1] if address and "!" in address else address
        parsed = _parse_a1_cell(pure) if pure else None
        # parsed is (col, row); order by row then col so a vertical axis reads top-down
        return (parsed[1], parsed[0]) if parsed else (0, 0)

    ordered = sorted(cells, key=_sort_key)
    return [_serialize_value(cell[1]) for cell in ordered]


def _build_time_axis_link(
    overlay_conn: sqlite3.Connection,
) -> dict[str, dict[str, Any]]:
    """Build the deterministic series -> time-axis link, gated against fabrication.

    Reads the IR-computed link directly from ir.binding_time_annotations
    (binding_id -> time_index_binding_id), GATED by
    `is_time_dependent = 1 AND MAX(shape_rows, shape_cols) > 1`. A 1x1 scalar that
    merely reads the time index is NOT a series and is excluded here, so it never
    gets a t-index.

    Returns a map binding_id -> axis dict
    {name: 'time', index_variable: <time_index_binding_id>, length, confidence,
     period_labels: [actual axis cells]}.
    Returns {} when the IR lacks binding_time_annotations (older builds).
    """
    try:
        rows = overlay_conn.execute(
            """
            SELECT a.binding_id,
                   a.time_index_binding_id,
                   a.confidence,
                   b.shape_rows,
                   b.shape_cols
            FROM ir.binding_time_annotations a
            JOIN ir.bindings b ON a.binding_id = b.binding_id
            WHERE a.is_time_dependent = 1
              AND MAX(b.shape_rows, b.shape_cols) > 1
            ORDER BY a.binding_id
            """
        ).fetchall()
    except sqlite3.OperationalError:
        # Older IR builds have no binding_time_annotations table.
        return {}

    # Resolve period labels once per distinct axis binding (deterministic, cached).
    period_labels_cache: dict[str, list[Any]] = {}
    link: dict[str, dict[str, Any]] = {}
    for binding_id, time_index_binding_id, confidence, shape_rows, shape_cols in rows:
        if time_index_binding_id not in period_labels_cache:
            period_labels_cache[time_index_binding_id] = _load_axis_period_labels(
                overlay_conn, time_index_binding_id
            )
        link[binding_id] = {
            "name": TIME_AXIS_NAME,
            "index_variable": time_index_binding_id,
            "length": max(shape_rows, shape_cols),
            "confidence": confidence,
            "period_labels": period_labels_cache[time_index_binding_id],
        }

    # Phase B2: propagate the time axis onto collapsed composite objects.
    # A composite's id is NOT an IR binding id, so the IR-keyed link above misses
    # it and the collapsed object would silently lose its B1 time axis. Resolve via
    # the composite's members: a composite gets a time axis ONLY when ALL of its
    # members carry the SAME gated time index (unanimous), so nothing is fabricated.
    # The 985 mortality family members agree unanimously, so this is well-defined.
    _propagate_time_axis_to_composites(overlay_conn, link)
    return link


def _propagate_time_axis_to_composites(
    overlay_conn: sqlite3.Connection,
    link: dict[str, dict[str, Any]],
) -> None:
    """Add composite_id -> time-axis entries to `link` (Phase B2, in place).

    A B2-collapsed object (985 mortality atoms -> one 985x1 vector) is the genuine
    SERIES; its individual members are 1x1 cells that merely READ the time axis, so
    B1's shape>1 gate excludes them from `link`. We must therefore resolve the
    composite's axis from the members' raw `binding_time_annotations` (gated only by
    is_time_dependent=1), and gate on the COMPOSITE's own shape (the composite is the
    series). A composite gets a time axis ONLY when:
      - every member is time-dependent and points to the SAME time index (unanimous),
      - AND the composite's own MAX(shape) > 1 (it is a real vector/grid, not a
        scalar) -- enforced because composite_shapes (computed downstream) always
        spans >1 for collapsed families; here we additionally require the time-index
        binding's length to be >1.
    Mixed / partial annotations get no axis (conservative; no fabrication).
    """
    try:
        member_rows = overlay_conn.execute(
            """
            SELECT cb.composite_id,
                   a.is_time_dependent,
                   a.time_index_binding_id,
                   a.confidence
            FROM composite_bindings cb
            JOIN ir.binding_time_annotations a ON a.binding_id = cb.ir_binding_id
            ORDER BY cb.composite_id, cb.ordinal
            """
        ).fetchall()
        member_counts = dict(
            overlay_conn.execute(
                "SELECT composite_id, COUNT(*) FROM composite_bindings GROUP BY composite_id"
            ).fetchall()
        )
    except sqlite3.OperationalError:
        # No composite_bindings or no binding_time_annotations (older builds).
        return

    # Aggregate per composite: collect (is_time_dependent, time_index, confidence).
    agg: dict[str, list[tuple[int, str, float]]] = {}
    for composite_id, is_td, tib, conf in member_rows:
        agg.setdefault(composite_id, []).append((is_td, tib, conf))

    period_labels_cache: dict[str, list[Any]] = {}
    for composite_id, annotations in agg.items():
        total_members = member_counts.get(composite_id, 0)
        # Require EVERY member to be annotated and time-dependent (unanimous).
        if len(annotations) != total_members or total_members == 0:
            continue
        if any(is_td != 1 for is_td, _tib, _c in annotations):
            continue
        index_vars = {tib for _is_td, tib, _c in annotations}
        if len(index_vars) != 1:
            continue
        time_index_binding_id = next(iter(index_vars))

        # Gate on the time index being a genuine axis (length > 1).
        shape_row = overlay_conn.execute(
            "SELECT shape_rows, shape_cols FROM ir.bindings WHERE binding_id = ?",
            (time_index_binding_id,),
        ).fetchone()
        if not shape_row:
            continue
        axis_length = max(shape_row[0], shape_row[1])
        if axis_length <= 1:
            continue

        if time_index_binding_id not in period_labels_cache:
            period_labels_cache[time_index_binding_id] = _load_axis_period_labels(
                overlay_conn, time_index_binding_id
            )
        # Use the minimum member confidence (conservative).
        confidence = min(c for _is_td, _tib, c in annotations if c is not None)
        link[composite_id] = {
            "name": TIME_AXIS_NAME,
            "index_variable": time_index_binding_id,
            "length": axis_length,
            "confidence": confidence,
            "period_labels": period_labels_cache[time_index_binding_id],
        }


def _infer_axes_from_shape(
    shape_rows: int,
    shape_cols: int,
    label: str,
    entity_type: str,
    binding_id: str | None = None,
    time_axis_link: dict[str, dict[str, Any]] | None = None,
) -> list[dict[str, Any]] | None:
    if entity_type not in ["time_series", "table"]:
        return None

    binding = {"shape_rows": shape_rows, "shape_cols": shape_cols}
    if entity_type == "time_series":
        axis = _infer_time_series_axis(binding, binding_id, time_axis_link)
        return [axis] if axis else None

    table_time_axis = time_axis_link.get(binding_id) if (time_axis_link and binding_id) else None
    return _infer_table_axes(binding, label, table_time_axis)


def _get_dependencies_from_maps(
    binding_id: str,
    is_composite: bool,
    composite_member_ids: dict[str, list[str]],
    parents_map: dict[str, set[str]],
    children_map: dict[str, set[str]],
    composite_lookup: dict[str, str],
) -> dict[str, list[str]]:
    if is_composite:
        binding_ids = composite_member_ids.get(binding_id)
        if not binding_ids:
            binding_ids = [binding_id]
    else:
        binding_ids = [binding_id]

    parent_ids: set[str] = set()
    child_ids: set[str] = set()
    for bid in binding_ids:
        parent_ids.update(parents_map.get(bid, set()))
        child_ids.update(children_map.get(bid, set()))

    resolved_parents = sorted({composite_lookup.get(pid, pid) for pid in parent_ids})
    resolved_children = sorted({composite_lookup.get(cid, cid) for cid in child_ids})

    return {"parents": resolved_parents, "children": resolved_children}


def create_metadata(
    overlay_conn: sqlite3.Connection, ir_conn: sqlite3.Connection
) -> dict[str, Any]:
    """
    Create metadata section for JSON spec.

    Args:
        overlay_conn: Connection to semantic overlay database
        ir_conn: Connection to Phase 1 IR database

    Returns:
        Metadata dict with model_name, generated_at, workbook_guid, overlay_version
    """
    # Get overlay metadata
    overlay_meta_rows = overlay_conn.execute("SELECT key, value FROM metadata").fetchall()
    overlay_meta = dict(overlay_meta_rows)

    # Get IR metadata (if available)
    ir_meta: dict[str, Any] = {}
    try:
        ir_meta_row = ir_conn.execute("SELECT * FROM meta").fetchone()
        if ir_meta_row:
            # Convert row to dict using column names
            cols = [desc[0] for desc in ir_conn.execute("SELECT * FROM meta").description]
            ir_meta = dict(zip(cols, ir_meta_row, strict=True))
        else:
            logger.warning("No metadata found in IR database meta table")
    except sqlite3.Error:
        # Fall back to fast schema metadata table.
        try:
            rows = ir_conn.execute("SELECT key, value FROM ir_metadata").fetchall()
            ir_meta = dict(rows)
        except sqlite3.Error as e:
            # Log database errors - these indicate schema issues that should be investigated
            logger.error(f"Failed to query IR metadata: {e}")
            ir_meta = {}
    except ValueError as e:
        # Log schema mismatch - indicates IR schema changed
        logger.error(f"IR metadata schema mismatch: {e}")
        ir_meta = {}

    # Build metadata. Model-name precedence (deterministic; SHA only as last
    # resort): verbatim docProps dc:title > display name derived from the
    # original filename > legacy meta.file_path > workbook SHA.
    metadata = {
        "model_name": _pick_model_name(ir_meta),
        "schema_version": "0.5.0",
        "generated_at": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "workbook_guid": overlay_meta.get("ir_workbook_guid", "unknown"),
        "overlay_version": overlay_meta.get("overlay_version", "0.1"),
    }

    return metadata


# Trailing tokens stripped from a workbook basename to derive a display name:
# parenthetical copy markers "(1)", version tags "v1.5"/"_v2", and ISO dates.
# Conservative by design — only matches at the END of the string.
_VERSION_TAIL_RE = re.compile(
    r"[\s_\-]*(?:\(\d+\)|v\d+(?:\.\d+)*|\d{4}-\d{2}-\d{2})\s*$",
    re.IGNORECASE,
)
_WORKBOOK_EXT_RE = re.compile(r"\.(?:xls[xmb]?|xlam|xla)$", re.IGNORECASE)

# A bare cell address used as a label, e.g. "Calculations!D110" or
# "'Cash Flow'!B5:B90" — the labeller's last-resort fallback, not a real name.
_RAW_ADDRESS_RE = re.compile(r"^(?:'[^']+'|[^!]+)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$")


def _looks_like_raw_address(label: str) -> bool:
    """True when ``label`` is just a Sheet!A1[:A1] cell address."""
    return bool(_RAW_ADDRESS_RE.match(label.strip()))


def _derive_display_name(filename: str) -> str:
    """Deterministically derive a display name from a workbook basename.

    Strips the spreadsheet extension and a small set of well-known TRAILING
    tokens (copy markers, version tags, ISO dates). Never touches mid-string
    content — aggressive cleanup of long descriptive filenames is a deferred
    semantic-layer upgrade.
    """
    name = _WORKBOOK_EXT_RE.sub("", filename)
    for _ in range(3):
        stripped = _VERSION_TAIL_RE.sub("", name)
        if stripped == name:
            break
        name = stripped
    name = name.strip()
    return name or filename


def _pick_model_name(ir_meta: dict[str, Any]) -> str:
    """Resolve the human-facing model name from IR metadata, SHA as last resort."""
    doc_title = ir_meta.get("doc_title")
    if doc_title:
        return doc_title
    filename = ir_meta.get("original_filename") or ir_meta.get("file_path")
    if filename:
        return _derive_display_name(filename)
    return ir_meta.get("workbook_sha256", "Unknown Model")


def _resolve_ir_binding_to_variable_id(overlay_conn: sqlite3.Connection, ir_binding_id: str) -> str:
    """
    Resolve an IR binding_id to its overlay variable_id.

    This is necessary because the JSON spec uses overlay variable_ids (which can be
    composite_* or time_series_* prefixed), but the IR dependency edges use raw
    IR binding_ids (64-char SHA-256 hashes).

    If the IR binding is part of a composite/time_series binding, return the composite_id.
    Otherwise, return the IR binding_id itself.

    Args:
        overlay_conn: Connection to overlay (must have IR attached)
        ir_binding_id: IR binding ID to resolve

    Returns:
        Overlay variable_id (either composite_id or ir_binding_id)
    """
    # Check if this IR binding is part of a composite
    composite_row = overlay_conn.execute(
        """
        SELECT composite_id 
        FROM composite_bindings 
        WHERE ir_binding_id = ?
        LIMIT 1
    """,
        (ir_binding_id,),
    ).fetchone()

    if composite_row:
        return composite_row[0]  # Return composite_id
    else:
        return ir_binding_id  # Return IR binding_id as-is


def _get_dependencies(overlay_conn: sqlite3.Connection, binding_id: str) -> dict[str, list[str]]:
    """
    Get dependency relationships for a binding.

    Handles both regular and composite bindings. For composite bindings,
    aggregates dependencies from all source bindings.

    IMPORTANT: Returns overlay variable_ids (which may be composite_* or time_series_*
    prefixed), not raw IR binding_ids. This ensures dependencies can be traversed
    using the variable_id field in the JSON spec.

    Args:
        overlay_conn: Connection to overlay (must have IR attached)
        binding_id: Binding ID to query

    Returns:
        Dict with 'parents' and 'children' keys containing lists of overlay variable_ids
    """
    # Check if this is a composite binding
    composite_members = overlay_conn.execute(
        """
        SELECT ir_binding_id
        FROM composite_bindings
        WHERE composite_id = ?
        ORDER BY ordinal
    """,
        (binding_id,),
    ).fetchall()

    # Get all binding IDs to query (self if regular, all members if composite)
    binding_ids = [row[0] for row in composite_members] if composite_members else [binding_id]

    # Collect parent and child binding IDs (using sets to deduplicate)
    parent_binding_ids: set[str] = set()
    child_binding_ids: set[str] = set()

    edges = detect_dependency_edges(overlay_conn, schema="ir")
    if edges is None:
        return {"parents": [], "children": []}

    for bid in binding_ids:
        # Get parents (bindings that depend on this one)
        # from_binding depends on to_binding, so parents have edges FROM them TO us
        parents = overlay_conn.execute(
            f"""
            SELECT DISTINCT {edges.from_col}
            FROM ir.{edges.table}
            WHERE {edges.to_col} = ?
            """,
            (bid,),
        ).fetchall()
        parent_binding_ids.update(row[0] for row in parents)

        # Get children (bindings this one depends on)
        # from_binding depends on to_binding, so children have edges FROM us TO them
        children = overlay_conn.execute(
            f"""
            SELECT DISTINCT {edges.to_col}
            FROM ir.{edges.table}
            WHERE {edges.from_col} = ?
            """,
            (bid,),
        ).fetchall()
        child_binding_ids.update(row[0] for row in children)

    # Resolve IR binding IDs to overlay variable IDs
    # (IR bindings that are part of composite/time_series bindings need to be mapped)
    resolved_parents = [
        _resolve_ir_binding_to_variable_id(overlay_conn, bid) for bid in parent_binding_ids
    ]
    resolved_children = [
        _resolve_ir_binding_to_variable_id(overlay_conn, bid) for bid in child_binding_ids
    ]

    # Deduplicate and sort (multiple IR bindings might resolve to same composite)
    return {
        "parents": sorted(list(set(resolved_parents))),
        "children": sorted(list(set(resolved_children))),
    }


def substitute_range_with_label(
    formula: str, range_a1: str, label: str, sheet_name: str | None = None
) -> str:
    """
    Substitute a range reference in a formula with its semantic label.

    Handles multiple range formats:
    - Unqualified: A1:B10
    - Absolute: $A$1:$B$10
    - Mixed: $A1:B$10
    - Sheet-qualified: Sheet!A1:B10
    - Quoted sheet: 'Sheet Name'!A1:B10

    Args:
        formula: The Excel formula string
        range_a1: The range to replace (in A1 notation, no $ signs)
        label: The semantic label to substitute
        sheet_name: Optional sheet name for qualified references

    Returns:
        Formula with range replaced by label
    """
    if not formula or not range_a1 or not label:
        return formula

    # Normalize the range (remove $ signs for matching)
    normalized_range = range_a1.replace("$", "")

    # Build regex pattern to match various forms of this range
    # We need to match: range, $range, Sheet!range, 'Sheet'!range
    # Allow $ anywhere in the range

    # Escape the range for regex, but then allow $ signs between characters
    # Split range into parts: start_cell : end_cell
    if ":" in normalized_range:
        start, end = normalized_range.split(":")

        # Build pattern allowing $ before each letter and number
        def make_flexible(cell: str) -> str:
            """Convert A1 to \\$?A\\$?1 pattern."""
            # Extract column letters and row number
            match = re.match(r"^([A-Z]+)(\d+)$", cell.upper())
            if not match:
                return re.escape(cell)
            col, row = match.groups()
            # Allow $ before column and before row
            return r"\$?" + re.escape(col) + r"\$?" + re.escape(row)

        flex_start = make_flexible(start)
        flex_end = make_flexible(end)
        flex_range = flex_start + ":" + flex_end
    else:
        # Single cell treated as range
        flex_range = re.escape(normalized_range).replace(r":", r"\$?:\$?")

    # Build list of patterns to try (order matters - try sheet-qualified first)
    patterns = []

    if sheet_name:
        escaped_sheet = re.escape(sheet_name)
        # Pattern 1: Quoted sheet first (more specific)
        patterns.append(r"'" + escaped_sheet + r"'!" + flex_range)
        # Pattern 2: Unquoted sheet
        patterns.append(escaped_sheet + r"!" + flex_range)

    # Pattern 3: Unqualified range (with optional $)
    patterns.append(flex_range)

    # Apply substitution for each pattern (case-insensitive for sheet names)
    result = formula
    for pattern in patterns:
        # Use word boundaries to avoid partial matches
        # But ranges might be in function args, so just use the pattern directly
        result = re.sub(pattern, label, result, flags=re.IGNORECASE)

    return result


def _build_binding_ranges(
    processed_rows: list[tuple],
) -> dict[tuple[str | None, BoundingBox], str]:
    """
    Build range-level lookup structure for O(m) matching.

    Uses existing BoundingBox from xl_marinade.core - no new dataclass needed.

    Returns:
        Dict mapping (sheet_normalized, BoundingBox) → label
    """
    binding_ranges: dict[tuple[str | None, BoundingBox], str] = {}

    for row in processed_rows:
        # Explicit unpacking matching extract_variable_data structure
        # row structure: (binding_id, label, actuarial_class, recon_required, recon_rationale,
        #                 sheet, address, is_composite, label_conf, class_conf, action, parameters_json)
        label = row[1]
        sheet = row[5]
        address = row[6]

        if not address or not label:
            continue

        # Normalize sheet name for case-insensitive matching
        sheet_normalized = sheet.lower().strip("'\"") if sheet else None

        try:
            # Strip sheet prefix if present in address
            pure_address = address
            if "!" in address:
                pure_address = address.split("!")[-1]

            bounds = parse_a1_address(pure_address.replace("$", ""))
            binding_ranges[(sheet_normalized, bounds)] = label
        except ValueError:
            # Skip invalid addresses (e.g., named ranges)
            continue

    return binding_ranges


def _replace_refs_in_ast(
    node: ASTNode,
    variable_map: dict[str, str],
    binding_ranges: dict[tuple[str | None, BoundingBox], str],
    current_sheet: str | None = None,
) -> ASTNode:
    """
    Recursively replace RefNode instances with semantic labels.

    IMMUTABLE: Creates new nodes, does not modify input.

    Handles both single-cell references (A1) and range references (A1:A10).
    For ranges, uses bounding box matching first, then falls back to variable_map.

    Args:
        node: AST node to transform
        variable_map: Map of canonical address → semantic label
        binding_ranges: Map of (sheet, bbox) → label for range matching
        current_sheet: Sheet context for relative refs

    Returns:
        New AST with refs replaced
    """
    if isinstance(node, RefNode):
        # Try to replace this reference
        ref = node.ref
        normalized = ref.replace("$", "")

        if ":" in normalized:
            # Range reference - try bounding box matching first (O(m))
            label = find_binding_for_range(normalized, current_sheet, binding_ranges)
            if label:
                return RefNode(label)

            # Fallback: Try direct lookup in variable_map (e.g. for exact string matches of ranges)
            # This handles cases where variable_map has "A1:A10" -> "Label" directly
            # Construct key
            lookup_key = normalized
            if "!" not in normalized and current_sheet:
                lookup_key = f"{current_sheet}!{normalized}"
            elif "!" in normalized:
                parts = normalized.split("!", 1)
                sheet_part = parts[0].strip("'\"")
                addr_part = parts[1]
                lookup_key = f"{sheet_part}!{addr_part}"

            if lookup_key in variable_map:
                return RefNode(variable_map[lookup_key])

            # No match - keep original
            return node
        else:
            # Single cell - try variable_map first (O(1))
            # Construct key for direct lookup
            lookup_key = normalized
            if "!" not in normalized and current_sheet:
                lookup_key = f"{current_sheet}!{normalized}"
            elif "!" in normalized:
                parts = normalized.split("!", 1)
                sheet_part = parts[0].strip("'\"")
                addr_part = parts[1]
                lookup_key = f"{sheet_part}!{addr_part}"

            if lookup_key in variable_map:
                return RefNode(variable_map[lookup_key])

            # Fallback: check if cell is in any binding range
            # This covers cases where range was too large for variable_map
            label = find_binding_for_range(normalized, current_sheet, binding_ranges)
            if label:
                return RefNode(label)

            # No match - keep original
            return node

    elif isinstance(node, FunctionNode):
        # Recurse into function arguments
        new_args = [
            _replace_refs_in_ast(arg, variable_map, binding_ranges, current_sheet)
            for arg in node.args
        ]
        return FunctionNode(node.name, new_args)

    elif isinstance(node, UnaryNode):
        new_operand = _replace_refs_in_ast(
            node.operand, variable_map, binding_ranges, current_sheet
        )
        return UnaryNode(node.operator, new_operand)

    elif isinstance(node, BinaryNode):
        new_left = _replace_refs_in_ast(node.left, variable_map, binding_ranges, current_sheet)
        new_right = _replace_refs_in_ast(node.right, variable_map, binding_ranges, current_sheet)
        return BinaryNode(node.operator, new_left, new_right)

    else:
        # ConstNode - no replacement needed
        return node


def _extract_array_formula_metadata(
    binding_id: str,
    overlay_conn: sqlite3.Connection,
    variable_map: dict[str, str],
    binding_ranges: dict[tuple[str | None, BoundingBox], str],
    sheet: str,
    cells: list[tuple[Any, Any]] | None = None,
    semantic_formula_builder: Callable[[str, str | None], str] | None = None,
) -> dict[str, Any] | None:
    """Extract array formula metadata if first cell differs from second.

    For bindings with multiple cells, compares first and second cell formulas.
    If different (e.g., initialization vs propagation), creates metadata dict
    with both formulas and their semantic versions.

    Args:
        binding_id: Binding identifier
        overlay_conn: Database connection to overlay (with IR attached)
        variable_map: Maps cell addresses to variable labels
        binding_ranges: Maps ranges to variable labels
        sheet: Current sheet name
        cells: Optional (value, formula) tuples to avoid per-binding query

    Returns:
        Dict with first_cell_formula, propagation_formula, and semantic versions,
        or None if metadata not needed (single cell, identical formulas, or errors)

    Example:
        >>> # Month counter: first=0, second=A8+1
        >>> meta = _extract_array_formula_metadata("bind123", conn, {}, {}, "Sheet1")
        >>> meta["first_cell_formula"]
        "=0"
        >>> meta["propagation_formula"]
        "=A8+1"
    """
    try:
        if cells is None:
            # Query first 2 cells only (performance: O(1) per binding)
            try:
                cell_rows = overlay_conn.execute(
                    """
                    SELECT ac.cell_address, ac.formula
                    FROM ir.cell_to_binding ctb
                    JOIN ir.agent_cells ac ON ctb.cell_id = ac.cell_id
                    WHERE ctb.binding_id = ?
                    ORDER BY ac.cell_address
                    LIMIT 2
                """,
                    (binding_id,),
                ).fetchall()
            except sqlite3.OperationalError:
                # Legacy/test schema: ir.cells(cell_address_a1, binding_id, formula_a1)
                cell_rows = overlay_conn.execute(
                    """
                    SELECT c.cell_address_a1 as cell_address, c.formula_a1 as formula
                    FROM ir.cells c
                    WHERE c.binding_id = ?
                    ORDER BY c.cell_address_a1
                    LIMIT 2
                """,
                    (binding_id,),
                ).fetchall()

            # Need at least 2 cells for array metadata
            if len(cell_rows) < 2:
                return None

            # Handle both tuple and dict-like row access
            # Try dict access first (if row_factory is set), fallback to tuple
            try:
                first_formula = cell_rows[0]["formula"]
                second_formula = cell_rows[1]["formula"]
            except (TypeError, KeyError):
                # Tuple access: (cell_address, formula)
                first_formula = cell_rows[0][1]
                second_formula = cell_rows[1][1]
        else:
            if len(cells) < 2:
                return None
            first_formula = cells[0][1]
            second_formula = cells[1][1]

        # Skip if either is None (constants)
        if first_formula is None or second_formula is None:
            return None

        # Skip if formulas are identical (not interesting)
        if first_formula == second_formula:
            return None

        # Generate semantic versions for both
        first_semantic = None
        second_semantic = None

        try:
            if semantic_formula_builder:
                first_semantic = semantic_formula_builder(first_formula, sheet)
            else:
                first_semantic = generate_semantic_formula(
                    first_formula, variable_map, binding_ranges, sheet
                )
        except Exception as e:
            logger.warning(f"Failed to generate semantic formula for first cell: {e}")

        try:
            if semantic_formula_builder:
                second_semantic = semantic_formula_builder(second_formula, sheet)
            else:
                second_semantic = generate_semantic_formula(
                    second_formula, variable_map, binding_ranges, sheet
                )
        except Exception as e:
            logger.warning(f"Failed to generate semantic formula for second cell: {e}")

        return {
            "first_cell_formula": first_formula,
            "propagation_formula": second_formula,
            "first_cell_semantic": first_semantic,
            "propagation_semantic": second_semantic,
        }

    except Exception as e:
        logger.warning(f"Failed to extract array formula metadata for {binding_id}: {e}")
        return None


def generate_semantic_formula(
    formula: str,
    variable_map: dict[str, str],
    binding_ranges: dict[tuple[str | None, BoundingBox], str],
    current_sheet: str | None = None,
) -> str:
    """
    Replace cell references in formula with semantic labels using AST.

    This ensures ONLY actual cell references are replaced, preserving
    string literals, function names, and other non-reference tokens.

    Args:
        formula: Excel formula string
        variable_map: Dict mapping canonical address (Sheet!A1 or A1) to Label
        binding_ranges: Dict mapping (sheet, bbox) to Label
        current_sheet: Name of sheet where formula resides (for relative refs)

    Returns:
        Formula with labels instead of addresses

    Example:
        >>> generate_semantic_formula("=A1+B1", {"A1": "Revenue", "B1": "Cost"}, {})
        "=Revenue+Cost"

        >>> generate_semantic_formula('="Check A1"', {"A1": "Revenue"}, {})
        '="Check A1"'  # String literal preserved
    """
    # Validation: handle empty input
    if not formula:
        return formula

    original_formula = formula
    had_leading_equals = formula.startswith("=")

    if not had_leading_equals:
        # Treat as non-formula unless it clearly looks like a formula expression.
        # This prevents truncation like "Just text" -> "Just" during parsing.
        looks_like_expr = (
            ":" in formula
            or re.search(r"\b\$?[A-Za-z]{1,3}\$?\d+\b", formula) is not None
            or any(op in formula for op in ("+", "-", "*", "/", "^", "&", "(", ")", ","))
        )
        if not looks_like_expr:
            return original_formula
        formula = f"={formula}"

    # Defensive: handle None or empty variable_map
    if not variable_map and not binding_ranges:
        return original_formula

    def _normalize_output(result: str) -> str:
        if not had_leading_equals and result.startswith("="):
            return result[1:]
        return result

    # Fast path: single-cell reference (no ranges/functions), avoid AST parse.
    stripped = formula.strip()
    if ":" not in stripped:
        sheet_name = None
        cell_ref = None
        sheet_match = re.match(
            r"^=\s*(?:(?P<sheet_q>'[^']+')|(?P<sheet_u>[^!]+))!(?P<cell>\$?[A-Za-z]{1,3}\$?\d+)\s*$",
            stripped,
        )
        if sheet_match:
            sheet_name = sheet_match.group("sheet_q") or sheet_match.group("sheet_u")
            if sheet_name:
                sheet_name = sheet_name.strip("'\"").strip()
            cell_ref = sheet_match.group("cell")
        else:
            cell_match = re.match(
                r"^=\s*(?P<cell>\$?[A-Za-z]{1,3}\$?\d+)\s*$",
                stripped,
            )
            if cell_match:
                cell_ref = cell_match.group("cell")

        if cell_ref:
            cell_ref = cell_ref.replace("$", "")
            sheet_for_lookup = sheet_name or current_sheet
            if sheet_for_lookup:
                key = f"{sheet_for_lookup}!{cell_ref}"
                label = variable_map.get(key)
                if not label and binding_ranges:
                    label = find_binding_for_range(
                        cell_ref,
                        sheet_for_lookup,
                        binding_ranges,
                    )
                if label:
                    return _normalize_output(f"={label}")

    try:
        # Parse to AST
        parser = FormulaParser()
        ast = parser.parse(formula)

        # Replace refs in AST (immutable transformation)
        transformed_ast = _replace_refs_in_ast(ast, variable_map, binding_ranges, current_sheet)

        # Convert back to string
        result = "=" + transformed_ast.to_string()

        return _normalize_output(result)

    except Exception as e:
        # If parsing fails, return original formula (graceful degradation)
        # In production, log this warning
        logger.warning(f"Failed to parse formula '{formula}': {e}. Returning original.")
        return original_formula


# NOTE: Garbage filtering logic moved to refiners/cleanup_refiner.py
# The generator is now a read-only projection - all filtering happens via mutations

# NOTE: Fragment merging logic moved to refiners/cleanup_refiner.py
# The generator is now a read-only projection - all merging happens via mutations


def _parse_cell_address(address_a1: str) -> tuple[int, int] | None:
    """
    Parse A1 cell address to (row, col) 0-indexed coordinates.

    Args:
        address_a1: Cell address like "Sheet!A1" or "A1:B2"

    Returns:
        (row, col) tuple or None if parse fails
    """
    import re

    # Strip sheet prefix if present
    if "!" in address_a1:
        address_a1 = address_a1.split("!")[-1]

    # Remove $ signs
    address_a1 = address_a1.replace("$", "")

    # Extract first cell if range
    if ":" in address_a1:
        address_a1 = address_a1.split(":")[0]

    # Parse column letters and row number
    match = re.match(r"^([A-Z]+)(\d+)$", address_a1.upper())
    if not match:
        return None

    col_letters, row_str = match.groups()

    # Convert column letters to 0-based index
    col_idx = 0
    for char in col_letters:
        col_idx = col_idx * 26 + (ord(char) - ord("A") + 1)
    col_idx -= 1

    row_idx = int(row_str) - 1
    return (row_idx, col_idx)


def _calculate_composite_bounding_box(
    member_data: list[tuple[str, int, int, str]],
) -> tuple[int, int]:
    """
    Calculate bounding box dimensions for composite binding members (Story 22).

    Fixes bug where max(shape_rows), max(shape_cols) was used instead of
    calculating the actual bounding box from member addresses.

    Example:
        Members: N10 (1x1), O10 (1x1), P10 (1x1)
        Old (WRONG): max(1,1,1) x max(1,1,1) = 1x1
        New (CORRECT): Bounding box N10:P10 = 1x3

    Args:
        member_data: List of (binding_id, shape_rows, shape_cols, address_a1) tuples

    Returns:
        (shape_rows, shape_cols) tuple for bounding box
    """
    if not member_data:
        return (1, 1)

    # Parse all member addresses to get coordinates
    coords = []
    for binding_id, shape_rows, shape_cols, address_a1 in member_data:
        parsed = _parse_cell_address(address_a1)
        if parsed:
            top_row, left_col = parsed
            # Calculate bottom-right corner
            bottom_row = top_row + shape_rows - 1
            right_col = left_col + shape_cols - 1
            coords.append(
                {
                    "top_row": top_row,
                    "left_col": left_col,
                    "bottom_row": bottom_row,
                    "right_col": right_col,
                }
            )

    if not coords:
        # Fallback to max aggregation if parsing failed
        logger.warning(
            f"Bounding box calculation fallback: Failed to parse addresses for "
            f"{len(member_data)} composite members. Using max() aggregation instead. "
            f"This may produce incorrect shape for non-overlapping fragments. "
            f"Addresses: {[m[3] for m in member_data]}"
        )
        return (max(m[1] for m in member_data), max(m[2] for m in member_data))

    # Calculate bounding box
    min_row = min(c["top_row"] for c in coords)
    max_row = max(c["bottom_row"] for c in coords)
    min_col = min(c["left_col"] for c in coords)
    max_col = max(c["right_col"] for c in coords)

    shape_rows = max_row - min_row + 1
    shape_cols = max_col - min_col + 1

    return (shape_rows, shape_cols)


def _infer_entity_type_from_shape(shape_rows: int, shape_cols: int, has_formula: bool) -> str:
    """
    Infer entity type from binding shape and formula presence.

    Rules:
    - 1x1 = scalar
    - 1xN or Nx1 (N>1) with formulas = time_series (calculation column/row)
    - 1xN or Nx1 (N>1) without formulas = vector (input data)
    - NxM (N>1, M>1) = table

    Args:
        shape_rows: Number of rows in binding
        shape_cols: Number of columns in binding
        has_formula: Whether binding contains formulas

    Returns:
        Entity type string: "scalar", "vector", "time_series", or "table"
    """
    if shape_rows == 1 and shape_cols == 1:
        return "scalar"
    elif shape_rows > 1 and shape_cols > 1:
        return "table"
    elif (shape_rows > 1 and shape_cols == 1) or (shape_rows == 1 and shape_cols > 1):
        # Single column/row - distinguish between time_series (calculated) and vector (input)
        if has_formula:
            return "time_series"
        else:
            return "vector"
    else:
        return "scalar"  # Fallback


def _infer_time_series_axis(
    binding: dict[str, Any],
    binding_id: str | None,
    time_axis_link: dict[str, dict[str, Any]] | None,
) -> dict[str, Any] | None:
    """
    Surface the real time axis for a time_series binding, or None.

    Looks up the IR-computed series -> time-axis link (built by
    `_build_time_axis_link`, gated by is_time_dependent=1 AND MAX(shape)>1). If the
    binding is a genuine, spatially-aligned time series it returns a labelled axis
    {name: 'time', index_variable: <axis binding>, length, period_labels}. Otherwise
    returns None — a 1x1 scalar or any unlinked binding never gets a fabricated
    t-index.

    The old label-substring / magic-number-length heuristic (which always returned
    index_variable=None) has been removed: it invented axis names from cell-count
    coincidences and could not point at the real axis.
    """
    if not binding_id or not time_axis_link:
        return None
    axis = time_axis_link.get(binding_id)
    if not axis:
        return None
    return {
        "name": axis["name"],
        "index_variable": axis["index_variable"],
        "length": axis["length"],
        "period_labels": axis["period_labels"],
        # A 1D series aligned to the IR-proven time index is a genuine, certain axis.
        "axis_confidence": "certain",
    }


def _infer_table_axes(
    binding: dict[str, Any],
    label: str,
    time_axis: dict[str, Any] | None = None,
) -> list[dict[str, Any]] | None:
    """
    Infer row and column axes for table entity.

    Uses generic row/column axes. If the IR proved this 2D grid is aligned to the
    time axis (Phase B1: `time_axis` is the gated link entry), the matching axis is
    named 'time' and carries the index_variable + period labels, instead of the
    generic 'Row'/'Column'. Future: Could infer the other axis from headers/labels.

    Args:
        binding: Dict with shape_rows, shape_cols
        label: Semantic label of the binding
        time_axis: The gated time-axis link entry for this binding, or None.

    Returns:
        List of two axis dicts (row, column), or None if cannot infer
    """
    shape_rows, shape_cols = binding["shape_rows"], binding["shape_cols"]

    # Phase B2: a generic Row/Column axis on a collapsed grid is SHAPE-INFERRED, not
    # certain — the second (non-time) axis of a 2D object is an age/cohort/scenario
    # dimension whose semantic name we have NOT proven. Flag it so the writer never
    # presents a possibly-mislabeled axis as definite.
    row_axis = {
        "name": "Row",
        "index_variable": None,
        "length": shape_rows,
        "axis_confidence": "shape_inferred",
    }
    col_axis = {
        "name": "Column",
        "index_variable": None,
        "length": shape_cols,
        "axis_confidence": "shape_inferred",
    }

    if time_axis:
        # Name whichever axis matches the IR-proven time-axis length 'time'. The
        # time axis is IR-proven, so it is 'certain'; the OTHER axis stays
        # shape_inferred.
        time_named = {
            "name": time_axis["name"],
            "index_variable": time_axis["index_variable"],
            "length": time_axis["length"],
            "period_labels": time_axis["period_labels"],
            "axis_confidence": "certain",
        }
        if shape_rows == time_axis["length"]:
            row_axis = time_named
        elif shape_cols == time_axis["length"]:
            col_axis = time_named

    return [row_axis, col_axis]


def _axes_from_geometry(
    geometry_axes: list[dict[str, Any]],
    geometry_kind: str,
    time_axis: dict[str, Any] | None,
) -> list[dict[str, Any]] | None:
    """Surface a collapsed family's axes from the SHAPE-AWARE classifier output.

    THE LOAD-BEARING SURFACING FIX (Phase B2 adversarial finding). The classifier in
    entity_grouper.classify_family_geometry already computes member-ACCURATE axis
    lengths (e.g. a 2-member-rows-apart family is length 2, not the 5-row bounding
    box). Earlier this metadata was discarded and axes were re-derived from the union
    bounding box, which inflated sparse/strided families into dense false objects
    (G2:G1011 = 2 cells surfaced as a 1010-element axis; C31:H35 = 2 rows surfaced as
    5). Here we surface the classifier's lengths verbatim instead.

    geometry_axes entries are {"orientation": row|column|block, "length": int,
    "confidence": str, ["stride": int]} as emitted by the classifier. We translate
    orientation -> axis name (Row/Column/Block) and carry the classifier's own
    per-axis confidence so the writer never presents a shape-inferred / strided axis
    as certain.

    B1 time axis (required-change #1 & #4): attach the IR-proven time index ONLY to
    the axis whose member-accurate LENGTH matches the time index extent — never to a
    multi-axis grid wholesale, and never against an inflated bounding-box length. The
    other axis stays its shape-inferred age/cohort/scenario self.
    """
    if not geometry_axes:
        return None

    name_for = {"row": "Row", "column": "Column", "block": "Block"}
    surfaced: list[dict[str, Any]] = []
    time_attached = False
    for ax in geometry_axes:
        orientation = ax.get("orientation", "")
        length = ax.get("length")
        confidence = ax.get("confidence", "shape_inferred")
        if time_axis is not None and not time_attached and length == time_axis["length"]:
            # The classifier's member-accurate length matches the IR time index:
            # this axis IS the (certain) time axis.
            surfaced.append(
                {
                    "name": time_axis["name"],
                    "index_variable": time_axis["index_variable"],
                    "length": time_axis["length"],
                    "period_labels": time_axis["period_labels"],
                    "axis_confidence": "certain",
                }
            )
            time_attached = True
            continue
        surfaced.append(
            {
                "name": name_for.get(orientation, "Axis"),
                "index_variable": None,
                "length": length,
                "axis_confidence": confidence,
            }
        )

    return surfaced


def _infer_axes_from_binding(
    overlay_conn: sqlite3.Connection,
    binding_id: str,
    entity_type: str,
    time_axis_link: dict[str, dict[str, Any]] | None = None,
) -> list[dict[str, Any]] | None:
    """
    Infer axes metadata from binding shape and label.

    Only infers axes for time_series and table entities. Returns None for other types.

    Args:
        overlay_conn: Connection to overlay (must have IR attached)
        binding_id: Binding ID to query
        entity_type: Entity type (scalar, vector, time_series, table)

    Returns:
        List of axis dicts, or None if axes not applicable or cannot be inferred
    """
    if entity_type not in ["time_series", "table"]:
        return None  # Only time_series and tables have axes

    # Get binding metadata (shape and label)
    try:
        binding_row = overlay_conn.execute(
            """
            SELECT b.shape_rows, b.shape_cols, sv.label
            FROM ir.agent_bindings b
            LEFT JOIN semantic_variables sv ON b.binding_id = sv.binding_id
            WHERE b.binding_id = ?
        """,
            (binding_id,),
        ).fetchone()
    except sqlite3.OperationalError:
        binding_row = overlay_conn.execute(
            """
            SELECT b.shape_rows, b.shape_cols, sv.label
            FROM ir.bindings b
            LEFT JOIN semantic_variables sv ON b.binding_id = sv.binding_id
            WHERE b.binding_id = ?
        """,
            (binding_id,),
        ).fetchone()

    if not binding_row:
        return None

    shape_rows, shape_cols, label = binding_row
    binding = {"shape_rows": shape_rows, "shape_cols": shape_cols}

    if entity_type == "time_series":
        axis = _infer_time_series_axis(binding, binding_id, time_axis_link)
        return [axis] if axis else None
    elif entity_type == "table":
        table_time_axis = time_axis_link.get(binding_id) if time_axis_link else None
        return _infer_table_axes(binding, label or "", table_time_axis)

    return None


def _get_entity_metadata(
    overlay_conn: sqlite3.Connection, binding_id: str
) -> dict[str, Any] | None:
    """
    Extract entity metadata from mutation that created this binding.

    For composite bindings created by Story 30-32, retrieves entity_type and axis information
    from the merge_bindings mutation metadata.

    VALIDATION (Story 20): Composite bindings marked as 'scalar' but spanning multiple cells
    are corrected to the appropriate entity_type based on shape.

    For non-composite bindings, infers entity_type from shape and formula presence.

    AXES INFERENCE (Story 21): For time_series and table entities, infers axes metadata
    from binding shape and label using heuristics.

    Args:
        overlay_conn: Connection to overlay database
        binding_id: Binding ID to query

    Returns:
        Entity metadata dict with entity_type, row_axis, column_axis, or None if not available
    """
    # Get label_source mutation for this binding
    row = overlay_conn.execute(
        """
        SELECT ml.parameters_json, ml.action
        FROM semantic_variables sv
        JOIN mutation_log ml ON sv.label_source = ml.mutation_id
        WHERE sv.binding_id = ?
    """,
        (binding_id,),
    ).fetchone()

    if not row or not row[0]:
        return None

    try:
        parameters = json.loads(row[0])
        action = row[1]

        # For merge_bindings mutations, entity metadata is in parameters.metadata
        if action == "merge_bindings":
            metadata = parameters.get("metadata", {})
        else:
            # For other mutations, no entity metadata expected
            return None

        entity_type = metadata.get("entity_type")

        if not entity_type:
            return None

        # VALIDATION (Story 20 + Story 25): Check if this is a composite binding with
        # contradictory entity_type metadata (scalar multi-cell OR time_series 2D)
        is_composite = (
            overlay_conn.execute(
                """
            SELECT 1 FROM composite_bindings WHERE composite_id = ? LIMIT 1
        """,
                (binding_id,),
            ).fetchone()
            is not None
        )

        if is_composite and entity_type in ("scalar", "time_series"):
            # Get shape to validate
            # For composite bindings, calculate shape from members
            try:
                member_rows = overlay_conn.execute(
                    """
                    SELECT cb.ir_binding_id, b.shape_rows, b.shape_cols, b.address
                    FROM composite_bindings cb
                    JOIN ir.agent_bindings b ON cb.ir_binding_id = b.binding_id
                    WHERE cb.composite_id = ?
                    ORDER BY cb.ordinal
                """,
                    (binding_id,),
                ).fetchall()
            except sqlite3.OperationalError:
                member_rows = overlay_conn.execute(
                    """
                    SELECT cb.ir_binding_id, b.shape_rows, b.shape_cols, b.address_a1
                    FROM composite_bindings cb
                    JOIN ir.bindings b ON cb.ir_binding_id = b.binding_id
                    WHERE cb.composite_id = ?
                    ORDER BY cb.ordinal
                """,
                    (binding_id,),
                ).fetchall()

            if member_rows:
                # Story 22 Fix: Calculate bounding box instead of max aggregation
                shape_rows, shape_cols = _calculate_composite_bounding_box(member_rows)

                # Check for contradictory metadata:
                # - scalar CANNOT be multi-cell (shape_rows > 1 OR shape_cols > 1)
                # - time_series CANNOT be 2D (shape_rows > 1 AND shape_cols > 1)
                needs_correction = False
                if (
                    entity_type == "scalar"
                    and (shape_rows > 1 or shape_cols > 1)
                    or entity_type == "time_series"
                    and shape_rows > 1
                    and shape_cols > 1
                ):
                    needs_correction = True

                if needs_correction:
                    # Get formula presence to infer correct type
                    try:
                        has_formula = (
                            overlay_conn.execute(
                                """
                            SELECT 1 FROM ir.cell_to_binding ctb
                            JOIN ir.agent_cells ac ON ctb.cell_id = ac.cell_id
                            WHERE ctb.binding_id IN (
                                SELECT ir_binding_id FROM composite_bindings
                                WHERE composite_id = ?
                            ) AND ac.formula IS NOT NULL AND ac.formula != ''
                            LIMIT 1
                        """,
                                (binding_id,),
                            ).fetchone()
                            is not None
                        )
                    except sqlite3.OperationalError:
                        has_formula = (
                            overlay_conn.execute(
                                """
                            SELECT 1
                            FROM ir.cells c
                            WHERE c.binding_id IN (
                                SELECT ir_binding_id FROM composite_bindings
                                WHERE composite_id = ?
                            ) AND c.formula_a1 IS NOT NULL AND c.formula_a1 != ''
                            LIMIT 1
                        """,
                                (binding_id,),
                            ).fetchone()
                            is not None
                        )

                    # Correct the entity_type
                    corrected_type = _infer_entity_type_from_shape(
                        shape_rows, shape_cols, has_formula
                    )
                    logger.warning(
                        f"Composite binding {binding_id} marked as {entity_type} but spans "
                        f"{shape_rows}x{shape_cols} - corrected to {corrected_type}"
                    )
                    entity_type = corrected_type

        result: dict[str, Any] = {"entity_type": entity_type}

        # Phase B2 load-bearing surfacing fix: prefer the SHAPE-AWARE classifier's
        # member-accurate geometry_axes over bounding-box re-derivation (see
        # _axes_from_geometry). This path has no time_axis_link in scope, so the time
        # axis is named in the batched extract_variable_data path instead.
        geometry_axes = metadata.get("geometry_axes")
        if geometry_axes:
            axes = _axes_from_geometry(geometry_axes, metadata.get("geometry_kind", ""), None)
            if axes:
                result["axes"] = axes
        else:
            # AXES INFERENCE (Story 21): Infer axes for time_series and table entities
            # NOTE: This replaces the old table axes format (object with row/column)
            # with the new list-based format (array of axis objects)
            inferred_axes = _infer_axes_from_binding(overlay_conn, binding_id, entity_type)
            if inferred_axes:
                result["axes"] = inferred_axes

        return result

    except json.JSONDecodeError:
        logger.warning(f"Failed to parse metadata_json for binding {binding_id}")
        return None


def extract_variable_data(overlay_conn: sqlite3.Connection) -> list[dict[str, Any]]:
    """
    Extract variable data from overlay + IR join.

    Args:
        overlay_conn: Connection to overlay (must have IR attached via attach_ir_to_overlay)

    Returns:
        List of variable dicts with variable_id, label, excel_location, formula,
        snapshot_value, label_confidence, classification_confidence, dependencies,
        provenance (knowledge_source, enrichment metadata)
    """
    start_time = time.perf_counter()
    # Get workbook name from IR metadata table (try fast schema first, fall back to legacy)
    try:
        meta_row = overlay_conn.execute(
            "SELECT value FROM ir.ir_metadata WHERE key = 'workbook_sha256' LIMIT 1"
        ).fetchone()
        workbook_name = meta_row[0] if meta_row else "unknown.xlsx"
    except sqlite3.Error:
        try:
            meta_row = overlay_conn.execute("SELECT file_path FROM ir.meta LIMIT 1").fetchone()
            workbook_name = meta_row[0] if meta_row else "unknown.xlsx"
        except sqlite3.Error as e:
            logger.warning(f"Failed to query IR meta for workbook name: {e}")
            workbook_name = "unknown.xlsx"

    # Query joining overlay and IR
    # Get labelled bindings with sheet/address/formula from IR (value is in cells)
    # And confidence + provenance from mutation_log
    # Filter for active bindings only (exclude superseded bindings from merges/splits)
    try:
        rows = overlay_conn.execute("""
            SELECT
                sv.binding_id,
                sv.label,
                sv.actuarial_class,
                sv.reconciliation_required,
                sv.reconciliation_rationale,
                b.sheet,
                b.address,
                sv.is_composite,
                sv.label_confidence,
                sv.classification_confidence,
                ml.action,
                ml.parameters_json,
                ml.metadata_json
            FROM semantic_variables sv
            LEFT JOIN ir.agent_bindings b ON sv.binding_id = b.binding_id
            LEFT JOIN mutation_log ml ON sv.label_source = ml.mutation_id
            WHERE sv.is_active = 1
            ORDER BY b.sheet, b.address
        """).fetchall()
    except sqlite3.OperationalError:
        rows = overlay_conn.execute("""
            SELECT
                sv.binding_id,
                sv.label,
                sv.actuarial_class,
                sv.reconciliation_required,
                sv.reconciliation_rationale,
                b.sheet,
                b.address_a1,
                sv.is_composite,
                sv.label_confidence,
                sv.classification_confidence,
                ml.action,
                ml.parameters_json,
                ml.metadata_json
            FROM semantic_variables sv
            LEFT JOIN ir.bindings b ON sv.binding_id = b.binding_id
            LEFT JOIN mutation_log ml ON sv.label_source = ml.mutation_id
            WHERE sv.is_active = 1
            ORDER BY b.sheet, b.address_a1
        """).fetchall()
    _log_timing("extract_variable_rows", time.perf_counter() - start_time, f"rows={len(rows)}")

    composite_members = _load_composite_members(overlay_conn)
    composite_member_ids: dict[str, list[str]] = {}
    composite_lookup: dict[str, str] = {}
    for composite_id, members in composite_members.items():
        member_ids = [member[0] for member in members]
        composite_member_ids[composite_id] = member_ids
        for ir_binding_id in member_ids:
            composite_lookup[ir_binding_id] = composite_id

    # Build variable map for semantic formulas (Sheet!Address -> Label)
    # For ranges, expand to individual cells so formulas can reference any cell
    map_start = time.perf_counter()
    variable_map = {}

    # Pre-process rows to handle composite bindings for variable_map
    processed_rows = []

    for row in rows:
        (
            binding_id,
            label,
            actuarial_class,
            recon_required,
            recon_rationale,
            sheet,
            address,
            is_composite,
            label_conf,
            class_conf,
            action,
            parameters_json,
            metadata_json,
        ) = row

        # Handle Composite Binding Resolution
        if is_composite:
            # Fetch members
            members = composite_members.get(binding_id, [])

            if members:
                # Use sheet from first member
                sheet = members[0][1]

                # Calculate union address
                ranges = [m[2].split("!")[-1] if "!" in m[2] else m[2] for m in members]

                # Calculate bounding box
                min_col, min_row = 999999, 999999
                max_col, max_row = -1, -1

                valid_ranges = False
                for r in ranges:
                    try:
                        # Normalize: strip $ signs for parsing
                        r_clean = r.replace("$", "")

                        if ":" in r_clean:
                            start, end = r_clean.split(":")
                            s_pos = _parse_a1_cell(start)
                            e_pos = _parse_a1_cell(end)
                            if s_pos and e_pos:
                                min_col = min(min_col, s_pos[0], e_pos[0])
                                min_row = min(min_row, s_pos[1], e_pos[1])
                                max_col = max(max_col, s_pos[0], e_pos[0])
                                max_row = max(max_row, s_pos[1], e_pos[1])
                                valid_ranges = True
                        else:
                            pos = _parse_a1_cell(r_clean)
                            if pos:
                                min_col = min(min_col, pos[0])
                                min_row = min(min_row, pos[1])
                                max_col = max(max_col, pos[0])
                                max_row = max(max_row, pos[1])
                                valid_ranges = True
                    except Exception:
                        continue

                if valid_ranges:
                    start_addr = f"{_col_index_to_letters(min_col)}{min_row + 1}"
                    end_addr = f"{_col_index_to_letters(max_col)}{max_row + 1}"
                    address = f"{start_addr}:{end_addr}" if start_addr != end_addr else start_addr
                else:
                    address = ranges[0] if ranges else "A1"
            else:
                sheet = "Unknown"
                address = "A1"

        processed_rows.append(
            (
                binding_id,
                label,
                actuarial_class,
                recon_required,
                recon_rationale,
                sheet,
                address,
                is_composite,
                label_conf,
                class_conf,
                action,
                parameters_json,
                metadata_json,
            )
        )

        if sheet and address and label:
            try:
                # address_a1 from bindings table includes sheet prefix (e.g., "Sheet1!A1:A10")
                # Strip the sheet prefix to get pure A1 notation
                # Split on last ! to handle sheets with quotes
                pure_address = address.split("!")[-1] if "!" in address else address

                # Expand range into individual cells
                # For "A1:A10", this creates entries for A1, A2, ..., A10
                # PERFORMANCE: Use smaller max_cells limit (1000) to keep variable_map size reasonable
                individual_cells = _expand_a1_range(pure_address, max_cells=1000)

                if not individual_cells:
                    # Range too large - store as range reference only
                    key = f"{sheet}!{pure_address.replace('$', '')}"
                    variable_map[key] = label
                else:
                    for cell in individual_cells:
                        # Store canonical key: Sheet!Address (no quotes, no $)
                        key = f"{sheet}!{cell}"
                        variable_map[key] = label

            except ValueError as e:
                # If range expansion fails, fall back to storing the range as-is
                logger.warning(f"Failed to expand range {address} for label '{label}': {e}")
                # Strip sheet prefix from address for fallback too
                pure_address = address.split("!")[-1] if "!" in address else address
                key = f"{sheet}!{pure_address.replace('$', '')}"
                variable_map[key] = label

    # Build binding_ranges for O(m) range matching
    binding_ranges = _build_binding_ranges(processed_rows)
    logger.info(
        f"Built variable_map with {len(variable_map)} entries for {len(processed_rows)} bindings"
    )
    logger.info(f"Built binding_ranges with {len(binding_ranges)} entries")
    _log_timing(
        "build_variable_maps",
        time.perf_counter() - map_start,
        f"bindings={len(processed_rows)} map={len(variable_map)} ranges={len(binding_ranges)}",
    )

    binding_ids_for_cells: set[str] = set()
    for binding_id, _, _, _, _, _, _, is_composite, _, _, _, _, _ in processed_rows:
        if is_composite:
            member_ids = composite_member_ids.get(binding_id)
            if member_ids:
                binding_ids_for_cells.update(member_ids)
            else:
                binding_ids_for_cells.add(binding_id)
        else:
            binding_ids_for_cells.add(binding_id)

    binding_id_list = sorted(binding_ids_for_cells)
    cells_start = time.perf_counter()
    binding_cells_by_id = _load_cells_by_binding(overlay_conn, binding_id_list)
    binding_shapes = _load_binding_shapes(overlay_conn, binding_id_list)
    dependency_parents, dependency_children = _build_dependency_maps(overlay_conn)
    # key_associations retired (Phase A); the series->axis index link is surfaced
    # in Phase B1 via the time axis on each variable's `axes` (see time_axis_link
    # below), not via `indices`, so `indices` stays empty here.
    indices_by_binding: dict[str, list[dict[str, Any]]] = {}
    # Phase B1: the real time-axis link, read from ir.binding_time_annotations and
    # gated by is_time_dependent=1 AND MAX(shape)>1 so 1x1 scalars never get a
    # t-index. Keyed by IR binding_id.
    time_axis_link = _build_time_axis_link(overlay_conn)
    has_formula_by_binding = {
        binding_id: any(cell[2] for cell in cells)
        for binding_id, cells in binding_cells_by_id.items()
    }
    composite_shapes: dict[str, tuple[int, int]] = {}
    for composite_id, member_ids in composite_member_ids.items():
        member_rows = []
        for member_id in member_ids:
            shape_info = binding_shapes.get(member_id)
            if shape_info:
                shape_rows, shape_cols, member_address = shape_info
                member_rows.append((member_id, shape_rows, shape_cols, member_address))
        if member_rows:
            composite_shapes[composite_id] = _calculate_composite_bounding_box(member_rows)
    entity_metadata_by_binding: dict[str, dict[str, Any]] = {}
    for (
        binding_id,
        label,
        _actuarial_class,
        _recon_required,
        _recon_rationale,
        _sheet,
        _address,
        is_composite,
        _label_conf,
        _class_conf,
        action,
        parameters_json,
        _metadata_json,
    ) in processed_rows:
        if action != "merge_bindings" or not parameters_json:
            continue
        try:
            parameters = json.loads(parameters_json)
        except json.JSONDecodeError:
            logger.warning(f"Failed to parse parameters_json for binding {binding_id}")
            continue

        metadata = parameters.get("metadata", {})
        entity_type = metadata.get("entity_type")
        if not entity_type:
            continue

        shape_rows = None
        shape_cols = None
        if is_composite:
            shape = composite_shapes.get(binding_id)
            if shape:
                shape_rows, shape_cols = shape
                if entity_type in ("scalar", "time_series"):
                    needs_correction = False
                    if (
                        entity_type == "scalar"
                        and (shape_rows > 1 or shape_cols > 1)
                        or entity_type == "time_series"
                        and shape_rows > 1
                        and shape_cols > 1
                    ):
                        needs_correction = True

                    if needs_correction:
                        member_ids = composite_member_ids.get(binding_id, [])
                        has_formula = any(
                            has_formula_by_binding.get(member_id) for member_id in member_ids
                        )
                        corrected_type = _infer_entity_type_from_shape(
                            shape_rows,
                            shape_cols,
                            has_formula,
                        )
                        logger.warning(
                            f"Composite binding {binding_id} marked as {entity_type} but spans "
                            f"{shape_rows}x{shape_cols} - corrected to {corrected_type}"
                        )
                        entity_type = corrected_type
        else:
            shape_info = binding_shapes.get(binding_id)
            if shape_info:
                shape_rows, shape_cols = shape_info[0], shape_info[1]

        result: dict[str, Any] = {"entity_type": entity_type}

        # Phase B2 load-bearing surfacing fix: if the SHAPE-AWARE classifier recorded
        # member-accurate geometry_axes for this collapsed family, surface THOSE
        # instead of re-deriving axes from the union bounding box (which inflates
        # sparse/strided families into dense false objects). Bounding-box-derived
        # axes are only used as a fallback for composites that predate B2 metadata.
        geometry_axes = metadata.get("geometry_axes")
        if geometry_axes:
            time_axis = time_axis_link.get(binding_id) if time_axis_link else None
            axes = _axes_from_geometry(geometry_axes, metadata.get("geometry_kind", ""), time_axis)
            if axes:
                result["axes"] = axes
        elif shape_rows is not None and shape_cols is not None:
            axes = _infer_axes_from_shape(
                shape_rows,
                shape_cols,
                label or "",
                entity_type,
                binding_id=binding_id,
                time_axis_link=time_axis_link,
            )
            if axes:
                result["axes"] = axes

        entity_metadata_by_binding[binding_id] = result
    _log_timing(
        "load_cells_and_dependencies",
        time.perf_counter() - cells_start,
        f"bindings={len(binding_id_list)}",
    )

    semantic_formula_cache: dict[tuple[str, str], str] = {}
    explanation_cache: dict[tuple[str, str, str], str | None] = {}

    def _get_semantic_formula_cached(formula: str, sheet_name: str | None) -> str:
        key = (formula, sheet_name or "")
        cached = semantic_formula_cache.get(key)
        if cached is not None:
            return cached

        result = generate_semantic_formula(
            formula,
            variable_map,
            binding_ranges,
            sheet_name,
        )
        semantic_formula_cache[key] = result
        return result

    variables = []
    vars_start = time.perf_counter()
    for (
        binding_id,
        label,
        actuarial_class,
        recon_required,
        recon_rationale,
        sheet,
        address,
        is_composite,
        label_conf_from_db,
        class_conf_from_db,
        _action,
        _parameters_json,
        metadata_json,
    ) in processed_rows:
        # Get ALL cells for this binding to properly handle ranges
        # Use evaluated_value (from dual-load) instead of value_snapshot

        if is_composite:
            member_ids = composite_member_ids.get(binding_id, [])
            if member_ids:
                all_cells_with_addr = []
                for member_id in member_ids:
                    all_cells_with_addr.extend(binding_cells_by_id.get(member_id, []))
                if len(all_cells_with_addr) > 1:
                    all_cells_with_addr.sort(key=lambda cell: cell[0])
            else:
                all_cells_with_addr = []
        else:
            all_cells_with_addr = binding_cells_by_id.get(binding_id, [])

        all_cells = []
        formula = None
        for _address, value, formula_a1 in all_cells_with_addr:
            all_cells.append((value, formula_a1))
            if formula is None and formula_a1:
                formula = formula_a1

        # Generate semantic formula
        semantic_formula = None
        if formula:
            semantic_formula = _get_semantic_formula_cached(formula, sheet)

            # Diagnostic: Detect binding ID patterns in semantic formulas (Story 23)
            # Pattern: Sheet::Cell::Cell (e.g., "Calculations::L999::L1007")
            if "::" in semantic_formula:
                logger.warning(
                    f"Binding ID pattern detected in semantic formula for {binding_id} (label: {label}): "
                    f"'{semantic_formula}'. This indicates a labelling fallback issue."
                )

        # Generate natural language explanation
        explanation = None
        if formula and semantic_formula:
            cache_key = (formula, semantic_formula, sheet or "")
            if cache_key in explanation_cache:
                explanation = explanation_cache[cache_key]
            else:
                try:
                    explanation = generate_explanation(
                        formula, semantic_formula, variable_map, sheet
                    )
                except Exception as e:
                    logger.warning(f"Failed to generate formula description for {binding_id}: {e}")
                    explanation = None
                explanation_cache[cache_key] = explanation

        # Extract array formula metadata if applicable
        array_formula_metadata = None
        try:
            array_formula_metadata = _extract_array_formula_metadata(
                binding_id,
                overlay_conn,
                variable_map,
                binding_ranges,
                sheet,
                cells=all_cells,
                semantic_formula_builder=_get_semantic_formula_cached,
            )
        except Exception as e:
            logger.warning(f"Failed to extract array formula metadata for {binding_id}: {e}")

        # Use confidence scores from database (written by Pass 4)
        # If not available (None), default to 0.0
        label_confidence = label_conf_from_db if label_conf_from_db is not None else 0.0
        classification_confidence = class_conf_from_db if class_conf_from_db is not None else None

        # Extract snapshot_value using evaluated_value (calculated results)
        snapshot_value = _extract_range_values(all_cells)

        # Get dependencies (parents and children)
        dependencies = _get_dependencies_from_maps(
            binding_id,
            is_composite,
            composite_member_ids,
            dependency_parents,
            dependency_children,
            composite_lookup,
        )

        # Get indices (for Assumptions/Inputs that have Index Lookup keys)
        indices = None
        if actuarial_class in ("Assumption", "Input"):
            index_list = indices_by_binding.get(binding_id)
            if index_list:
                indices = index_list

        # Get entity metadata (for tables, record sets, vectors, scalars)
        entity_metadata = entity_metadata_by_binding.get(binding_id)
        entity_type = entity_metadata.get("entity_type") if entity_metadata else None
        axes = entity_metadata.get("axes") if entity_metadata else None

        # If no entity_type from metadata, infer from shape
        if not entity_type:
            # For composite bindings, calculate shape from members
            if is_composite:
                shape = composite_shapes.get(binding_id)
                if shape:
                    shape_rows, shape_cols = shape
                    has_formula = bool(formula)
                    entity_type = _infer_entity_type_from_shape(shape_rows, shape_cols, has_formula)
            else:
                # Regular binding - get shape from IR bindings table
                shape_row = binding_shapes.get(binding_id)
                if shape_row:
                    shape_rows, shape_cols = shape_row[0], shape_row[1]
                    has_formula = bool(formula)  # Use formula extracted earlier
                    entity_type = _infer_entity_type_from_shape(shape_rows, shape_cols, has_formula)

        # AXES INFERENCE (Story 21): If no axes from metadata, infer from entity_type
        # This runs for ALL bindings (composite and regular) after entity_type is determined
        if not axes and entity_type:
            if is_composite:
                shape = composite_shapes.get(binding_id)
                if shape:
                    shape_rows, shape_cols = shape
                    inferred_axes = _infer_axes_from_shape(
                        shape_rows,
                        shape_cols,
                        label or "",
                        entity_type,
                        binding_id=binding_id,
                        time_axis_link=time_axis_link,
                    )
                    if inferred_axes:
                        axes = inferred_axes
            else:
                shape_row = binding_shapes.get(binding_id)
                if shape_row:
                    shape_rows, shape_cols = shape_row[0], shape_row[1]
                    inferred_axes = _infer_axes_from_shape(
                        shape_rows,
                        shape_cols,
                        label or "",
                        entity_type,
                        binding_id=binding_id,
                        time_axis_link=time_axis_link,
                    )
                    if inferred_axes:
                        axes = inferred_axes

        # Strip sheet prefix from address if present (address_a1 from IR includes sheet)
        # Templates expect: {"sheet": "Projection", "address": "A1:O7"}
        # Not: {"sheet": "Projection", "address": "Projection!A1:O7"}
        pure_address = address
        if address and "!" in address:
            # Handle quoted sheet names like "'Sheet Name'!A1"
            pure_address = address.split("!")[-1]

        if not label:
            label = f"[Unlabeled] {pure_address}"
        elif label_confidence == 0.0 and _looks_like_raw_address(label):
            # A 0.0-confidence raw cell address (e.g. "Calculations!D110") is the
            # labeller's last-resort fallback, not a name. Don't render it verbatim.
            label = f"[Unlabeled] {pure_address}"

        # Extract provenance from mutation metadata (Sprint 7)
        provenance = None
        if metadata_json:
            try:
                metadata = (
                    json.loads(metadata_json) if isinstance(metadata_json, str) else metadata_json
                )
                # Extract enrichment provenance fields
                knowledge_source = metadata.get("knowledge_source")
                confidence_initial = metadata.get("confidence_initial")
                sprint = metadata.get("sprint")

                if knowledge_source or confidence_initial is not None or sprint:
                    provenance = {}
                    if knowledge_source:
                        provenance["knowledge_source"] = knowledge_source
                    if confidence_initial is not None:
                        provenance["confidence_initial"] = confidence_initial
                    if sprint:
                        provenance["sprint"] = sprint
                    # Add reasoning if available
                    if metadata.get("reasoning"):
                        provenance["reasoning"] = metadata["reasoning"]
            except (json.JSONDecodeError, TypeError) as e:
                logger.warning(f"Failed to parse metadata_json for binding {binding_id}: {e}")

        variable = {
            "variable_id": binding_id,
            "label": label,
            "actuarial_class": actuarial_class,
            "entity_type": entity_type,
            "axes": axes,
            "reconciliation_required": bool(recon_required) if recon_required is not None else None,
            "reconciliation_rationale": recon_rationale,
            "excel_location": {
                "sheet": sheet,
                "address": pure_address,  # Use sheet-stripped address
                "workbook": workbook_name,
            },
            "formula": formula if formula else f"={address}",
            "semantic_formula": semantic_formula,
            "explanation": explanation,
            "array_formula_metadata": array_formula_metadata,
            "snapshot_value": snapshot_value,
            "label_confidence": label_confidence,
            "classification_confidence": classification_confidence,
            "dependencies": dependencies,
            "indices": indices,
            "provenance": provenance,  # Sprint 7: enrichment provenance
        }

        # Filter Index Lookup variables (documented as dimensions of their parent)
        # Garbage variables are already filtered via disable_binding mutations
        if actuarial_class != "Index Lookup":
            variables.append(variable)

    # NOTE: No post-processing needed - garbage filtering and fragment merging
    # are now handled by CleanupRefiner via mutations before this point

    _log_timing("build_variables", time.perf_counter() - vars_start, f"variables={len(variables)}")
    _log_timing("extract_variable_data_total", time.perf_counter() - start_time)
    return variables


def validate_json_spec(spec: dict[str, Any]) -> list[str]:
    """
    Validate JSON spec against schema.

    Args:
        spec: Model spec dict to validate

    Returns:
        List of validation errors (empty if valid)
    """
    if jsonschema is None:
        logger.warning("jsonschema not available; skipping spec validation")
        return []

    # Load schema lazily to avoid module-level side effects
    try:
        with open(SCHEMA_PATH, encoding="utf-8") as f:
            schema = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        return [f"Failed to load schema from {SCHEMA_PATH}: {e}"]

    errors = []
    try:
        jsonschema.validate(spec, schema)
    except jsonschema.ValidationError as e:
        errors.append(f"Validation error: {e.message}")
    except jsonschema.SchemaError as e:
        errors.append(f"Schema error: {e.message}")

    return errors


def generate_json_spec(overlay_db_path: str, ir_db_path: str, output_path: str) -> None:
    """
    Generate JSON model specification from overlay database.

    Args:
        overlay_db_path: Path to semantic_overlay.db
        ir_db_path: Path to Phase 1 IR database
        output_path: Where to write model_spec.json

    Raises:
        ValueError: If validation fails
    """
    start_total = time.perf_counter()
    # Connect to databases. overlay_conn needs uri=True so attach_ir_to_overlay()'s
    # `ATTACH DATABASE 'file:...?mode=ro'` below is parsed as a URI on this connection.
    overlay_conn = sqlite3.connect(overlay_db_path, uri=True)
    ir_conn = connect_read_only(ir_db_path)

    try:
        # Attach IR to overlay for joins
        from xl_marinade.core.labelling.overlay_database import attach_ir_to_overlay

        attach_start = time.perf_counter()
        attach_ir_to_overlay(overlay_conn, ir_db_path)
        _log_timing("attach_ir", time.perf_counter() - attach_start)

        # Build spec
        vars_start = time.perf_counter()
        variables = extract_variable_data(overlay_conn)
        _log_timing(
            "extract_variable_data", time.perf_counter() - vars_start, f"variables={len(variables)}"
        )

        # Generate reconciliation scope summary
        requiring_recon = [v["variable_id"] for v in variables if v.get("reconciliation_required")]
        skipped_recon = [
            v["variable_id"] for v in variables if v.get("reconciliation_required") is False
        ]

        recon_rationale = (
            f"{len(requiring_recon)} of {len(variables)} variables require "
            "reconciliation testing. Key actuarial outputs and calculations "
            "requiring validation in re-implementation. Helpers and formatting excluded."
        )

        reconciliation_scope = {
            "variables_requiring_reconciliation": requiring_recon,
            "variables_skipped": skipped_recon,
            "rationale": recon_rationale,
        }

        # Generate orphans summary (Story 15)
        orphan_rows = overlay_conn.execute("""
            SELECT binding_id FROM semantic_variables WHERE is_orphan = 1
        """).fetchall()
        orphan_binding_ids = [row[0] for row in orphan_rows]

        total_bindings_row = overlay_conn.execute("""
            SELECT COUNT(*) FROM semantic_variables
        """).fetchone()
        total_bindings = total_bindings_row[0] if total_bindings_row else 0

        orphan_count = len(orphan_binding_ids)
        orphan_percentage = (orphan_count / total_bindings * 100) if total_bindings > 0 else 0

        orphans_rationale = (
            f"{orphan_count} of {total_bindings} bindings ({orphan_percentage:.1f}%) are orphaned "
            f"(not reachable from any root cell). Orphaned bindings represent unused inputs, "
            f"deprecated calculations, or dead code in the model. They are excluded from documentation "
            f"but preserved in the IR for audit purposes."
        )

        orphans = {
            "orphan_binding_ids": orphan_binding_ids,
            "orphan_count": orphan_count,
            "total_bindings": total_bindings,
            "rationale": orphans_rationale,
        }

        meta_start = time.perf_counter()
        metadata = create_metadata(overlay_conn, ir_conn)
        _log_timing("create_metadata", time.perf_counter() - meta_start)

        spec = {
            "metadata": metadata,
            "reconciliation_scope": reconciliation_scope,
            "orphans": orphans,
            "variables": variables,
        }

        # Validate
        validate_start = time.perf_counter()
        errors = validate_json_spec(spec)
        _log_timing("validate_spec", time.perf_counter() - validate_start)
        if errors:
            raise ValueError(f"Spec validation failed: {errors}")

        # Write to file with pretty formatting and trailing newline
        write_start = time.perf_counter()
        with open(output_path, "w", encoding="utf-8", newline="\n") as f:
            json.dump(spec, f, indent=2)
            f.write("\n")
        _log_timing("write_spec_file", time.perf_counter() - write_start)

    finally:
        overlay_conn.close()
        ir_conn.close()
        _log_timing("total", time.perf_counter() - start_total)
