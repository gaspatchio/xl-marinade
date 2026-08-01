# ABOUTME: OFFSET-formula edge synthesiser (R21 Fix D). Static analysis of OFFSET
# ABOUTME: function calls in cell formulas emits `via_offset_static` and
# ABOUTME: `via_offset_volatile` binding_edges from the formula-bearing cell to
# ABOUTME: the OFFSET anchor's column/range so the dependency walker can reach
# ABOUTME: through OFFSET-mediated bridges (R20 Risk Drivers→Calculation Engine).

"""OFFSET-formula edge synthesiser (R21 Fix D).

Background (per `r20-multifactor-investigation-concept-tagger-coverage-gap-2026-04-29.md`
and `r21-multifactor-investigation-vba-paste-edge-synthesis-2026-04-29.md` §7):

Actuarial / projection workbooks frequently bridge between an output sheet
(e.g. `Risk Drivers`) and a calculation core (e.g. `Calculation Engine`)
via OFFSET formulas of the form

    OFFSET('Calculation Engine'!$AN$8, row_expr, col_expr, height, width)

Risk Drivers row 4 carries 164 such OFFSET refs, all anchored at constant
cells in Calculation Engine. The existing `ReferenceExtractor` treats
OFFSET as a fully-opaque dynamic function — it emits one `external` edge
with a `DYNAMIC:OFFSET(...)` marker and extracts no precedents from
inside the call. Consequence: the formula DAG cannot reach from Risk
Drivers back to Calculation Engine through OFFSET-mediated dependencies.

This module's job is the *parser layer*: walk a formula AST, find every
OFFSET call, and classify each as one of three resolution kinds:

- **static**: anchor is a literal Ref AND row/col offsets (and any
  explicit height/width) are numeric constants. The exact target
  (anchor + offsets, dimensions = height/width or anchor's dimensions)
  is computed.

- **volatile**: anchor is a literal Ref but row OR col offset is a
  formula (Binary, Function, Ref). We CANNOT compute the exact target,
  but we know the result lies within the anchor sheet. The classifier
  records whichever offsets/dimensions ARE known so the IR resolver
  can emit a coarse range (e.g. anchor's column from anchor row down
  to the sheet's max-populated row when both row offset and dimensions
  are formula-driven). Better to over-include cells than to omit the
  dependency entirely.

- **skipped**: anchor is INDIRECT, a nested OFFSET, or a non-Ref
  expression (function call, identifier). No edge is synthesisable
  because the anchor itself is dynamic. Marked with a typed
  `skipped_reason`.

The parser is **typed-field, not phrase-list-driven**: it operates
purely on the syntactic shape of the OFFSET call (Ref vs Const vs other),
not on workbook-specific cell addresses, sheet names, or function-name
heuristics. The same parser handles a lifetime-policy mortality
projection workbook and a regulatory-capital simulation workbook with
no per-domain configuration.

The companion driver `synthesize_offset_edges(conn)` consumes these
`OffsetCall` records, resolves the anchor to a sheet_id + (r1, c1, r2,
c2) rectangle on the IR's `cells` table, and emits
`binding_edges (kind='via_offset_static' | 'via_offset_volatile')`
rows. It runs as a post-pass after binding edges are wired (so
target bindings exist) and only inserts new edges — it never modifies
or removes existing rows.
"""

from __future__ import annotations

import logging
import sqlite3
from dataclasses import dataclass
from typing import Any

logger = logging.getLogger(__name__)


# --- Data classes ----------------------------------------------------------


@dataclass
class OffsetCall:
    """A single OFFSET(...) call extracted from a formula AST.

    `kind` is one of:
    - 'static'   — anchor is a Ref AND row/col offsets and any explicit
                   dimensions are numeric constants; the target is an
                   exact (row, col, height, width) tuple.
    - 'volatile' — anchor is a Ref but row OR col offset (or an explicit
                   dimension) is non-constant; the target is a coarse
                   range using the constants we know.
    - 'skipped'  — anchor is INDIRECT, a nested OFFSET, or a non-Ref
                   expression; no edge is synthesisable.

    For 'static' and 'volatile' kinds, `target_sheet`/`target_row`/`target_col`
    describe the anchor-relative resolved target. For 'static',
    `target_height` and `target_width` are always concrete integers >= 1.
    For 'volatile', they may be None — the resolver will substitute the
    anchor sheet's data extent.

    `coarsening_note` (volatile only) records *why* the range was
    coarsened (which arguments were formula-driven), so the implementation
    report can audit the synthesiser's behaviour on real workbooks.
    """

    kind: str  # 'static' | 'volatile' | 'skipped'
    anchor_ref: str
    target_sheet: str | None = None
    target_row: int | None = None
    target_col: int | None = None
    target_height: int | None = None
    target_width: int | None = None
    skipped_reason: str | None = None
    coarsening_note: str | None = None


# --- AST-walking parser -----------------------------------------------------


def _is_const_int(node: Any) -> int | None:
    """Return the integer value of a numeric Const, or None if not constant.

    Excel formula parsers wrap unary `-` around the Const for negative
    numbers (e.g. `OFFSET(A1, -2, 0)` → `Unary(-, Const(2))`). We unwrap
    a single `-` (or `+`) Unary and resolve to the signed integer.
    Anything more elaborate (Binary, Function, Ref) is treated as
    non-constant.
    """
    if not isinstance(node, dict):
        return None
    if node.get("type") == "Unary":
        op = node.get("operator")
        operand = node.get("operand")
        inner = _is_const_int(operand)
        if inner is None:
            return None
        if op == "-":
            return -inner
        if op == "+":
            return inner
        return None
    if node.get("type") != "Const":
        return None
    v = node.get("value")
    if isinstance(v, bool):
        # bools are Const but not what OFFSET expects as offsets
        return None
    if isinstance(v, (int, float)):
        try:
            return int(v)
        except (ValueError, OverflowError):
            return None
    return None


def _classify_anchor(node: Any) -> tuple[str, str | None]:
    """Classify the OFFSET anchor (first argument).

    Returns `(kind, ref)` where:
    - kind='ref'      — anchor is a literal Ref; ref carries the address text
    - kind='indirect' — anchor is INDIRECT(...) (out of scope)
    - kind='offset'   — anchor is a nested OFFSET(...) (out of scope)
    - kind='nonref'   — anchor is some other expression (out of scope)
    """
    if not isinstance(node, dict):
        return ("nonref", None)
    t = node.get("type")
    if t == "Ref":
        return ("ref", node.get("ref"))
    if t == "Function":
        name = (node.get("name") or "").upper()
        if name == "INDIRECT":
            return ("indirect", None)
        if name == "OFFSET":
            return ("offset", None)
    return ("nonref", None)


def _walk_collect_offsets(node: Any, out: list[dict[str, Any]]) -> None:
    """Walk an AST collecting every OFFSET function call.

    Outer OFFSETs are kept; inner OFFSETs nested as the *anchor* of an
    outer OFFSET are NOT collected (the outer's `skipped_reason='nested_anchor'`
    handling already records them). Inner OFFSETs that appear elsewhere
    (e.g. as a sibling argument inside SUMPRODUCT) ARE collected.
    """
    if not isinstance(node, dict):
        return

    if node.get("type") == "Function" and (node.get("name") or "").upper() == "OFFSET":
        out.append(node)
        # Recurse into args 1..n (skip args[0] = anchor, since a nested
        # OFFSET there is handled by the outer's skip logic; recursing
        # would double-count it).
        args = node.get("args") or []
        for arg in args[1:]:
            _walk_collect_offsets(arg, out)
        return

    # For non-OFFSET nodes, walk all children.
    for key in ("left", "right", "operand"):
        child = node.get(key)
        if child is not None:
            _walk_collect_offsets(child, out)
    for arg in node.get("args") or []:
        _walk_collect_offsets(arg, out)


def parse_offset_calls(formula: str, current_sheet: str) -> list[OffsetCall]:
    """Parse all OFFSET(...) calls in a formula.

    Args:
        formula: A1-form formula text (with or without leading '=').
        current_sheet: Name of the sheet the formula lives on (used as
            the default qualifier for anchor refs that omit a sheet
            name).

    Returns:
        List of OffsetCall — one per OFFSET found. Calls whose anchor
        is INDIRECT, a nested OFFSET, or a non-Ref expression have
        `kind='skipped'` with `skipped_reason` set.
    """
    if not formula:
        return []

    # Defer the import to avoid circulars at module-load time.
    from xl_marinade.core.parser import parse_formula
    from xl_marinade.core.ref_converter import parse_cell_address

    try:
        ast = parse_formula(formula)
    except (ValueError, SyntaxError, KeyError, AttributeError):
        return []

    nodes: list[dict[str, Any]] = []
    _walk_collect_offsets(ast, nodes)

    results: list[OffsetCall] = []
    for fn in nodes:
        args = fn.get("args") or []
        if len(args) < 3:
            results.append(
                OffsetCall(
                    kind="skipped",
                    anchor_ref="",
                    skipped_reason="invalid_arity",
                )
            )
            continue

        anchor_kind, anchor_ref_text = _classify_anchor(args[0])
        if anchor_kind != "ref" or not anchor_ref_text:
            results.append(
                OffsetCall(
                    kind="skipped",
                    anchor_ref=anchor_ref_text or "",
                    skipped_reason=(
                        "indirect_anchor"
                        if anchor_kind == "indirect"
                        else "nested_anchor"
                        if anchor_kind == "offset"
                        else "nonref_anchor"
                    ),
                )
            )
            continue

        # Parse the anchor ref to get sheet, row, col, and (for range anchors)
        # the embedded height/width.
        parsed = parse_cell_address(anchor_ref_text)
        anchor_row = parsed.get("row", 0)
        anchor_col = parsed.get("col", 0)
        if not isinstance(anchor_row, int) or not isinstance(anchor_col, int):
            results.append(
                OffsetCall(
                    kind="skipped",
                    anchor_ref=anchor_ref_text,
                    skipped_reason="anchor_unparseable",
                )
            )
            continue
        if anchor_row < 1 or anchor_col < 1:
            results.append(
                OffsetCall(
                    kind="skipped",
                    anchor_ref=anchor_ref_text,
                    skipped_reason="anchor_unparseable",
                )
            )
            continue

        anchor_sheet_raw = parsed.get("sheet", "")
        anchor_sheet = anchor_sheet_raw if isinstance(anchor_sheet_raw, str) else ""
        if not anchor_sheet:
            anchor_sheet = current_sheet
        anchor_height_raw = parsed.get("height", 1)
        anchor_height = anchor_height_raw if isinstance(anchor_height_raw, int) else 1
        anchor_width_raw = parsed.get("width", 1)
        anchor_width = anchor_width_raw if isinstance(anchor_width_raw, int) else 1

        row_offset = _is_const_int(args[1])
        col_offset = _is_const_int(args[2])
        height_const = _is_const_int(args[3]) if len(args) >= 4 else None
        width_const = _is_const_int(args[4]) if len(args) >= 5 else None

        # Detect explicit-but-volatile dimensions before falling back
        # to anchor defaults (so we can record the coarsening reason).
        height_arg_present = len(args) >= 4
        width_arg_present = len(args) >= 5
        height_arg_volatile = height_arg_present and height_const is None
        width_arg_volatile = width_arg_present and width_const is None

        # Default dimensions: if height/width are absent, OFFSET inherits
        # the anchor's own dimensions per Excel semantics.
        if not height_arg_present:
            height_const = anchor_height
        if not width_arg_present:
            width_const = anchor_width

        is_static = (
            row_offset is not None
            and col_offset is not None
            and not height_arg_volatile
            and not width_arg_volatile
        )

        if is_static:
            target_row = anchor_row + (row_offset or 0)
            target_col = anchor_col + (col_offset or 0)
            target_height = height_const if height_const and height_const >= 1 else 1
            target_width = width_const if width_const and width_const >= 1 else 1
            # Bounds check: target row/col must be positive.
            if target_row < 1 or target_col < 1:
                results.append(
                    OffsetCall(
                        kind="skipped",
                        anchor_ref=anchor_ref_text,
                        skipped_reason="static_out_of_bounds",
                    )
                )
                continue
            results.append(
                OffsetCall(
                    kind="static",
                    anchor_ref=anchor_ref_text,
                    target_sheet=anchor_sheet,
                    target_row=target_row,
                    target_col=target_col,
                    target_height=target_height,
                    target_width=target_width,
                )
            )
            continue

        # Volatile: at least one of row/col offset is non-constant, or
        # an explicit dimension is formula-driven. Build a coarse range
        # using whatever constants are available.
        coarsening: list[str] = []
        if row_offset is None:
            coarsening.append("row_offset_volatile")
        if col_offset is None:
            coarsening.append("col_offset_volatile")
        if height_arg_volatile:
            coarsening.append("height_volatile")
        if width_arg_volatile:
            coarsening.append("width_volatile")

        # Volatile starting position: when the offset is volatile, the
        # earliest row/col we can guarantee is the anchor itself
        # (anchoring downward / rightward is the conservative choice).
        if row_offset is None:
            target_row = anchor_row
        else:
            target_row = anchor_row + row_offset
        if col_offset is None:
            target_col = anchor_col
        else:
            target_col = anchor_col + col_offset
        if target_row < 1:
            target_row = 1
        if target_col < 1:
            target_col = 1

        # When the dimension is a known constant, propagate it. When
        # it's missing-and-volatile or explicitly volatile, leave None
        # — the resolver clips to the anchor sheet's populated extent.
        target_height: int | None
        target_width: int | None
        if height_arg_volatile or (height_const is None):
            target_height = None
        elif height_const >= 1:
            target_height = height_const
        else:
            target_height = None

        if width_arg_volatile or (width_const is None):
            target_width = None
        elif width_const >= 1:
            target_width = width_const
        else:
            target_width = None

        results.append(
            OffsetCall(
                kind="volatile",
                anchor_ref=anchor_ref_text,
                target_sheet=anchor_sheet,
                target_row=target_row,
                target_col=target_col,
                target_height=target_height,
                target_width=target_width,
                coarsening_note=",".join(coarsening) or None,
            )
        )

    return results


# --- IR resolution: OffsetCall → binding-edge tuple -------------------------


def _col_letter_to_num(col_str: str) -> int:
    result = 0
    for ch in col_str.upper():
        if ch < "A" or ch > "Z":
            return 0
        result = result * 26 + (ord(ch) - 64)
    return result


def _col_num_to_letter(col: int) -> str:
    out = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        out = chr(65 + rem) + out
    return out or "A"


def _bindings_for_rect(
    conn: sqlite3.Connection,
    sheet_id: int,
    r1: int,
    c1: int,
    r2: int,
    c2: int,
) -> set[str]:
    """Find binding_ids whose extent overlaps `sheet[r1:r2, c1:c2]`."""
    rows = conn.execute(
        """
        SELECT DISTINCT ctb.binding_id
        FROM cells c
        JOIN cell_to_binding ctb ON c.cell_id = ctb.cell_id
        WHERE c.sheet_id = ?
          AND c.row BETWEEN ? AND ?
          AND c.col BETWEEN ? AND ?
        """,
        (sheet_id, r1, r2, c1, c2),
    ).fetchall()
    return {row[0] for row in rows}


def _resolve_offset_to_rect(
    conn: sqlite3.Connection,
    target_sheet_id: int,
    call: OffsetCall,
) -> tuple[int, int, int, int] | None:
    """Resolve an OffsetCall to a concrete (r1, c1, r2, c2) on its target sheet.

    For 'static' calls this is the exact range. For 'volatile' calls,
    missing dimensions are clipped to the populated extent of the target
    sheet — over-include cells rather than miss the dependency entirely.
    """
    if call.target_row is None or call.target_col is None:
        return None
    r1 = max(1, call.target_row)
    c1 = max(1, call.target_col)

    if call.kind == "static":
        h = call.target_height or 1
        w = call.target_width or 1
        return (r1, c1, r1 + h - 1, c1 + w - 1)

    # Volatile: fill missing dimensions from the populated extent.
    max_row_row = conn.execute(
        "SELECT MAX(row), MAX(col) FROM cells WHERE sheet_id = ?",
        (target_sheet_id,),
    ).fetchone()
    if max_row_row is None or max_row_row[0] is None or max_row_row[1] is None:
        return None
    sheet_max_row = int(max_row_row[0])
    sheet_max_col = int(max_row_row[1])

    if call.target_height is not None and call.target_height >= 1:
        r2 = r1 + call.target_height - 1
    else:
        r2 = sheet_max_row
    if call.target_width is not None and call.target_width >= 1:
        c2 = c1 + call.target_width - 1
    else:
        c2 = sheet_max_col

    # Clip to sheet's populated extent so we don't over-extend past
    # the actual data tail.
    if r2 > sheet_max_row:
        r2 = sheet_max_row
    if c2 > sheet_max_col:
        c2 = sheet_max_col
    if r2 < r1 or c2 < c1:
        return None
    return (r1, c1, r2, c2)


# --- Public driver ---------------------------------------------------------


def synthesize_offset_edges(conn: sqlite3.Connection) -> dict[str, int]:
    """Synthesise `via_offset_*` binding_edges for every OFFSET call in the IR.

    Walks every cell whose `formula_a1` contains the literal token
    'OFFSET' (case-insensitive); for each such formula, parses the AST,
    resolves each OFFSET to a target rectangle on its anchor sheet, and
    emits a binding edge from the cell's binding to every binding whose
    extent overlaps the rectangle.

    Returns a metrics dict:
    - calls_seen           — total OFFSET calls parsed
    - calls_static         — calls resolved as exact (kind='via_offset_static')
    - calls_volatile       — calls resolved as coarse (kind='via_offset_volatile')
    - calls_skipped        — calls skipped (INDIRECT / nested / non-ref / etc.)
    - edges_inserted       — new edge rows in binding_edges (de-duped)
    """
    metrics = {
        "calls_seen": 0,
        "calls_static": 0,
        "calls_volatile": 0,
        "calls_skipped": 0,
        "edges_inserted": 0,
    }

    # Schema gates
    tables = {
        r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    }
    needed = {"cells", "binding_edges", "cell_to_binding", "sheets"}
    if not needed.issubset(tables):
        return metrics

    # sheets lookup
    sheet_id_of: dict[str, int] = {}
    sheet_name_of: dict[int, str] = {}
    for sheet_id, sheet_name in conn.execute("SELECT sheet_id, sheet_name FROM sheets"):
        sheet_id_of[sheet_name.lower()] = sheet_id
        sheet_name_of[sheet_id] = sheet_name

    # Find every formula-bearing cell whose A1 form contains 'OFFSET'.
    # We use cells.formula_a1 (per-cell A1 form) so column-relative
    # OFFSETs that shift across cells are parsed correctly.
    rows = conn.execute(
        """
        SELECT c.cell_id, c.sheet_id, c.formula_a1, ctb.binding_id
        FROM cells c
        JOIN cell_to_binding ctb ON c.cell_id = ctb.cell_id
        WHERE c.formula_a1 IS NOT NULL
          AND UPPER(c.formula_a1) LIKE '%OFFSET%'
        """
    ).fetchall()

    edges: set[tuple[str, str, str]] = set()  # (from_bid, to_bid, kind)

    for cell_id, sheet_id, formula_a1, from_bid in rows:
        current_sheet = sheet_name_of.get(sheet_id, "")
        try:
            calls = parse_offset_calls(formula_a1, current_sheet)
        except Exception as exc:  # noqa: BLE001 — defensive: never crash the build
            logger.debug(
                "OFFSET parse failed for cell_id=%s formula=%r: %s",
                cell_id,
                (formula_a1 or "")[:80],
                exc,
            )
            continue
        metrics["calls_seen"] += len(calls)

        for call in calls:
            if call.kind == "skipped":
                metrics["calls_skipped"] += 1
                continue
            if call.target_sheet is None:
                metrics["calls_skipped"] += 1
                continue
            target_sheet_id = sheet_id_of.get(call.target_sheet.lower())
            if target_sheet_id is None:
                metrics["calls_skipped"] += 1
                continue
            try:
                rect = _resolve_offset_to_rect(conn, target_sheet_id, call)
            except sqlite3.Error as exc:
                logger.debug(
                    "OFFSET resolver failed for cell_id=%s anchor=%s: %s",
                    cell_id,
                    call.anchor_ref,
                    exc,
                )
                metrics["calls_skipped"] += 1
                continue
            if rect is None:
                metrics["calls_skipped"] += 1
                continue
            r1, c1, r2, c2 = rect
            target_bids = _bindings_for_rect(conn, target_sheet_id, r1, c1, r2, c2)
            if not target_bids:
                metrics["calls_skipped"] += 1
                continue
            edge_kind = "via_offset_static" if call.kind == "static" else "via_offset_volatile"
            if call.kind == "static":
                metrics["calls_static"] += 1
            else:
                metrics["calls_volatile"] += 1
            for tbid in target_bids:
                if tbid == from_bid:
                    continue
                edges.add((from_bid, tbid, edge_kind))

    if edges:
        # Insert new edges. The schema's PRIMARY KEY is (from, to) so an
        # existing 'formula' edge between the same pair will block our
        # 'via_offset_*' insert (INSERT OR IGNORE). That's fine: the
        # formula edge already encodes the dependency. Fix D's value is
        # the edges that didn't exist before.
        rows_to_insert = [(from_bid, to_bid, 1, kind, None) for (from_bid, to_bid, kind) in edges]
        # Snapshot the via_offset_* edge count before insert so we can
        # report the actual delta (executemany.rowcount is unreliable
        # across drivers and may report -1 or the cumulative count).
        before = conn.execute(
            "SELECT COUNT(*) FROM binding_edges WHERE kind LIKE 'via_offset_%'"
        ).fetchone()[0]
        conn.executemany(
            """
            INSERT OR IGNORE INTO binding_edges
                (from_binding_id, to_binding_id, edge_count, kind, provenance_proc)
            VALUES (?, ?, ?, ?, ?)
            """,
            rows_to_insert,
        )
        after = conn.execute(
            "SELECT COUNT(*) FROM binding_edges WHERE kind LIKE 'via_offset_%'"
        ).fetchone()[0]
        metrics["edges_inserted"] = after - before
    conn.commit()
    return metrics
