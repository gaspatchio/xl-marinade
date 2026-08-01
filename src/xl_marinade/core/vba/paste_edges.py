# ABOUTME: VBA paste-edge synthesiser (R21 Fix B). Static analysis of VBA paste-as-values
# ABOUTME: statements (PasteSpecial / .Value=.Value / .Value=Array) emits `via_vba_paste`
# ABOUTME: binding_edges from source-template bindings to paste-target output bindings.

"""VBA paste-edge synthesiser (R21 Fix B).

Background (per `r21-multifactor-investigation-vba-paste-edge-synthesis-2026-04-29.md`):

Actuarial / cashflow workbooks frequently use a VBA macro that copies a
"template row" of formulas (e.g. `Risk Drivers!A2:CK2`) into an output
block (e.g. `Risk Drivers!A8:CK847`) one row per scenario iteration. The
copy is a *paste-as-values* — the destination cells receive constants, not
formulas. Consequence: the formula DAG cannot reach from the output block
back to Calculation Engine through the formula precedents that live ONLY
on the template row. Without a bridging edge, the lineage walker
dead-ends at the output block.

This module emits those bridging edges as
`binding_edges (kind='via_vba_paste')` rows, with `provenance_proc`
naming the VBA procedure responsible. The synthesiser handles three
syntactic patterns:

1. **PasteSpecial idiom** (a model's `loop_calc()`): state-machine
   emulation of the `Selection` cursor across `.Select` → `.Copy` →
   `.PasteSpecial Paste:=xlPasteValues` lines.
2. **`.Value = .Value` direct assignment** (a large model's `Calculations`):
   `lhs_range.Value = rhs_range.Value` line-level pattern match with
   loop-variable string-concat resolution.
3. **`.Value = Array(...)` tuple assignment** (a large model's risk-margin row):
   the source is a constant tuple, no edge emitted (no source binding to
   chain to); a debug log records the skip.

False-positive control: paste statements inside `On Error GoTo`
error-handler blocks are skipped (R21 §3.6 measured 0/9 paste statements
in bench workbooks fall in handlers, but the check is trivial and
future-proofs the synthesiser).

Integration: `synthesize_paste_edges(conn)` is called from
`fast_extraction_pipeline._wire_vba_edges` after the formula-DAG and
VBA-static-ref edges are wired. The synthesiser only inserts new edge
rows; it never modifies or removes existing edges.
"""

from __future__ import annotations

import logging
import re
import sqlite3
from dataclasses import dataclass

logger = logging.getLogger(__name__)


# --- Regex catalog ---------------------------------------------------------

# Sheet activation: Sheets("X").Select  / Worksheets("X").Activate
_SHEET_SELECT_RE = re.compile(
    r'(?:Sheets|Worksheets)\("([^"]+)"\)\s*\.\s*(?:Select|Activate)\b',
    re.IGNORECASE,
)

# Cursor reset: Range("X").Select where X is a literal A1 or named range
_RANGE_SELECT_RE = re.compile(
    r'Range\("([^"]+)"\)\s*\.\s*Select\b',
    re.IGNORECASE,
)

# Cursor reset with offset: Range("name").Offset(r,c).Select
_RANGE_OFFSET_SELECT_RE = re.compile(
    r'Range\("([^"]+)"\)\s*\.\s*Offset\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)\s*\.\s*Select\b',
    re.IGNORECASE,
)

# Selection extension: Range(Selection, Selection.End(xlToRight)).Select etc.
_EXTEND_TO_RIGHT_RE = re.compile(
    r"Range\(Selection,\s*Selection\.End\(\s*xlToRight\s*\)\)\s*\.\s*Select",
    re.IGNORECASE,
)
_EXTEND_DOWN_RE = re.compile(
    r"Range\(Selection,\s*Selection\.End\(\s*xlDown\s*\)\)\s*\.\s*Select",
    re.IGNORECASE,
)

# Selection.Copy
_SELECTION_COPY_RE = re.compile(r"Selection\s*\.\s*Copy\b", re.IGNORECASE)

# Append-row pattern: Selection.End(xlDown).Offset(r,c).Select
_APPEND_OFFSET_SELECT_RE = re.compile(
    r"Selection\s*\.\s*End\(\s*xlDown\s*\)\s*\.\s*Offset\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)\s*\.\s*Select",
    re.IGNORECASE,
)

# PasteSpecial Paste:=xlPasteValues / xlPasteFormulas / xlPasteFormats
_PASTE_SPECIAL_RE = re.compile(
    r"(?:Selection|ActiveSheet\.Paste|ActiveCell)\s*\.\s*PasteSpecial\s+Paste\s*:=\s*xl(\w+)",
    re.IGNORECASE,
)

# Direct value-assignment patterns (a large model's `.Value = .Value` family).
# LHS: optional ws-var / sheet-qual prefix, then Range("..." | "X" & v & ":Y" & v).
# RHS: same structure, ending with .Value (range-to-range copy) OR Array(...) (tuple).
_VALUE_ASSIGN_RANGE_RE = re.compile(
    r'(?P<lhs_pre>(?:\w+\.)?(?:Sheets\("[^"]+"\)\.|Worksheets\("[^"]+"\)\.|))'
    r"Range\((?P<lhs_arg>[^()]+)\)\s*\.\s*Value\s*=\s*"
    r'(?P<rhs_pre>(?:\w+\.)?(?:Sheets\("[^"]+"\)\.|Worksheets\("[^"]+"\)\.|))'
    r"Range\((?P<rhs_arg>[^()]+)\)\s*\.\s*Value\b",
    re.IGNORECASE,
)
_VALUE_ASSIGN_ARRAY_RE = re.compile(
    r'(?P<lhs_pre>(?:\w+\.)?(?:Sheets\("[^"]+"\)\.|Worksheets\("[^"]+"\)\.|))'
    r"Range\((?P<lhs_arg>[^()]+)\)\s*\.\s*Value\s*=\s*Array\(",
    re.IGNORECASE,
)

# Worksheet-variable assignment for resolving wsCalc.Range("...") prefixes:
#   Set wsCalc = Worksheets("Calculations")
_WS_VAR_ASSIGN_RE = re.compile(
    r'Set\s+(\w+)\s*=\s*\w*\.?(?:Worksheets?|Sheets)\("([^"]+)"\)',
    re.IGNORECASE,
)

# Error-handler block markers (false-positive control).
_ON_ERROR_GOTO_RE = re.compile(r"On\s+Error\s+GoTo\s+(\w+)", re.IGNORECASE)
_LABEL_RE = re.compile(r"^\s*(\w+)\s*:\s*$")

# A1 cell address (case-insensitive). Used to discriminate named ranges.
_A1_RE = re.compile(r"^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$", re.IGNORECASE)

# String-built loop-row argument: e.g.  "A" & lastRow + 1 & ":ALI" & lastRow + 12
_LOOP_BUILT_RE = re.compile(
    r'^"([A-Z$]+)"\s*&[^&]+&\s*"\s*:\s*([A-Z$]+)"\s*&',
    re.IGNORECASE,
)


# --- Selection cursor state ------------------------------------------------


@dataclass
class _Selection:
    """The state machine's current Selection cursor."""

    sheet: str | None = None
    base: str | None = None  # A1 cell ref or named range
    base_kind: str = "literal"  # 'literal' | 'named_range'
    extent_to_right: bool = False
    extent_down: bool = False
    append_offset: tuple[int, int] | None = None  # End(xlDown).Offset(r,c)


@dataclass
class PasteEvent:
    """A single resolved paste statement.

    `kind` is the syntactic category:
    - 'paste_special_values' — Selection.PasteSpecial Paste:=xlPasteValues
    - 'paste_special_formulas' — Selection.PasteSpecial Paste:=xlPasteFormulas
    - 'value_assign_range' — `lhs.Value = rhs.Value`
    - 'value_assign_array' — `lhs.Value = Array(...)` (no source binding)

    `source_sheet`/`source_range_a1` describe the COPY origin (the template
    row); `target_sheet`/`target_range_a1` describe the paste destination.
    `procedure` is `module_name::procedure_name` (provenance attribution).
    """

    procedure: str
    line_number: int
    kind: str
    source_sheet: str | None
    source_range_a1: str | None
    target_sheet: str | None
    target_range_a1: str | None
    notes: str = ""


# --- Helpers ----------------------------------------------------------------


def _is_a1(target: str) -> bool:
    return bool(_A1_RE.match(target.strip()))


def _strip_inline_comment(line: str) -> str:
    """Drop a trailing `'comment` outside of string literals."""
    in_string = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_string = not in_string
        elif ch == "'" and not in_string:
            return line[:i]
    return line


def _classify_value_assign_arg(arg: str) -> str:
    """Classify the inner argument of a `Range(...)` reference.

    Returns one of:
    - 'literal_a1'  — `"A2:ALI13"` (statically resolvable)
    - 'loop_built'  — `"A" & lastRow + 1 & ":ALI" & lastRow + 12` (loop append)
    - 'string_built' — concat with `&` but no recognised loop var pattern
    - 'cells_call' — `Cells(r,c)` (out of scope for v1)
    - 'other'      — unrecognised; skip
    """
    s = arg.strip()
    if _A1_RE.match(s.strip('"')):
        return "literal_a1"
    if "&" in s and _LOOP_BUILT_RE.search(s):
        return "loop_built"
    if "&" in s:
        return "string_built"
    if "Cells(" in s:
        return "cells_call"
    return "other"


def _extract_sheet_qual(prefix: str, ws_vars: dict[str, str]) -> str | None:
    """Resolve a sheet qualifier from a `lhs_pre`/`rhs_pre` prefix.

    Handles `Sheets("X").`, `Worksheets("X").`, and `wsVar.` variants.
    Returns None for unqualified prefix (paste applies to the cursor's
    active sheet).
    """
    if not prefix:
        return None
    m = re.search(
        r'(?:Sheets|Worksheets)\("([^"]+)"\)\.?',
        prefix,
        flags=re.IGNORECASE,
    )
    if m:
        return m.group(1)
    # ws_var prefix: extract identifier and look up in ws_vars
    m = re.match(r"(\w+)\.", prefix)
    if m:
        return ws_vars.get(m.group(1).lower())
    return None


def _normalize_a1(arg: str) -> str | None:
    """Strip surrounding quotes and `$` markers from an A1 literal."""
    s = arg.strip().strip('"')
    s = s.replace("$", "")
    if _A1_RE.match(s):
        return s.upper()
    return None


# --- State-machine parse for PasteSpecial idiom -----------------------------


def _parse_paste_special_events(procedure: str, body: str) -> list[PasteEvent]:
    """Parse PasteSpecial-idiom paste events from a procedure body.

    Walks lines, tracking the current sheet and the Selection cursor
    (with extension flags and append-offset). Emits one PasteEvent per
    PasteSpecial line. Skips events whose source is undetermined (no
    Selection.Copy preceded the paste).
    """
    events: list[PasteEvent] = []
    cur_sheet: str | None = None
    sel = _Selection()
    last_copied: _Selection | None = None
    last_after_copy_dest: _Selection | None = None
    in_error_handler = False
    error_label: str | None = None

    for i, raw in enumerate(body.splitlines(), 1):
        line = _strip_inline_comment(raw).strip()
        if not line:
            continue

        # Error-handler tracking (false-positive control).
        m = _ON_ERROR_GOTO_RE.search(line)
        if m:
            error_label = m.group(1)
            continue
        m = _LABEL_RE.match(line)
        if m and error_label and m.group(1) == error_label:
            in_error_handler = True
            continue
        if in_error_handler:
            # Skip everything inside the handler body.
            continue

        # Sheet activation
        m = _SHEET_SELECT_RE.search(line)
        if m:
            cur_sheet = m.group(1)
            sel = _Selection(sheet=cur_sheet)
            continue

        # Range("...").Select  → cursor reset
        m = _RANGE_SELECT_RE.match(line)
        if m:
            target = m.group(1)
            sel = _Selection(
                sheet=cur_sheet,
                base=target,
                base_kind="literal" if _is_a1(target) else "named_range",
            )
            if last_copied is not None:
                last_after_copy_dest = sel
            continue

        # Range("name").Offset(r,c).Select  → cursor reset with offset
        m = _RANGE_OFFSET_SELECT_RE.match(line)
        if m:
            target = m.group(1)
            roff, coff = int(m.group(2)), int(m.group(3))
            sel = _Selection(
                sheet=cur_sheet,
                base=f"{target}.Offset({roff},{coff})",
                base_kind="literal" if _is_a1(target) else "named_range",
            )
            if last_copied is not None:
                last_after_copy_dest = sel
            continue

        # Range(Selection, Selection.End(xlToRight)).Select
        if _EXTEND_TO_RIGHT_RE.search(line):
            sel.extent_to_right = True
            continue
        if _EXTEND_DOWN_RE.search(line):
            sel.extent_down = True
            continue

        # Selection.Copy  → snapshot Selection as source
        if _SELECTION_COPY_RE.search(line):
            last_copied = _Selection(
                sheet=sel.sheet,
                base=sel.base,
                base_kind=sel.base_kind,
                extent_to_right=sel.extent_to_right,
                extent_down=sel.extent_down,
            )
            continue

        # Selection.End(xlDown).Offset(r,c).Select  → append-row pattern
        m = _APPEND_OFFSET_SELECT_RE.search(line)
        if m:
            sel.append_offset = (int(m.group(1)), int(m.group(2)))
            if last_copied is not None:
                last_after_copy_dest = sel
            continue

        # PasteSpecial Paste:=xl... → emit event
        m = _PASTE_SPECIAL_RE.search(line)
        if m:
            paste_token = "xl" + m.group(1)
            paste_kind_lower = paste_token.lower()
            # Skip format-only pastes — no data edge.
            if paste_kind_lower == "xlpasteformats":
                continue

            target_sel = last_after_copy_dest or sel
            source_sel = last_copied
            if source_sel is None or source_sel.base is None:
                # Paste with no preceding Copy — cannot resolve source.
                continue

            kind = (
                "paste_special_values"
                if paste_kind_lower == "xlpastevalues"
                else "paste_special_formulas"
            )

            source_range_a1 = source_sel.base
            target_range_a1 = target_sel.base if target_sel.base else None
            # If the target is "named_range.Offset(r,c).Select" we keep the
            # named-range anchor as the symbolic target (the synthesiser
            # resolves it against `defined_names` / cell_to_binding so that
            # the entire output block becomes the target binding).
            events.append(
                PasteEvent(
                    procedure=procedure,
                    line_number=i,
                    kind=kind,
                    source_sheet=source_sel.sheet,
                    source_range_a1=source_range_a1,
                    target_sheet=target_sel.sheet,
                    target_range_a1=target_range_a1,
                    notes=(
                        f"src_extents=({source_sel.extent_to_right},"
                        f"{source_sel.extent_down}) "
                        f"tgt_append={target_sel.append_offset}"
                    ),
                )
            )
            continue

    return events


# --- Direct value-assignment parse (.Value = .Value family) -----------------


def _parse_value_assign_events(procedure: str, body: str) -> list[PasteEvent]:
    """Parse `.Value = .Value` and `.Value = Array(...)` events."""
    events: list[PasteEvent] = []
    ws_vars: dict[str, str] = {}
    in_error_handler = False
    error_label: str | None = None

    # Pre-pass: collect Set ws = Worksheets("X") assignments.
    for raw in body.splitlines():
        line = _strip_inline_comment(raw).strip()
        m = _WS_VAR_ASSIGN_RE.search(line)
        if m:
            ws_vars[m.group(1).lower()] = m.group(2)

    for i, raw in enumerate(body.splitlines(), 1):
        line = _strip_inline_comment(raw).strip()
        if not line:
            continue

        m = _ON_ERROR_GOTO_RE.search(line)
        if m:
            error_label = m.group(1)
            continue
        m = _LABEL_RE.match(line)
        if m and error_label and m.group(1) == error_label:
            in_error_handler = True
            continue
        if in_error_handler:
            continue

        # Range(...).Value = Range(...).Value  (R21 large-model pattern)
        m = _VALUE_ASSIGN_RANGE_RE.search(line)
        if m:
            lhs_arg = m.group("lhs_arg")
            rhs_arg = m.group("rhs_arg")
            lhs_kind = _classify_value_assign_arg(lhs_arg)
            rhs_kind = _classify_value_assign_arg(rhs_arg)

            # Both sides must be statically resolvable for an edge.
            # 'literal_a1' is fully resolvable; 'loop_built' on the LHS is
            # a loop-append (resolvable to "the entire output region"
            # downstream).
            target_a1 = _normalize_a1(lhs_arg) if lhs_kind == "literal_a1" else None
            source_a1 = _normalize_a1(rhs_arg) if rhs_kind == "literal_a1" else None
            target_loop = lhs_kind == "loop_built"

            if source_a1 is None or (target_a1 is None and not target_loop):
                # Skip dynamic / unresolvable; out of scope for v1.
                continue

            source_sheet = _extract_sheet_qual(m.group("rhs_pre"), ws_vars)
            target_sheet = _extract_sheet_qual(m.group("lhs_pre"), ws_vars)

            events.append(
                PasteEvent(
                    procedure=procedure,
                    line_number=i,
                    kind="value_assign_range",
                    source_sheet=source_sheet,
                    source_range_a1=source_a1,
                    target_sheet=target_sheet,
                    target_range_a1=target_a1 if target_a1 else "<loop_append>",
                    notes=f"lhs={lhs_kind} rhs={rhs_kind}",
                )
            )
            continue

        # Range(...).Value = Array(...)  (tuple literal — no source binding)
        m = _VALUE_ASSIGN_ARRAY_RE.search(line)
        if m:
            lhs_arg = m.group("lhs_arg")
            lhs_kind = _classify_value_assign_arg(lhs_arg)
            target_a1 = _normalize_a1(lhs_arg) if lhs_kind == "literal_a1" else None
            target_loop = lhs_kind == "loop_built"
            if target_a1 is None and not target_loop:
                continue
            target_sheet = _extract_sheet_qual(m.group("lhs_pre"), ws_vars)
            events.append(
                PasteEvent(
                    procedure=procedure,
                    line_number=i,
                    kind="value_assign_array",
                    source_sheet=None,
                    source_range_a1=None,
                    target_sheet=target_sheet,
                    target_range_a1=target_a1 if target_a1 else "<loop_append>",
                    notes=f"lhs={lhs_kind} (tuple_literal)",
                )
            )

    return events


def parse_paste_events(procedure: str, body: str) -> list[PasteEvent]:
    """Parse all paste events in a procedure body (both syntactic patterns).

    Returns a flat list of PasteEvents; callers translate these into
    binding-edge inserts via `_resolve_event_to_edge`.
    """
    out: list[PasteEvent] = []
    out.extend(_parse_paste_special_events(procedure, body))
    out.extend(_parse_value_assign_events(procedure, body))
    return out


# --- IR resolution: PasteEvent → binding-edge tuple -------------------------


def _col_letter_to_num(col_str: str) -> int:
    result = 0
    for ch in col_str.upper():
        result = result * 26 + (ord(ch) - 64)
    return result


def _bindings_for_range(
    conn: sqlite3.Connection,
    sheet_id: int,
    address: str,
) -> set[str]:
    """Find binding_ids whose extent overlaps `sheet!address`.

    `address` may be a single cell ('A2') or a range ('A2:CK2'). Returns
    the set of binding_ids (the source-template row may map to multiple
    bindings — e.g. one per formula family — and the target output block
    may also map to multiple bindings).
    """
    m = re.match(
        r"([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$",
        address.strip().upper(),
    )
    if not m:
        return set()
    c1 = _col_letter_to_num(m.group(1))
    r1 = int(m.group(2))
    c2 = _col_letter_to_num(m.group(3)) if m.group(3) else c1
    r2 = int(m.group(4)) if m.group(4) else r1

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


def _resolve_named_range_to_address(
    named_ranges: dict[str, str], name: str
) -> tuple[str, str] | None:
    """Resolve a named-range identifier to (sheet, address_a1).

    Returns None if the name isn't defined or the destination isn't a
    sheet-qualified A1 reference.
    """
    dest = named_ranges.get(name.upper())
    if not dest or "!" not in dest:
        return None
    sheet, addr = dest.split("!", 1)
    sheet = sheet.strip("'").strip()
    addr = addr.replace("$", "")
    return sheet, addr


def _output_block_address_for_named_range(
    conn: sqlite3.Connection,
    sheet_id: int,
    anchor_address: str,
) -> str:
    """Compute the A1 range of the output block anchored at a named range.

    Convention: the output block extends from the row AFTER the anchor
    (which holds headers in the loop_calc idiom) down to the last
    populated row on that sheet. Columns span the anchor's column extent
    (or up to the sheet's last column when the anchor is a single cell).
    """
    # Anchor parsing: handle both A1 and A1:Z1 forms.
    m = re.match(
        r"([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$",
        anchor_address.strip().upper(),
    )
    if not m:
        return anchor_address
    c1 = _col_letter_to_num(m.group(1))
    r1 = int(m.group(2))
    c2 = _col_letter_to_num(m.group(3)) if m.group(3) else c1
    r2 = int(m.group(4)) if m.group(4) else r1
    # Find the last populated row on this sheet for the column range.
    row = conn.execute(
        """
        SELECT MAX(row) FROM cells
        WHERE sheet_id = ? AND col BETWEEN ? AND ?
        """,
        (sheet_id, c1, c2),
    ).fetchone()
    last_row = row[0] if row and row[0] else r2
    if last_row <= r1:
        return anchor_address
    # Output block covers rows after the anchor through last populated row.
    block_r1 = max(r1 + 1, r2 + 1)
    return f"{m.group(1)}{block_r1}:{m.group(3) or m.group(1)}{last_row}"


def _resolve_source_extent_to_a1(
    conn: sqlite3.Connection,
    sheet_id: int,
    base_address: str,
    notes: str,
) -> str | None:
    """Resolve a paste source's data-bounded extent to an A1 range.

    `base_address` is the A1 anchor (e.g. "A2"); `notes` carries the
    state-machine's extent flags. We always resolve to the entire
    populated row span starting at `base_address` — the column extent is
    the maximum populated column for that row on that sheet. This is the
    "data-bounded extent" decision (R21 §2.3).
    """
    # Single-cell anchor (no extent applied) → return as-is.
    m = re.match(r"([A-Z]{1,3})(\d+)$", base_address.strip().upper())
    if not m:
        # Already a range (e.g. "A2:CK2") — return verbatim.
        return base_address.strip().upper()
    base_col = m.group(1)
    base_row = int(m.group(2))
    base_col_num = _col_letter_to_num(base_col)
    # Read the maximum populated column on `base_row` on `sheet_id`.
    row = conn.execute(
        """
        SELECT MAX(col) FROM cells
        WHERE sheet_id = ? AND row = ? AND col >= ?
        """,
        (sheet_id, base_row, base_col_num),
    ).fetchone()
    last_col = row[0] if row and row[0] else base_col_num
    if last_col <= base_col_num:
        return f"{base_col}{base_row}"
    end_col_letter = _col_num_to_letter(last_col)
    return f"{base_col}{base_row}:{end_col_letter}{base_row}"


def _col_num_to_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result or "A"


# --- Public driver ---------------------------------------------------------


def synthesize_paste_edges(conn: sqlite3.Connection) -> dict[str, int]:
    """Synthesise `via_vba_paste` binding_edges for every paste event in
    the workbook's VBA.

    Returns a metrics dict:
    - events_seen        — total paste events parsed (all kinds)
    - events_resolved    — events whose source AND target resolved to bindings
    - events_unresolved  — events skipped (source/target unresolvable)
    - edges_inserted     — new edge rows in binding_edges (de-duped)
    """
    metrics = {
        "events_seen": 0,
        "events_resolved": 0,
        "events_unresolved": 0,
        "edges_inserted": 0,
    }

    # Schema gates
    tables = {
        r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    }
    needed = {
        "vba_procedures",
        "vba_modules",
        "binding_edges",
        "cell_to_binding",
        "cells",
        "sheets",
    }
    if not needed.issubset(tables):
        return metrics

    # sheets lookup
    sheet_id_of: dict[str, int] = {}
    for sheet_id, sheet_name in conn.execute("SELECT sheet_id, sheet_name FROM sheets"):
        sheet_id_of[sheet_name.lower()] = sheet_id

    # named-range lookup (for resolving "out_*" paste targets)
    named_ranges: dict[str, str] = {}
    if "defined_names" in tables:
        try:
            import json

            for name, dests in conn.execute("SELECT name, destinations FROM defined_names"):
                try:
                    parsed = json.loads(dests)
                    if isinstance(parsed, list) and parsed:
                        first = parsed[0]
                        if isinstance(first, str) and not first.startswith("#"):
                            named_ranges[name.upper()] = first
                except (json.JSONDecodeError, IndexError):
                    pass
        except sqlite3.OperationalError:
            pass

    # Gather all VBA procedures.
    procs = conn.execute(
        """
        SELECT m.name AS module_name, p.name AS proc_name, p.body
        FROM vba_procedures p
        JOIN vba_modules m ON p.module_id = m.module_id
        """
    ).fetchall()

    edges: set[tuple[str, str, str]] = set()  # (from_bid, to_bid, proc_qual)

    for module_name, proc_name, body in procs:
        if not body:
            continue
        proc_qual = f"{module_name}::{proc_name}"
        events = parse_paste_events(proc_qual, body)
        metrics["events_seen"] += len(events)

        for ev in events:
            try:
                resolved = _resolve_event_to_edges(
                    conn=conn,
                    event=ev,
                    sheet_id_of=sheet_id_of,
                    named_ranges=named_ranges,
                )
            except Exception as exc:  # noqa: BLE001 — defensive: never crash the build
                logger.debug(
                    "paste-edge resolution failed for %s line %d: %s",
                    ev.procedure,
                    ev.line_number,
                    exc,
                )
                metrics["events_unresolved"] += 1
                continue
            if not resolved:
                metrics["events_unresolved"] += 1
                continue
            metrics["events_resolved"] += 1
            for from_bid, to_bid in resolved:
                if from_bid == to_bid:
                    continue
                edges.add((from_bid, to_bid, proc_qual))

    if edges:
        # sorted(): binding_edges' PK is (from,to), so when two procedures
        # paste into the same target the first row wins provenance_proc.
        # Set iteration order varies per run (str hash randomization), which
        # made the winning label nondeterministic across otherwise identical
        # extractions.
        rows = [
            (from_bid, to_bid, 1, "via_vba_paste", proc_qual)
            for (from_bid, to_bid, proc_qual) in sorted(edges)
        ]
        conn.executemany(
            """
            INSERT OR IGNORE INTO binding_edges
                (from_binding_id, to_binding_id, edge_count, kind, provenance_proc)
            VALUES (?, ?, ?, ?, ?)
            """,
            rows,
        )
        metrics["edges_inserted"] = len(rows)
    conn.commit()
    return metrics


def _resolve_event_to_edges(
    conn: sqlite3.Connection,
    event: PasteEvent,
    sheet_id_of: dict[str, int],
    named_ranges: dict[str, str],
) -> list[tuple[str, str]]:
    """Translate a single PasteEvent into 0..N (from_bid, to_bid) tuples.

    Returns an empty list when source or target can't be resolved to any
    binding (which is then counted as `events_unresolved`).
    """
    if event.kind == "value_assign_array":
        # Array(...) RHS has no source binding; nothing to chain.
        return []

    if event.source_range_a1 is None or event.source_sheet is None:
        return []

    # Resolve source: start from the anchor; for state-machine events with
    # extent flags, extend to the data-bounded last column.
    source_sheet_id = sheet_id_of.get(event.source_sheet.lower())
    if source_sheet_id is None:
        return []
    source_a1 = _resolve_source_extent_to_a1(
        conn=conn,
        sheet_id=source_sheet_id,
        base_address=event.source_range_a1.split(".")[0],
        notes=event.notes,
    )
    if source_a1 is None:
        return []
    source_bids = _bindings_for_range(conn, source_sheet_id, source_a1)
    if not source_bids:
        return []

    # Resolve target. Three target representations:
    #   (a) Named range — `out_risk_driver`, `out_policy`, etc. — chain to
    #       the entire output block.
    #   (b) A1 literal — single cell or static range — direct resolution.
    #   (c) Loop append `<loop_append>` (a large model's `.Value=.Value`) — resolve
    #       via the LHS sheet's last populated row span.
    target_sheet = event.target_sheet or event.source_sheet  # intra-sheet default
    target_sheet_id = sheet_id_of.get(target_sheet.lower())
    if target_sheet_id is None:
        return []

    target_a1: str | None = None
    target_raw = event.target_range_a1 or ""
    target_no_offset = target_raw.split(".")[0].strip()

    if target_no_offset == "<loop_append>":
        # Compute the entire writeable region from row 2 onward across the
        # sheet's column extent. This is the conservative whole-block
        # representation used for a large model's `.Value=.Value` patterns.
        max_row = conn.execute(
            "SELECT MAX(row), MAX(col) FROM cells WHERE sheet_id = ?",
            (target_sheet_id,),
        ).fetchone()
        if not max_row or not max_row[0] or not max_row[1]:
            return []
        target_a1 = f"A2:{_col_num_to_letter(max_row[1])}{max_row[0]}"
    elif _A1_RE.match(target_no_offset):
        target_a1 = target_no_offset.upper()
    elif target_no_offset:
        # Treat as named range. Resolve via defined_names.
        resolved = _resolve_named_range_to_address(named_ranges, target_no_offset)
        if resolved is None:
            return []
        named_sheet, anchor_addr = resolved
        # Override target_sheet with the named-range's sheet.
        target_sheet_id = sheet_id_of.get(named_sheet.lower(), target_sheet_id)
        target_a1 = _output_block_address_for_named_range(
            conn=conn,
            sheet_id=target_sheet_id,
            anchor_address=anchor_addr,
        )

    if target_a1 is None:
        return []
    target_bids = _bindings_for_range(conn, target_sheet_id, target_a1)
    if not target_bids:
        return []

    # Cross-product: each source binding chains to each target binding.
    # The synthesised edge is (source → target) representing
    # "target inherits source's column-aligned dependency".
    edges: list[tuple[str, str]] = []
    for sbid in source_bids:
        for tbid in target_bids:
            edges.append((sbid, tbid))
    return edges
