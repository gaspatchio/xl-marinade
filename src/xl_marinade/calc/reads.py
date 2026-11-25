"""Whether a formula reads a map field's cell: provably, provably not, or unknown.

The formula graph holds an edge for every cell and range reference a formula writes,
empty cells included, and for the arguments of ``OFFSET`` and ``INDIRECT``. Some reads
have no edge: where an ``OFFSET`` or ``INDIRECT`` lands away from the stored inputs'
landing cell, the arguments of ``CELL``, ``INFO`` and ``GETPIVOTDATA`` (which the
extractor treats as opaque), a defined name holding a formula, and a 3-D reference
(``S1:S3!B1``). This module finds those so a field is refused as unread only when no
formula can read it.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from xl_marinade.calc.formula_ast import (
    ERROR_CODES,
    SHEET,
    function_name,
    iter_nodes,
    parse_formula,
)
from xl_marinade.calc.graph import NameDef, WorkbookGraph
from xl_marinade.calc.models import CellRef, _col_to_number
from xl_marinade.core.parser import ConstNode, FunctionNode, RefNode

Rect = tuple[str, int, int, int, int]

Read = bool | str
"""``True``/``False`` whether a requested output reads a field, or ``"unknown"``."""
UNKNOWN = "unknown"

_LANDING = ("OFFSET", "INDIRECT")
_OPAQUE = ("CELL", "INFO", "GETPIVOTDATA")
_A1_TEXT = re.compile(r"\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?")
_SHEET_PREFIX = re.compile(rf"^({SHEET})!")
_STRING = re.compile(r'"(?:[^"]|"")*"')
_CELL = r"\$?([A-Za-z]{1,3})\$?([0-9]+)"
# ``S1:S3!B1`` or ``'S 1:S 3'!B1:C2``: the same address on every sheet in the span.
_THREE_D = re.compile(
    rf"(?:'(?:[^':]|'')+:(?:[^']|'')+'|(?<![\w.'])[^\W\d][\w.]*:[^\W\d][\w.]*)!{_CELL}(?::{_CELL})?"
)
_WHOLE = re.compile(r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[0-9]+:\$?[0-9]+")
_NAME = re.compile(r"(?:[^\W\d]|[_\\])[\w.\\?]*")
_MAX_DEPTH = 8


@dataclass(slots=True)
class HiddenReads:
    """Reads the graph has no edge for.

    ``readers`` are provable reads with the formula cell that makes each.
    ``every_sheet`` and ``sheets`` say where a read may land that cannot be placed;
    ``addresses`` are the cells 3-D references read on sheets whose span is not known
    here (row 1, column 1, row 2, column 2).
    """

    readers: list[tuple[Rect, CellRef]] = field(default_factory=list)
    every_sheet: bool = False
    sheets: set[str] = field(default_factory=set)
    addresses: list[tuple[int, int, int, int]] = field(default_factory=list)

    def readers_of(self, cell: CellRef) -> set[CellRef]:
        """The formula cells whose hidden reads contain ``cell``."""
        return {reader for rect, reader in self.readers if _contains(rect, cell)}

    def maybe_read(self, cell: CellRef) -> bool:
        """True when a read that cannot be placed may reach ``cell``."""
        if self.every_sheet or cell.sheet in self.sheets:
            return True
        return any(
            r1 <= cell.row <= r2 and c1 <= cell.col <= c2 for r1, c1, r2, c2 in self.addresses
        )


def _contains(rect: Rect, cell: CellRef) -> bool:
    sheet, r1, c1, r2, c2 = rect
    return sheet == cell.sheet and r1 <= cell.row <= r2 and c1 <= cell.col <= c2


def field_read(
    graph: WorkbookGraph,
    cell: CellRef,
    reach: frozenset[CellRef],
    hidden: HiddenReads,
) -> Read | None:
    """Whether a requested output reads ``cell``: True or False when a formula
    provably reads it, ``"unknown"`` when only a read that cannot be placed may, and
    None when no formula can read it."""
    readers = set(graph.readers(cell)) | hidden.readers_of(cell)
    if readers and any(r in reach for r in readers):
        return True
    if hidden.maybe_read(cell):
        return UNKNOWN
    return False if readers else None


def hidden_reads(graph: WorkbookGraph) -> HiddenReads:
    """Scan every formula once for reads that have no edge in the graph."""
    out = HiddenReads()
    names = [d.name for d in graph.names()]
    name_re = (
        re.compile(
            r"(?<![\w.$\\])(?:"
            + "|".join(re.escape(n) for n in sorted(set(names), key=len, reverse=True))
            + r")(?![\w.(\[])",
            re.IGNORECASE,
        )
        if names
        else None
    )
    for cell in graph.formula_cells():
        text = graph.formula(cell) or ""
        bare = _STRING.sub('""', text)
        for m in _THREE_D.finditer(bare):
            r1, c1 = int(m.group(2)), _col_to_number(m.group(1))
            r2, c2 = (int(m.group(4)), _col_to_number(m.group(3))) if m.group(3) else (r1, c1)
            out.addresses.append((min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)))
        upper = bare.upper()
        calls = any(f"{name}(" in upper.replace(" ", "") for name in (*_LANDING, *_OPAQUE))
        if not calls and (name_re is None or name_re.search(bare) is None):
            continue
        root = parse_formula(text)
        if root is None:
            out.every_sheet = True
            continue
        for node in iter_nodes(root):
            if isinstance(node, FunctionNode):
                _call(graph, cell, node, out)
            elif isinstance(node, RefNode):
                _name_ref(graph, cell, node.ref, out)
    return out


def _call(graph: WorkbookGraph, cell: CellRef, node: FunctionNode, out: HiddenReads) -> None:
    name = function_name(node.name)
    first = node.args[0] if node.args else None
    if name in _LANDING:
        # Where it lands beyond the stored landing cell has no edge: that sheet.
        if name == "INDIRECT" and isinstance(first, ConstNode) and isinstance(first.value, str):
            landing = _sheets_of(graph, first.value, cell.sheet)
        elif name == "OFFSET" and isinstance(first, RefNode):
            landing = _sheets_of(graph, first.ref, cell.sheet)
        else:
            landing = None
        if landing is None:
            out.every_sheet = True
        else:
            out.sheets.update(landing)
    elif name in _OPAQUE:
        # The extractor records no edge for these arguments; they are plain reads.
        for arg in node.args:
            for inner in iter_nodes(arg):
                if isinstance(inner, RefNode):
                    for rect in graph.reference_rects(cell.sheet, inner.ref):
                        out.readers.append((rect, cell))


def _name_ref(graph: WorkbookGraph, cell: CellRef, ref: str, out: HiddenReads) -> None:
    found = graph.lookup_name(ref, cell.sheet)
    if found is None:
        if _unresolved_name(ref):
            out.every_sheet = True  # a name form calc cannot resolve: it may read anything
        return
    if found.relative:
        # Read relative to each cell using it: somewhere on its sheets.
        sheets = {rect[0] for rect in found.rects}
        if sheets:
            out.sheets.update(sheets)
        else:
            out.every_sheet = True
        return
    if found.kind == "cells":
        # Resolved as Excel resolves it (the formula's own sheet first, or the sheet a
        # ``Sheet!Name`` names): the IR has no edge for every such form.
        out.readers.extend((rect, cell) for rect in found.rects)
        return
    if found.kind != "formula":
        return  # cells in another workbook: none of this one's
    rects = _name_text_rects(graph, found, 0)
    if rects is None:
        out.every_sheet = True
    else:
        out.readers.extend((rect, cell) for rect in rects)


def _unresolved_name(ref: str) -> bool:
    """True for ``Sheet!Name`` that no definition answers (a cell, range, whole
    column or row, or another workbook's reference is not a name)."""
    head, bang, tail = ref.rpartition("!")
    if not bang or head.lstrip("'").startswith("[") or not tail:
        return False
    if _A1_TEXT.fullmatch(tail) or _WHOLE.fullmatch(tail):
        return False
    return _NAME.fullmatch(tail) is not None


def _name_text_rects(graph: WorkbookGraph, found: NameDef, depth: int) -> list[Rect] | None:
    """The cells a formula-holding name reads, from its text; None when that cannot be
    told (no text in the IR, a dynamic or opaque call, a chain too deep)."""
    if depth > _MAX_DEPTH or not found.areas:
        return None
    rects: list[Rect] = []
    for text in found.areas:
        if text.strip().upper() in ERROR_CODES:
            continue
        root = parse_formula(text)
        if root is None:
            return None
        for node in iter_nodes(root):
            if isinstance(node, FunctionNode) and function_name(node.name) in (
                *_LANDING,
                *_OPAQUE,
            ):
                return None
            if not isinstance(node, RefNode):
                continue
            if _SHEET_PREFIX.match(node.ref) or _A1_TEXT.fullmatch(node.ref):
                got = graph.reference_rects(found.scope or "", node.ref)
                if not got:
                    return None
                rects.extend(got)
                continue
            inner = graph.lookup_name(node.ref, found.scope)
            if inner is None or inner.relative:
                return None
            if inner.kind == "cells":
                rects.extend(inner.rects)
            elif inner.kind == "formula":
                deeper = _name_text_rects(graph, inner, depth + 1)
                if deeper is None:
                    return None
                rects.extend(deeper)
    return rects


def _sheets_of(graph: WorkbookGraph, ref: str, own: str) -> set[str] | None:
    """The sheets a reference (or ``INDIRECT`` text) written on ``own`` points into;
    None when that cannot be told."""
    ref = ref.strip()
    m = _SHEET_PREFIX.match(ref)
    if m is not None:
        written = m.group(1)
        if written.startswith("'") and written.endswith("'") and len(written) >= 2:
            written = written[1:-1].replace("''", "'")
        sheet = graph.sheet(written)
        return {sheet} if sheet is not None else set()  # an unknown sheet reads #REF!
    if _A1_TEXT.fullmatch(ref):
        return {own}
    found = graph.lookup_name(ref, own)
    if found is None or found.kind != "cells" or not found.rects:
        return None
    return {rect[0] for rect in found.rects}
