"""Static scan for formulas that match Excel on the stored inputs only by luck.

A formula found here is refused even when the engine reproduces its stored value:
an approximate lookup over unsorted keys, an implicit intersection, a random or
clock function, a circular reference or a data table can agree with Excel on one
set of inputs and silently disagree on the next. When in doubt the scan refuses.
"""

from __future__ import annotations

import re
from collections import defaultdict
from collections.abc import Callable, Collection, Iterator
from dataclasses import dataclass, field
from datetime import date
from typing import Literal

from loguru import logger

from xl_marinade.calc.formula_ast import (
    ALWAYS_ARRAY,
    ENGINE_LACKS,
    ERROR_CODES,
    SHEET,
    UNKNOWN,
    function_name,
    literal,
    parse_formula,
    sign,
)
from xl_marinade.calc.graph import NameDef, StoredValue, WorkbookGraph
from xl_marinade.calc.models import CellRef, Reason, ReasonKind, _col_to_number
from xl_marinade.core.levels_cycles import detect_cycles
from xl_marinade.core.parser import (
    ASTNode,
    BinaryNode,
    FunctionNode,
    RefNode,
    UnaryNode,
)
from xl_marinade.errors import CalcInputError

UNSORTED = "approximate lookup on unsorted column"
MIXED = "approximate lookup on mixed text and numbers"
NON_RANGE = "approximate lookup on a non-range table"
WHOLE_LINE = "approximate lookup on whole column"
IMPLICIT = "implicit intersection on multi-cell name"
IMPLICIT_RANGE = "implicit intersection on multi-cell range"
NAME_FORMULA = "name holds a formula"
CIRCULAR = "circular reference"
DATA_TABLE = "data table"
UNPARSEABLE = "unparseable formula"
NOT_HELD = "external reference not held"
INPUT_KEYS = "approximate lookup on input-dependent keys"
RELATIVE_NAME = "name with a relative reference"
ENGINE_LACKS_PREFIX = "function IronCalc does not implement: "
SHEETS_OF_REFERENCE = f"{ENGINE_LACKS_PREFIX}SHEETS with a reference"
"""IronCalc answers #N/IMPL! for SHEETS with an argument (SHEETS() works)."""

_MAX_ROW = 1_048_576
_MAX_COL = 16_384
_RANDOM = frozenset({"RAND", "RANDBETWEEN", "RANDARRAY"})
_CLOCK = frozenset({"TODAY", "NOW"})
# Text keys made only of these characters sort the same in Python and in Excel's
# collation (which ignores hyphens/apostrophes, puts ``_`` before digits and files
# accented letters beside their base letter).
_SAFE_TEXT = re.compile(r"^[a-z0-9 ]*$")

_CELL_RANGE = re.compile(
    rf"^(?:(?P<s1>{SHEET})!)?\$?(?P<c1>[A-Za-z]{{1,3}})\$?(?P<r1>[0-9]+)"
    rf"(?::(?:(?P<s2>{SHEET})!)?\$?(?P<c2>[A-Za-z]{{1,3}})\$?(?P<r2>[0-9]+))?$"
)
_COL_RANGE = re.compile(
    rf"^(?:(?P<s1>{SHEET})!)?\$?(?P<c1>[A-Za-z]{{1,3}}):"
    rf"(?:(?P<s2>{SHEET})!)?\$?(?P<c2>[A-Za-z]{{1,3}})$"
)
_ROW_RANGE = re.compile(
    rf"^(?:(?P<s1>{SHEET})!)?\$?(?P<r1>[0-9]+):(?:(?P<s2>{SHEET})!)?\$?(?P<r2>[0-9]+)$"
)

Order = Literal["asc", "desc"]
Axis = Literal["col", "row", "auto"]

# (sheet, r1, c1, r2, c2)
_Rect = tuple[str, int, int, int, int]


def _unquote(sheet: str) -> str:
    if len(sheet) >= 2 and sheet.startswith("'") and sheet.endswith("'"):
        return sheet[1:-1].replace("''", "'")
    return sheet


def _walk(root: ASTNode) -> Iterator[tuple[ASTNode, bool]]:
    """Every node with a flag: True when no implicit intersection can apply to it.

    That is a node directly a function argument, or anywhere inside the arguments
    of a function that array-evaluates in every Excel version (``SUMPRODUCT``...).
    """
    stack: list[tuple[ASTNode, bool, bool]] = [(root, False, False)]
    while stack:
        node, direct, in_array = stack.pop()
        yield node, direct or in_array
        if isinstance(node, FunctionNode):
            inner = in_array or function_name(node.name) in ALWAYS_ARRAY
            stack.extend((arg, True, inner) for arg in reversed(node.args))
        elif isinstance(node, BinaryNode):
            stack.extend([(node.right, False, in_array), (node.left, False, in_array)])
        elif isinstance(node, UnaryNode):
            stack.append((node.operand, False, in_array))


def _lookup_spec(name: str, args: list[ASTNode]) -> tuple[ASTNode, Axis, set[Order]] | None:
    """For an approximate lookup: its key table, key axis and required orders."""
    if name in {"VLOOKUP", "HLOOKUP"} and len(args) >= 2:
        if len(args) >= 4 and sign(literal(args[3])) == 0:
            return None
        return args[1], ("col" if name == "VLOOKUP" else "row"), {"asc"}
    if name == "MATCH" and len(args) >= 2:
        direction = 1 if len(args) < 3 else sign(literal(args[2]))
        if direction == 0:
            return None
        orders: set[Order] = (
            {"asc", "desc"} if direction is None else {"asc" if direction > 0 else "desc"}
        )
        return args[1], "auto", orders
    if name == "LOOKUP" and len(args) >= 2:
        return args[1], "auto", {"asc"}
    search_at = {"XLOOKUP": 5, "XMATCH": 3}.get(name)
    if search_at is not None and len(args) >= 2:
        if len(args) <= search_at:
            return None
        mode = literal(args[search_at])
        if mode in (2, 2.0):
            return args[1], "auto", {"asc"}
        if mode in (-2, -2.0):
            return args[1], "auto", {"desc"}
        if mode is UNKNOWN or isinstance(mode, str):
            return args[1], "auto", {"asc", "desc"}
    return None


def _same_sheet(m: re.Match[str], default: str) -> str | None:
    sheet = _unquote(m.group("s1")) if m.group("s1") else default
    second = m.group("s2")
    if second is not None and _unquote(second) != sheet:
        return None
    return sheet


@dataclass(slots=True)
class _Scan:
    """Per-scan state: the graph plus caches, so copied-down formulas stay cheap."""

    graph: WorkbookGraph
    as_of: date | None
    names: dict[tuple[str, str], NameDef | None] = field(default_factory=dict)
    lookups: dict[tuple[str, str, Axis, frozenset[Order]], str | None] = field(default_factory=dict)
    calls: dict[CellRef, set[str]] = field(default_factory=dict)
    """Each formula cell's function names as written, for the engine to be asked about."""


def _multi_cell(ref: str, default_sheet: str) -> bool:
    parsed = _ref_rect(ref, default_sheet)
    if parsed is None:
        return False
    _sheet, r1, c1, r2, c2 = parsed[0]
    return (r1, c1) != (r2, c2)


def _name_def(scan: _Scan, cell: CellRef, ref: str) -> NameDef | None:
    """The definition ``ref`` (``name`` or ``Sheet!name``) reads in ``cell``'s formula,
    resolved from ``cell``'s sheet as Excel does; None if it is not a name."""
    key = (cell.sheet, ref)
    if key not in scan.names:
        scan.names[key] = scan.graph.lookup_name(ref, cell.sheet)
    return scan.names[key]


def _multi_cell_name(found: NameDef) -> bool:
    """True for a name over more than one cell (a whole column or row always is)."""
    if found.whole or len(found.rects) > 1:
        return True
    return any((r1, c1) != (r2, c2) for _s, r1, c1, r2, c2 in found.rects)


def _ref_rect(ref: str, default_sheet: str) -> tuple[_Rect, bool] | None:
    """A cell, range, whole-column or whole-row reference as a rectangle.

    The flag is True for whole columns/rows. None for anything else (names,
    structured references, references spanning two sheets).
    """
    if (m := _CELL_RANGE.match(ref)) is not None:
        sheet = _same_sheet(m, default_sheet)
        if sheet is None:
            return None
        r1, c1 = int(m.group("r1")), _col_to_number(m.group("c1"))
        r2 = int(m.group("r2")) if m.group("r2") else r1
        c2 = _col_to_number(m.group("c2")) if m.group("c2") else c1
        return (sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)), False
    if (m := _COL_RANGE.match(ref)) is not None:
        sheet = _same_sheet(m, default_sheet)
        if sheet is None:
            return None
        a, b = sorted((_col_to_number(m.group("c1")), _col_to_number(m.group("c2"))))
        return (sheet, 1, a, _MAX_ROW, b), True
    if (m := _ROW_RANGE.match(ref)) is not None:
        sheet = _same_sheet(m, default_sheet)
        if sheet is None:
            return None
        a, b = sorted((int(m.group("r1")), int(m.group("r2"))))
        return (sheet, a, 1, b, _MAX_COL), True
    return None


def _contains(rect: _Rect, cell: CellRef) -> bool:
    sheet, r1, c1, r2, c2 = rect
    return cell.sheet == sheet and r1 <= cell.row <= r2 and c1 <= cell.col <= c2


def _refers_to_self(scan: _Scan, cell: CellRef, ref: str) -> bool:
    """True when the reference covers the formula's own cell (the IR drops self-edges)."""
    found = _name_def(scan, cell, ref)
    if found is not None:
        return found.kind == "cells" and any(_contains(r, cell) for r in found.rects)
    parsed = _ref_rect(ref, cell.sheet)
    return parsed is not None and _contains(parsed[0], cell)


def _table_cells(scan: _Scan, cell: CellRef, ref: str) -> tuple[_Rect, list[CellRef]] | None:
    """The table's bounding rectangle and the cells to read; None if not a plain range."""
    graph = scan.graph
    found = _name_def(scan, cell, ref)
    if found is not None:
        if found.kind != "cells" or not found.rects:
            return None
        if found.whole:
            if len(found.rects) != 1:
                return None
            # A whole column/row name: the IR's range edges list every stored cell read.
            inside = [p for p in graph.precedents(cell) if _contains(found.rects[0], p)]
            return found.rects[0], sorted(inside)
        named = graph.name_cells(ref, cell.sheet)
        rows = [c.row for c in named]
        cols = [c.col for c in named]
        rect = (named[0].sheet, min(rows), min(cols), max(rows), max(cols))
        area = (rect[3] - rect[1] + 1) * (rect[4] - rect[2] + 1)
        if len({c.sheet for c in named}) != 1 or len(named) != area:
            return None
        return rect, named

    parsed = _ref_rect(ref, cell.sheet)
    if parsed is None:
        return None
    rect, whole = parsed
    if whole:
        # Whole columns/rows: the IR's range edges already list every stored cell read.
        inside = [p for p in graph.precedents(cell) if _contains(rect, p)]
        return rect, sorted(inside)
    sheet, r1, c1, r2, c2 = rect
    corner = CellRef(sheet, r2, c2).a1().rpartition("!")[2]
    try:
        return rect, graph.resolve_target(f"{CellRef(sheet, r1, c1).a1()}:{corner}")
    except CalcInputError:
        return None


def _key_cells(rect: _Rect, cells: list[CellRef], axis: Axis) -> list[CellRef]:
    _sheet, r1, c1, r2, c2 = rect
    if axis == "auto":
        # LOOKUP/MATCH on a 2-D table search the first row when it is wider than tall.
        axis = "row" if (c2 - c1) > (r2 - r1) else "col"
    if axis == "col":
        return sorted((c for c in cells if c.col == c1), key=lambda c: c.row)
    return sorted((c for c in cells if c.row == r1), key=lambda c: c.col)


def _kind(value: StoredValue) -> str:
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, float):
        return "number"
    if isinstance(value, str) and value in ERROR_CODES:
        return "error"
    return "text"


def _ordered(values: list[float | str | bool], order: Order) -> bool:
    pairs = zip(values, values[1:], strict=False)
    if order == "asc":
        return all(a <= b for a, b in pairs)  # type: ignore[operator]
    return all(a >= b for a, b in pairs)  # type: ignore[operator]


def _key_problem(graph: WorkbookGraph, keys: list[CellRef], orders: set[Order]) -> str | None:
    values = [v for c in keys if (v := graph.stored(c)) is not None and v != ""]
    kinds = {_kind(v) for v in values}
    if "error" in kinds:
        return UNSORTED
    if len(kinds) > 1:
        return MIXED
    norm: list[float | str | bool] = [v.casefold() if isinstance(v, str) else v for v in values]
    if any(isinstance(v, str) and not _SAFE_TEXT.match(v) for v in norm):
        return UNSORTED
    if all(_ordered(norm, order) for order in orders):
        return None
    return UNSORTED


def _lookup_problem(
    scan: _Scan, cell: CellRef, table: ASTNode, axis: Axis, orders: set[Order]
) -> str | None:
    """Why an approximate lookup over ``table`` is unsafe; None when its keys are sorted."""
    if not isinstance(table, RefNode):
        return NON_RANGE
    # Whole-column tables read the same IR cells from any formula on the sheet.
    key = (cell.sheet, table.ref, axis, frozenset(orders))
    if key in scan.lookups:
        return scan.lookups[key]
    parsed = _ref_rect(table.ref, cell.sheet)
    found = _name_def(scan, cell, table.ref)
    if (parsed is not None and parsed[1]) or (found is not None and found.whole):
        # IronCalc 0.9.1's approximate MATCH over a whole column/row returns the last
        # row (1048576) where Excel returns the last key at or below the value.
        scan.lookups[key] = WHOLE_LINE
        return WHOLE_LINE
    resolved = _table_cells(scan, cell, table.ref)
    problem = (
        NON_RANGE
        if resolved is None
        else _key_problem(scan.graph, _key_cells(resolved[0], resolved[1], axis), orders)
    )
    scan.lookups[key] = problem
    return problem


def _formula_reasons(scan: _Scan, cell: CellRef, text: str) -> set[str]:
    root = parse_formula(text)
    if root is None:
        return {UNPARSEABLE}
    found: set[str] = set()
    for node, exempt in _walk(root):
        if isinstance(node, RefNode):
            if _refers_to_self(scan, cell, node.ref):
                found.add(CIRCULAR)
            definition = _name_def(scan, cell, node.ref)
            kind = definition.kind if definition is not None else None
            if kind == "external":
                # A name over another workbook's cells: the engine copy holds no cached
                # value for it, so off the stored inputs it can only be wrong.
                found.add(NOT_HELD)
            elif kind == "formula":
                found.add(NAME_FORMULA)
            if definition is not None and definition.relative:
                # Excel reads it relative to this cell; the engine reads the stored text.
                found.add(f"{RELATIVE_NAME}: {definition.name}")
            if exempt or scan.graph.is_array_formula(cell):
                continue
            if definition is not None:
                if kind == "cells" and _multi_cell_name(definition):
                    found.add(IMPLICIT)
            elif _multi_cell(node.ref, cell.sheet):
                found.add(IMPLICIT_RANGE)
            continue
        if not isinstance(node, FunctionNode):
            continue
        scan.calls.setdefault(cell, set()).add(node.name)
        name = function_name(node.name)
        if name == "SHEETS" and node.args:
            found.add(SHEETS_OF_REFERENCE)
        if name in ENGINE_LACKS:
            # The engine's #NAME? is an Excel error code: an error handler around the
            # call turns it into a plausible value no trust run can tell from Excel's.
            found.add(f"{ENGINE_LACKS_PREFIX}{name}")
        if name in _RANDOM:
            found.add(name)
        if name in _CLOCK and scan.as_of is None:
            found.add(name)
        spec = _lookup_spec(name, node.args)
        if spec is None:
            continue
        problem = _lookup_problem(scan, cell, *spec)
        if problem is not None:
            found.add(problem)
    return found


def _cycle_cells(graph: WorkbookGraph) -> set[CellRef]:
    """Formula cells on a circular reference (self-references included)."""
    formulas = set(graph.formula_cells())
    succ: dict[CellRef, set[CellRef]] = {
        c: {p for p in graph.precedents(c) if p in formulas} for c in formulas
    }
    pred: dict[CellRef, set[CellRef]] = defaultdict(set)
    for c, ps in succ.items():
        for p in ps:
            pred[p].add(c)
    # Trim cells with no in- or out-edge left: they cannot be on a cycle. This keeps
    # the (recursive) Tarjan input small on long acyclic projection chains.
    alive = set(formulas)
    queue = [c for c in alive if not succ[c] or not pred[c]]
    while queue:
        c = queue.pop()
        if c not in alive:
            continue
        alive.discard(c)
        for n in succ[c] | pred[c]:
            succ[n].discard(c)
            pred[n].discard(c)
            if n in alive and (not succ[n] or not pred[n]):
                queue.append(n)
    if not alive:
        return set()
    ids = {c.a1(): c for c in alive}
    edges = sorted((c.a1(), p.a1()) for c in alive for p in succ[c])
    try:
        cycles = detect_cycles(sorted(ids), edges)
    except RecursionError:
        logger.warning("cycle scan too deep; refusing all {} cells left after trimming", len(ids))
        return alive
    return {ids[b] for cyc in cycles for b in cyc.bindings}


def input_dependent_lookups(
    graph: WorkbookGraph, inputs: set[CellRef]
) -> dict[CellRef, list[Reason]]:
    """Approximate lookups whose key cells change with the mapped inputs.

    The scan checks key order on the stored values only; keys computed from an input
    can be sorted for one policy and not the next, so such a lookup is refused.
    Returns ``refused_pattern`` reasons like :func:`scan_patterns`.
    """
    moving = set(graph.downstream(inputs))
    scan = _Scan(graph, None)
    found: dict[CellRef, list[Reason]] = {}
    # A lookup over moving keys reads them, so it is downstream of the inputs too.
    for cell in sorted(c for c in moving if graph.formula(c) is not None):
        root = parse_formula(graph.formula(cell) or "")
        if root is None:
            continue
        for node, _exempt in _walk(root):
            if not isinstance(node, FunctionNode):
                continue
            spec = _lookup_spec(function_name(node.name), node.args)
            if spec is None or not isinstance(spec[0], RefNode):
                continue
            resolved = _table_cells(scan, cell, spec[0].ref)
            if resolved is None:
                continue
            if any(k in moving for k in _key_cells(resolved[0], resolved[1], spec[1])):
                found[cell] = [Reason(ReasonKind.refused_pattern, cell, INPUT_KEYS)]
                break
    return found


def scan_patterns(
    graph: WorkbookGraph,
    *,
    as_of: date | None,
    data_table_cells: list[CellRef],
    engine_lacks: Callable[[Collection[str]], frozenset[str]] | None = None,
) -> dict[CellRef, list[Reason]]:
    """Find cells whose formulas must be refused even when they reproduce.

    ``engine_lacks`` is asked once about every function name the formulas call and
    returns those the installed engine cannot compute (``engine.functions_engine_lacks``);
    each call to one is refused by name, as :data:`ENGINE_LACKS` is. Without it only
    that fixed list is used.

    Returns a map from cell to its ``refused_pattern`` reasons (root = the cell),
    sorted and without duplicates; cells with no reason are absent.
    """
    scan = _Scan(graph, as_of)
    details: dict[CellRef, set[str]] = defaultdict(set)
    for cell in graph.formula_cells():
        text = graph.formula(cell)
        if text is None:
            continue
        found = _formula_reasons(scan, cell, text)
        if found:
            details[cell].update(found)
    if engine_lacks is not None and scan.calls:
        # The engine's #NAME? or #N/IMPL! is an Excel error code: an error handler
        # around the call turns it into a plausible value on every input.
        lacking = engine_lacks({n for names in scan.calls.values() for n in names})
        for cell, names in scan.calls.items():
            for raw in names:
                if function_name(raw) in lacking:
                    details[cell].add(f"{ENGINE_LACKS_PREFIX}{function_name(raw)}")
    for cell in _cycle_cells(graph):
        details[cell].add(CIRCULAR)
    for cell in data_table_cells:
        details[cell].add(DATA_TABLE)
    return {
        cell: sorted(
            (Reason(ReasonKind.refused_pattern, cell, d) for d in details[cell]),
            key=lambda r: (r.kind, r.root or CellRef("", 0, 0), r.detail),
        )
        for cell in sorted(details)
    }
