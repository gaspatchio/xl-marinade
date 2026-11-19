"""Static scan for formulas that match Excel on the stored inputs only by luck.

A formula found here is refused even when the engine reproduces its stored value:
an approximate lookup over unsorted keys, an implicit intersection, a random or
clock function, a circular reference or a data table can agree with Excel on one
set of inputs and silently disagree on the next. When in doubt the scan refuses.
"""

from __future__ import annotations

import re
from collections import defaultdict
from collections.abc import Iterator
from dataclasses import dataclass, field
from datetime import date
from typing import Literal

from loguru import logger

from xl_marinade.calc.graph import StoredValue, WorkbookGraph
from xl_marinade.calc.models import CellRef, Reason, ReasonKind, _col_to_number
from xl_marinade.core.levels_cycles import detect_cycles
from xl_marinade.core.parser import (
    ASTNode,
    BinaryNode,
    ConstNode,
    FormulaParser,
    FunctionNode,
    RefNode,
    UnaryNode,
)
from xl_marinade.errors import CalcInputError

UNSORTED = "approximate lookup on unsorted column"
MIXED = "approximate lookup on mixed text and numbers"
NON_RANGE = "approximate lookup on a non-range table"
IMPLICIT = "implicit intersection on multi-cell name"
IMPLICIT_RANGE = "implicit intersection on multi-cell range"
NAME_FORMULA = "name holds a formula"
CIRCULAR = "circular reference"
DATA_TABLE = "data table"
UNPARSEABLE = "unparseable formula"

_MAX_ROW = 1_048_576
_MAX_COL = 16_384
_FN_PREFIXES = ("_XLFN.", "_XLWS.", "_XLL.")
_RANDOM = frozenset({"RAND", "RANDBETWEEN", "RANDARRAY"})
_CLOCK = frozenset({"TODAY", "NOW"})
_ERRORS = frozenset(
    {
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#GETTING_DATA",
        "#SPILL!",
        "#CALC!",
        "#FIELD!",
        "#BLOCKED!",
        "#CONNECT!",
        "#BUSY!",
        "#UNKNOWN!",
    }
)

# Text keys made only of these characters sort the same in Python and in Excel's
# collation (which ignores hyphens/apostrophes, puts ``_`` before digits and files
# accented letters beside their base letter).
_SAFE_TEXT = re.compile(r"^[a-z0-9 ]*$")

_SHEET = r"(?:'(?:[^']|'')+'|[^'!:\s{}();,\"#\[\]]+)"
_CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]+"
_NAME = r"[A-Za-z_\\][\w.\\?]*"
_END = rf"(?:{_CELL}|\$?[A-Za-z]{{1,3}}|\$?[0-9]+|{_NAME})"
# What a well-formed RefNode may look like; anything else is a mangled parse.
_VALID_REF = re.compile(rf"^(?:{_SHEET}!)?{_END}(?::(?:{_SHEET}!)?{_END})?$")
_STRUCTURED = re.compile(r"^[A-Za-z_\\][\w.\\]*\[.*\]$")
_CELL_RANGE = re.compile(
    rf"^(?:(?P<s1>{_SHEET})!)?\$?(?P<c1>[A-Za-z]{{1,3}})\$?(?P<r1>[0-9]+)"
    rf"(?::(?:(?P<s2>{_SHEET})!)?\$?(?P<c2>[A-Za-z]{{1,3}})\$?(?P<r2>[0-9]+))?$"
)
_COL_RANGE = re.compile(
    rf"^(?:(?P<s1>{_SHEET})!)?\$?(?P<c1>[A-Za-z]{{1,3}}):"
    rf"(?:(?P<s2>{_SHEET})!)?\$?(?P<c2>[A-Za-z]{{1,3}})$"
)
_ROW_RANGE = re.compile(
    rf"^(?:(?P<s1>{_SHEET})!)?\$?(?P<r1>[0-9]+):(?:(?P<s2>{_SHEET})!)?\$?(?P<r2>[0-9]+)$"
)

Order = Literal["asc", "desc"]
Axis = Literal["col", "row", "auto"]
_UNKNOWN = object()

# (sheet, r1, c1, r2, c2)
_Rect = tuple[str, int, int, int, int]


def _unquote(sheet: str) -> str:
    if len(sheet) >= 2 and sheet.startswith("'") and sheet.endswith("'"):
        return sheet[1:-1].replace("''", "'")
    return sheet


def _function_name(raw: str) -> str:
    """Upper-case name with ``_xlfn.`` / ``_xlws.`` / ``_xll.`` prefixes removed."""
    name = raw.upper()
    stripped = True
    while stripped:
        stripped = False
        for prefix in _FN_PREFIXES:
            if name.startswith(prefix):
                name = name[len(prefix) :]
                stripped = True
    return name


def _parse(text: str) -> ASTNode | None:
    """Parse the formula; None when the parser raised or the parse looks mangled.

    The core parser silently stops at text it cannot place (an intersection space,
    ``INDEX(...):B3``) and turns array constants into garbage references, so a
    parse that left text behind or holds an implausible reference is rejected too.
    """
    parser = FormulaParser()
    try:
        root = parser.parse(text)
    except (ValueError, IndexError, RecursionError):
        return None
    if parser.formula[parser.pos :].strip():
        return None
    for node, _direct in _walk(root):
        if isinstance(node, RefNode) and not _plausible_ref(node.ref):
            return None
    return root


def _plausible_ref(ref: str) -> bool:
    if _VALID_REF.match(ref):
        return True
    return bool(_STRUCTURED.match(ref)) and ref.count("[") == ref.count("]")


def _walk(root: ASTNode) -> Iterator[tuple[ASTNode, bool]]:
    """Every node with a flag: True when it is directly a function argument."""
    stack: list[tuple[ASTNode, bool]] = [(root, False)]
    while stack:
        node, direct = stack.pop()
        yield node, direct
        if isinstance(node, FunctionNode):
            stack.extend((arg, True) for arg in reversed(node.args))
        elif isinstance(node, BinaryNode):
            stack.extend([(node.right, False), (node.left, False)])
        elif isinstance(node, UnaryNode):
            stack.append((node.operand, False))


def _literal(node: ASTNode) -> object:
    """The value of a literal argument (``TRUE``, ``0``, ``-1``), else ``_UNKNOWN``."""
    if isinstance(node, ConstNode):
        return node.value
    if isinstance(node, UnaryNode) and node.operator in {"+", "-"}:
        inner = _literal(node.operand)
        if isinstance(inner, int | float) and not isinstance(inner, bool):
            return -inner if node.operator == "-" else inner
    return _UNKNOWN


def _sign(value: object) -> int | None:
    """Sign of a numeric/boolean literal; None for text or a non-literal."""
    if isinstance(value, bool):
        return 1 if value else 0
    if isinstance(value, int | float):
        return (value > 0) - (value < 0)
    return None


def _lookup_spec(name: str, args: list[ASTNode]) -> tuple[ASTNode, Axis, set[Order]] | None:
    """For an approximate lookup: its key table, key axis and required orders."""
    if name in {"VLOOKUP", "HLOOKUP"} and len(args) >= 2:
        if len(args) >= 4 and _sign(_literal(args[3])) == 0:
            return None
        return args[1], ("col" if name == "VLOOKUP" else "row"), {"asc"}
    if name == "MATCH" and len(args) >= 2:
        sign = 1 if len(args) < 3 else _sign(_literal(args[2]))
        if sign == 0:
            return None
        orders: set[Order] = {"asc", "desc"} if sign is None else {"asc" if sign > 0 else "desc"}
        return args[1], "auto", orders
    if name == "LOOKUP" and len(args) >= 2:
        return args[1], "auto", {"asc"}
    search_at = {"XLOOKUP": 5, "XMATCH": 3}.get(name)
    if search_at is not None and len(args) >= 2:
        if len(args) <= search_at:
            return None
        mode = _literal(args[search_at])
        if mode in (2, 2.0):
            return args[1], "auto", {"asc"}
        if mode in (-2, -2.0):
            return args[1], "auto", {"desc"}
        if mode is _UNKNOWN or isinstance(mode, str):
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
    names: dict[str, tuple[list[CellRef], frozenset[CellRef]]] = field(default_factory=dict)
    lookups: dict[tuple[str, str, Axis, frozenset[Order]], str | None] = field(default_factory=dict)


def _name_candidates(ref: str) -> list[str]:
    """``ref`` itself and, for ``Sheet!name``, the bare name."""
    return list(dict.fromkeys([ref, ref.rpartition("!")[2]]))


def _multi_cell(ref: str, default_sheet: str) -> bool:
    parsed = _ref_rect(ref, default_sheet)
    if parsed is None:
        return False
    _sheet, r1, c1, r2, c2 = parsed[0]
    return (r1, c1) != (r2, c2)


def _named_cells(scan: _Scan, ref: str) -> tuple[list[CellRef], frozenset[CellRef]]:
    """Cells of a defined name written as ``name`` or ``Sheet!name``; empty if not a name."""
    hit = scan.names.get(ref)
    if hit is not None:
        return hit
    result: tuple[list[CellRef], frozenset[CellRef]] = ([], frozenset())
    for candidate in _name_candidates(ref):
        named = scan.graph.name_cells(candidate)
        if named:
            result = (named, frozenset(named))
            break
    scan.names[ref] = result
    return result


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
    if cell in _named_cells(scan, ref)[1]:
        return True
    parsed = _ref_rect(ref, cell.sheet)
    return parsed is not None and _contains(parsed[0], cell)


def _table_cells(scan: _Scan, cell: CellRef, ref: str) -> tuple[_Rect, list[CellRef]] | None:
    """The table's bounding rectangle and the cells to read; None if not a plain range."""
    graph = scan.graph
    named = _named_cells(scan, ref)[0]
    if named:
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
    if isinstance(value, str) and value in _ERRORS:
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
    resolved = _table_cells(scan, cell, table.ref)
    problem = (
        NON_RANGE
        if resolved is None
        else _key_problem(scan.graph, _key_cells(resolved[0], resolved[1], axis), orders)
    )
    scan.lookups[key] = problem
    return problem


def _formula_reasons(scan: _Scan, cell: CellRef, text: str) -> set[str]:
    root = _parse(text)
    if root is None:
        return {UNPARSEABLE}
    found: set[str] = set()
    for node, direct in _walk(root):
        if isinstance(node, RefNode):
            if _refers_to_self(scan, cell, node.ref):
                found.add(CIRCULAR)
            if any(scan.graph.is_formula_name(n) for n in _name_candidates(node.ref)):
                found.add(NAME_FORMULA)
            if direct or scan.graph.is_array_formula(cell):
                continue
            named = _named_cells(scan, node.ref)[0]
            if len(named) > 1:
                found.add(IMPLICIT)
            elif not named and _multi_cell(node.ref, cell.sheet):
                found.add(IMPLICIT_RANGE)
            continue
        if not isinstance(node, FunctionNode):
            continue
        name = _function_name(node.name)
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


def scan_patterns(
    graph: WorkbookGraph, *, as_of: date | None, data_table_cells: list[CellRef]
) -> dict[CellRef, list[Reason]]:
    """Find cells whose formulas must be refused even when they reproduce.

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
