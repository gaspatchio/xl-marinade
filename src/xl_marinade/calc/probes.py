"""Probe cells: formulas on a hidden sheet that report which path a policy took.

For each formula cell the plan lifts parts of its formula onto ``_marinade_probes``
so the engine computes, per policy, which branch each ``IF``-like function took,
where each lookup landed, the inputs to rounding functions and branch
comparisons, and where ``OFFSET`` / ``INDIRECT`` point. Probes are per cell; each
records its formula family and its ordinal (index among the cell's probe nodes of
that kind, in AST pre-order) so readings can be aggregated per family.

Lifted text must compute the same on the probe sheet as in its own cell, so every
reference is sheet-qualified and absolute (:func:`anchor`). Anything that cannot
be lifted faithfully, and every probe beyond the cap, is reported as unplaced.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from enum import StrEnum
from typing import Protocol

from loguru import logger

from xl_marinade.calc.engine import Engine, EngineValue
from xl_marinade.calc.formula_ast import (
    ALWAYS_ARRAY,
    END,
    ERROR_CODES,
    SHEET,
    function_name,
    literal,
    parse_formula,
    sign,
)
from xl_marinade.calc.graph import WorkbookGraph
from xl_marinade.calc.models import CellRef, _col_to_number
from xl_marinade.core.parser import (
    ASTNode,
    BinaryNode,
    ConstNode,
    FunctionNode,
    RefNode,
    UnaryNode,
)
from xl_marinade.core.ref_converter import quote_sheet_name

PROBE_SHEET = "_marinade_probes"
NEAR_TOLERANCE = 1e-12

_COMPARISONS = frozenset({"<", "<=", ">", ">=", "=", "<>"})
_NEAR_COMPARE = _COMPARISONS | {"lookup"}
# Functions that bind names: nothing beneath them can be lifted on its own.
_SCOPING = frozenset({"LET", "LAMBDA", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY"})
# Functions whose result depends on the cell holding the formula when called bare.
_POSITIONAL = frozenset({"ROW", "COLUMN", "SHEET"})
_ROUNDING: dict[str, str] = {
    "ROUND": "ROUND",
    "ROUNDUP": "ROUNDUP",
    "ROUNDDOWN": "ROUNDDOWN",
    "MROUND": "MROUND",
    "INT": "INT",
    "TRUNC": "TRUNC",
    "CEILING": "CEILING",
    "CEILING.MATH": "CEILING",
    "CEILING.PRECISE": "CEILING",
    "ISO.CEILING": "CEILING",
    "FLOOR": "FLOOR",
    "FLOOR.MATH": "FLOOR",
    "FLOOR.PRECISE": "FLOOR",
}
_BRANCHES = frozenset({"IF", "IFS", "SWITCH", "CHOOSE", "IFERROR", "IFNA"})
_LOOKUPS = frozenset({"VLOOKUP", "HLOOKUP", "MATCH", "LOOKUP"})
_DYNAMIC = frozenset({"OFFSET", "INDIRECT"})

_REF = re.compile(rf"^(?:(?P<s1>{SHEET})!)?(?P<a>{END})(?::(?:(?P<s2>{SHEET})!)?(?P<b>{END}))?$")
_CELL_END = re.compile(r"^\$?(?P<col>[A-Za-z]{1,3})\$?(?P<row>[0-9]+)$")
_COL_END = re.compile(r"^\$?(?P<col>[A-Za-z]{1,3})$")
_ROW_END = re.compile(r"^\$?(?P<row>[0-9]+)$")
_TEXT_REF = re.compile(
    rf"^(?:(?P<sheet>{SHEET})!)?\$?(?P<c1>[A-Za-z]{{1,3}})\$?(?P<r1>[0-9]+)"
    r"(?::\$?(?P<c2>[A-Za-z]{1,3})\$?(?P<r2>[0-9]+))?$"
)


class ProbeKind(StrEnum):
    """What a probe reports."""

    branch = "branch"
    lookup = "lookup"
    rounding = "rounding"
    comparison = "comparison"
    dynamic_ref = "dynamic_ref"


@dataclass(slots=True)
class Probe:
    """One probe: the formulas lifted from one node of one cell's formula.

    ``meta`` always holds ``function`` (branch, lookup, rounding, dynamic_ref) or
    ``operator`` (comparison: the operator, or ``lookup`` for an approximate
    lookup's value against its matched key).
    """

    kind: ProbeKind
    cell: CellRef
    family: str
    ordinal: int
    formulas: tuple[str, ...]
    meta: dict[str, str]


@dataclass(slots=True)
class ProbePlan:
    """The probes to place, in order, and per cell the kinds that could not be placed."""

    probes: list[Probe]
    unplaced: dict[CellRef, list[ProbeKind]]


@dataclass(slots=True)
class ProbeReadings:
    """What the probes computed for one policy."""

    branch: dict[tuple[CellRef, int], str]
    lookup_edge: dict[tuple[CellRef, int], str]
    near_rounding: set[CellRef]
    near_comparison: set[CellRef]
    dynamic_refs: dict[CellRef, list[CellRef]]
    dynamic_unread: set[CellRef] = field(default_factory=set)
    """Cells with a dynamic-ref probe that errored or could not be parsed."""


class NameScopes(Protocol):
    """What :func:`anchor` needs to know about defined names."""

    def is_sheet_scoped_name(self, name: str) -> bool: ...

    def is_formula_name(self, name: str) -> bool: ...


# -- near flips ----------------------------------------------------------------


def _near(y: float, target: float) -> bool:
    return abs(y - target) <= NEAR_TOLERANCE * max(1.0, abs(y))


def _half(y: float) -> bool:
    return _near(y, math.floor(y) + 0.5)


def _whole(y: float) -> bool:
    # Exactly on the boundary is not near (typed whole numbers sit there); only a
    # value a hair off it is.
    return y != round(y) and _near(y, float(round(y)))


def near_flip(fn: str, x: float, arg: float | None) -> bool:
    """True when a tiny change in ``x`` could flip the result of ``fn``.

    ``fn`` is a rounding function (``arg`` its digits / multiple / significance,
    None when absent) or a comparison operator / ``lookup`` (``arg`` the other side).

    Raises:
        ValueError: for an unknown ``fn``.
    """
    name = fn.upper()
    if name.lower() in _NEAR_COMPARE:
        if arg is None or not (math.isfinite(x) and math.isfinite(arg)):
            return False
        return abs(x - arg) <= NEAR_TOLERANCE * max(1.0, abs(x), abs(arg))
    if name not in _ROUNDING.values():
        raise ValueError(f"near_flip: unknown function {fn!r}")
    if not math.isfinite(x) or (arg is not None and not math.isfinite(arg)):
        return False
    try:
        if name == "INT":
            return _whole(x)
        if name in {"ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC"}:
            y = x * 10.0 ** math.trunc(arg or 0.0)
            return _half(y) if name == "ROUND" else _whole(y)
        if name == "MROUND":
            return arg is not None and arg != 0 and _half(x / arg)
        sig = 1.0 if arg is None else arg
        return sig != 0 and _whole(x / sig)
    except OverflowError:
        return False


# -- anchoring -----------------------------------------------------------------


def _end(text: str) -> tuple[str, str]:
    """Classify one reference end and make it absolute: (kind, text)."""
    if (m := _CELL_END.match(text)) is not None:
        return "cell", f"${m.group('col').upper()}${m.group('row')}"
    if (m := _COL_END.match(text)) is not None:
        return "col", f"${m.group('col').upper()}"
    if (m := _ROW_END.match(text)) is not None:
        return "row", f"${m.group('row')}"
    return "name", text


def _anchor_ref(ref: str, sheet: str, names: NameScopes | None) -> str:
    m = _REF.match(ref)
    if m is None:
        raise ValueError(f"cannot lift reference {ref!r}")
    s1, s2, b = m.group("s1"), m.group("s2"), m.group("b")
    kind, a = _end(m.group("a"))
    prefix = s1 if s1 is not None else quote_sheet_name(sheet)
    if b is None:
        if kind == "cell":
            return f"{prefix}!{a}"
        name = m.group("a")
        if name.lower().startswith("_xlpm."):
            raise ValueError(f"cannot lift LET/LAMBDA parameter {name!r}")
        if s1 is None and names is not None:
            if names.is_sheet_scoped_name(name):
                raise ValueError(f"cannot lift sheet-scoped name {name!r}")
            if names.is_formula_name(name):
                raise ValueError(f"cannot lift formula name {name!r}")
        return ref
    kind_b, b_abs = _end(b)
    if kind != kind_b or kind == "name" or (s2 is not None and s2 != prefix):
        raise ValueError(f"cannot lift reference {ref!r}")
    return f"{prefix}!{a}:{b_abs}"


def _const(value: object) -> str:
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, str):
        # The parser keeps error literals and text literals alike as str, so neither
        # can be written back faithfully.
        if value.upper() in ERROR_CODES:
            raise ValueError(f"cannot lift error-code constant {value!r}")
        return _text_literal(value)
    if isinstance(value, float):
        return repr(value)
    return str(value)


def _text_literal(value: str) -> str:
    escaped = value.replace('"', '""')
    return f'"{escaped}"'


def _emit(node: ASTNode, sheet: str, names: NameScopes | None) -> str:
    if isinstance(node, ConstNode):
        return _const(node.value)
    if isinstance(node, RefNode):
        return _anchor_ref(node.ref, sheet, names)
    if isinstance(node, UnaryNode):
        inner = _emit(node.operand, sheet, names)
        if isinstance(node.operand, BinaryNode):
            inner = f"({inner})"
        return f"{inner}%" if node.operator == "%" else f"{node.operator}{inner}"
    if isinstance(node, BinaryNode):
        left = _emit(node.left, sheet, names)
        right = _emit(node.right, sheet, names)
        if isinstance(node.left, BinaryNode):
            left = f"({left})"
        if isinstance(node.right, BinaryNode):
            right = f"({right})"
        return f"{left}{node.operator}{right}"
    if isinstance(node, FunctionNode):
        return _emit_function(node, sheet, names)
    raise ValueError(f"cannot lift node {node!r}")


def _emit_function(node: FunctionNode, sheet: str, names: NameScopes | None) -> str:
    name = function_name(node.name)
    if name in _SCOPING:
        raise ValueError(f"cannot lift {name}: it binds names")
    if name in _POSITIONAL and not node.args:
        raise ValueError(f"cannot lift {name}() off its own cell")
    if name == "CELL" and len(node.args) < 2:
        raise ValueError("cannot lift CELL without a reference")
    args = [_emit(a, sheet, names) for a in node.args]
    if name == "INDIRECT":
        if not args:
            raise ValueError("cannot lift INDIRECT without arguments")
        if len(node.args) > 1 and sign(literal(node.args[1])) != 1:
            raise ValueError("cannot lift INDIRECT with R1C1-style text")
        # Text without a sheet resolves against the formula's own sheet, not the probe's.
        text = args[0]
        prefix = _text_literal(f"{quote_sheet_name(sheet)}!")
        args[0] = f'IF(ISERROR(FIND("!",{text})),{prefix}&({text}),{text})'
    return f"{node.name}({','.join(args)})"


def anchor(node: ASTNode, sheet: str, *, names: NameScopes | None = None) -> str:
    """Formula text for ``node`` that computes the same on the probe sheet.

    Unqualified references are qualified with ``sheet`` (quoted like
    :meth:`CellRef.a1`); every cell, column and row reference becomes absolute;
    defined names stay names. ``INDIRECT`` text without a sheet is prefixed with
    ``sheet`` at run time.

    Raises:
        ValueError: for a sheet-scoped or formula name (when ``names`` is given), an
            array constant or unrecognised reference, ``LET``/``LAMBDA`` and friends,
            a bare ``ROW()``/``COLUMN()``/``SHEET()``, ``CELL`` without a reference,
            or R1C1-style ``INDIRECT``.
    """
    return _emit(node, sheet, names)


# -- planning ------------------------------------------------------------------


Builder = Callable[[], tuple[tuple[str, ...], dict[str, str]]]


@dataclass(slots=True)
class _Planner:
    graph: WorkbookGraph
    cap: int
    probes: list[Probe] = field(default_factory=list)
    unplaced: dict[CellRef, list[ProbeKind]] = field(default_factory=dict)
    cell: CellRef = CellRef("", 0, 0)
    family: str = ""
    counters: dict[ProbeKind, int] = field(default_factory=dict)

    def lift(self, node: ASTNode) -> str:
        return anchor(node, self.cell.sheet, names=self.graph)

    def add(self, kind: ProbeKind, scoped: bool, build: Builder) -> None:
        ordinal = self.counters.get(kind, 0)
        self.counters[kind] = ordinal + 1
        if scoped or len(self.probes) >= self.cap:
            self.unplaced.setdefault(self.cell, []).append(kind)
            return
        try:
            formulas, meta = build()
        except ValueError as exc:
            logger.debug("probe {} {} unplaced: {}", self.cell.a1(), kind, exc)
            self.unplaced.setdefault(self.cell, []).append(kind)
            return
        self.probes.append(Probe(kind, self.cell, self.family, ordinal, formulas, meta))

    def visit(self, node: ASTNode, in_cond: bool, blocked: bool) -> None:
        """Probe ``node`` and its children; ``blocked`` probes go unplaced (under
        ``LET``/``LAMBDA``, in an array formula or an always-array function)."""
        if isinstance(node, BinaryNode):
            if in_cond and node.operator in _COMPARISONS:
                self.add(ProbeKind.comparison, blocked, lambda: self._comparison(node))
            self.visit(node.left, in_cond, blocked)
            self.visit(node.right, in_cond, blocked)
            return
        if isinstance(node, UnaryNode):
            self.visit(node.operand, in_cond, blocked)
            return
        if not isinstance(node, FunctionNode):
            return
        name = function_name(node.name)
        blocked = blocked or name in _SCOPING
        self._probe_function(node, name, blocked)
        inner = blocked or name in ALWAYS_ARRAY
        conds = _condition_args(name, len(node.args))
        for i, arg in enumerate(node.args):
            self.visit(arg, in_cond or i in conds, inner)

    def _probe_function(self, node: FunctionNode, name: str, scoped: bool) -> None:
        if name in _BRANCHES and node.args:
            self.add(ProbeKind.branch, scoped, lambda: self._branch(node, name))
        elif name in _LOOKUPS and len(node.args) >= 2:
            self.add(ProbeKind.lookup, scoped, lambda: self._lookup(node, name))
            if _lookup_mode(node, name) != "0":
                self.add(ProbeKind.comparison, scoped, lambda: self._lookup_key(node, name))
        elif name in _ROUNDING and node.args:
            self.add(ProbeKind.rounding, scoped, lambda: self._rounding(node, name))
        elif name in _DYNAMIC and node.args:
            self.add(ProbeKind.dynamic_ref, scoped, lambda: self._dynamic(node, name))

    # -- builders --

    def _comparison(self, node: BinaryNode) -> tuple[tuple[str, ...], dict[str, str]]:
        return (self.lift(node.left), self.lift(node.right)), {"operator": node.operator}

    def _branch(self, node: FunctionNode, name: str) -> tuple[tuple[str, ...], dict[str, str]]:
        # Only the deciding arguments are lifted: what a branch returns is irrelevant.
        args = node.args
        meta = {"function": name}
        if name == "IFS":
            pairs = [f'{self.lift(c)},"{i}"' for i, c in enumerate(args[0:-1:2], 1)]
            return (f'{node.name}({",".join(pairs)},TRUE,"none")',), meta
        if name == "SWITCH":
            cases = [f'{self.lift(v)},"{i}"' for i, v in enumerate(args[1:-1:2], 1)]
            # No default: an unmatched value is #N/A in the cell, so it reads "error".
            default = ',"default"' if len(args) % 2 == 0 else ""
            return (f"{node.name}({','.join([self.lift(args[0]), *cases])}{default})",), meta
        first = self.lift(args[0])
        if name == "IF":
            return (f'IF({first},"TRUE","FALSE")',), meta
        if name == "CHOOSE":
            return (f"TRUNC({first})",), {**meta, "choices": str(len(args) - 1)}
        test = "ISERROR" if name == "IFERROR" else "ISNA"
        return (f"{test}({first})",), meta

    def _keys(self, node: FunctionNode, name: str) -> tuple[str, str]:
        """Key vector and its length, as formula text."""
        table = self.lift(node.args[1])
        if name == "VLOOKUP":
            return f"INDEX({table},0,1)", f"ROWS({table})"
        if name == "HLOOKUP":
            return f"INDEX({table},1,0)", f"COLUMNS({table})"
        if name == "LOOKUP" and len(node.args) == 2:
            # Array form: the first row when wider than tall, else the first column.
            wide = f"COLUMNS({table})>ROWS({table})"
            keys = f"IF({wide},INDEX({table},1,0),INDEX({table},0,1))"
            return keys, f"MAX(ROWS({table}),COLUMNS({table}))"
        return table, f"ROWS({table})*COLUMNS({table})"

    def _match(self, node: FunctionNode, name: str) -> str:
        keys, _size = self._keys(node, name)
        mode = _lookup_mode(node, name)
        if mode is None:
            mode_arg = node.args[3] if name in {"VLOOKUP", "HLOOKUP"} else node.args[2]
            lifted = self.lift(mode_arg)
            mode = f"IF({lifted},1,0)" if name in {"VLOOKUP", "HLOOKUP"} else lifted
        return f"MATCH({self.lift(node.args[0])},{keys},{mode})"

    def _lookup(self, node: FunctionNode, name: str) -> tuple[tuple[str, ...], dict[str, str]]:
        _keys, size = self._keys(node, name)
        mode = _lookup_mode(node, name)
        label = "exact" if mode == "0" else "dynamic" if mode is None else "approximate"
        return (self._match(node, name), size), {"function": name, "mode": label}

    def _lookup_key(self, node: FunctionNode, name: str) -> tuple[tuple[str, ...], dict[str, str]]:
        keys, _size = self._keys(node, name)
        key = f"INDEX({keys},{self._match(node, name)})"
        if _lookup_mode(node, name) is None:
            if name in {"VLOOKUP", "HLOOKUP"}:
                key = f'IF({self.lift(node.args[3])},{key},"exact")'
            else:
                key = f'IF(({self.lift(node.args[2])})=0,"exact",{key})'
        return (self.lift(node.args[0]), key), {"operator": "lookup"}

    def _rounding(self, node: FunctionNode, name: str) -> tuple[tuple[str, ...], dict[str, str]]:
        meta = {"function": _ROUNDING[name]}
        if name == "INT":
            return (self.lift(node.args[0]),), meta
        return tuple(self.lift(a) for a in node.args[:2]), meta

    def _dynamic(self, node: FunctionNode, name: str) -> tuple[tuple[str, ...], dict[str, str]]:
        if name == "INDIRECT":
            self.lift(node)  # raises for R1C1-style text
            return (self.lift(node.args[0]),), {"function": name, "sheet": self.cell.sheet}
        ref = self.lift(node)
        corner = f"OFFSET({ref},0,0,1,1)"
        formulas = (f"ROW({corner})", f"COLUMN({corner})", f"ROWS({ref})", f"COLUMNS({ref})")
        return formulas, {"function": name, "sheet": self._static_sheet(node.args[0])}

    def _static_sheet(self, node: ASTNode) -> str:
        """The sheet a reference expression lands on, when knowable without running it."""
        if isinstance(node, FunctionNode) and node.args:
            if function_name(node.name) in {"OFFSET", "INDEX"}:
                return self._static_sheet(node.args[0])
        if not isinstance(node, RefNode):
            raise ValueError("OFFSET base is not a static reference")
        m = _REF.match(node.ref)
        if m is None:
            raise ValueError(f"cannot lift reference {node.ref!r}")
        if _end(m.group("a"))[0] == "name" and m.group("b") is None:
            named = self.graph.name_cells(m.group("a"))
            if not named:
                raise ValueError(f"name {node.ref!r} has no cells")
            return named[0].sheet
        s1 = m.group("s1")
        return _unquote(s1) if s1 is not None else self.cell.sheet


def _unquote(sheet: str) -> str:
    if len(sheet) >= 2 and sheet.startswith("'") and sheet.endswith("'"):
        return sheet[1:-1].replace("''", "'")
    return sheet


def _condition_args(name: str, n: int) -> set[int]:
    """Argument positions that decide a branch."""
    if name in {"IF", "CHOOSE"}:
        return {0}
    if name == "IFS":
        return set(range(0, n - 1, 2))
    if name == "SWITCH":
        cases = (n - 1) // 2
        return {0, *range(1, 2 * cases, 2)}
    return set()


def _lookup_mode(node: FunctionNode, name: str) -> str | None:
    """MATCH mode text: ``"1"``/``"-1"`` approximate, ``"0"`` exact, None when not literal."""
    if name == "LOOKUP":
        return "1"
    at = 3 if name in {"VLOOKUP", "HLOOKUP"} else 2
    if len(node.args) <= at:
        return "1"
    s = sign(literal(node.args[at]))
    if s is None:
        return None
    if name in {"VLOOKUP", "HLOOKUP"}:
        return "0" if s == 0 else "1"
    return str(s)


def plan_probes(graph: WorkbookGraph, cells: Iterable[CellRef], *, cap: int) -> ProbePlan:
    """Plan probes for the branch, lookup, rounding, comparison and dynamic-ref
    nodes of each cell's formula; cells are taken in sorted order.

    A node that :func:`anchor` cannot lift, any node beneath ``LET``/``LAMBDA``,
    every node of an array formula or inside an always-array function's arguments
    (``SUMPRODUCT``...), and every probe beyond ``cap`` is recorded in ``unplaced`` (one entry per probe).
    Unparseable formulas produce no probes (the pattern scan refuses them).
    """
    planner = _Planner(graph, cap)
    for cell in sorted(set(cells)):
        text = graph.formula(cell)
        if text is None:
            continue
        root = parse_formula(text)
        if root is None:
            logger.debug("no probes for {}: unparseable formula", cell.a1())
            continue
        planner.cell = cell
        planner.family = graph.family(cell)
        planner.counters = {}
        planner.visit(root, False, graph.is_array_formula(cell))
    return ProbePlan(planner.probes, planner.unplaced)


# -- placing and reading -------------------------------------------------------


def _slot(index: int, column: int) -> CellRef:
    return CellRef(PROBE_SHEET, index + 1, column + 1)


def place_probes(engine: Engine, plan: ProbePlan) -> None:
    """Write the plan to the hidden ``_marinade_probes`` sheet: one probe per row,
    its formulas across the columns, in plan order. The caller evaluates."""
    if not plan.probes:
        return
    engine.add_hidden_sheet(PROBE_SHEET)
    for i, probe in enumerate(plan.probes):
        for j, formula in enumerate(probe.formulas):
            engine.set_formula(_slot(i, j), formula)


def _number(value: EngineValue) -> float | None:
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, float):
        return value
    return None


def _is_error(value: EngineValue) -> bool:
    return isinstance(value, str) and value.startswith("#")


def _branch_outcome(probe: Probe, values: list[EngineValue]) -> str:
    value = values[0]
    function = probe.meta["function"]
    if function in {"IFERROR", "IFNA"}:
        return "ok" if value is False else "error"
    if function == "CHOOSE":
        n = _number(value)
        if n is None or _is_error(value):
            return "error"
        index = math.trunc(n)
        return str(index) if 1 <= index <= int(probe.meta["choices"]) else "error"
    if not isinstance(value, str) or _is_error(value):
        return "error"
    return value


def _lookup_edge(values: list[EngineValue]) -> str:
    pos, size = values[0], values[1]
    if not isinstance(pos, float) or not isinstance(size, float):
        return "not_found"
    if pos == 1:
        return "first"
    return "last" if pos == size else "interior"


def _rounding_near(probe: Probe, values: list[EngineValue]) -> bool:
    x = _number(values[0])
    arg = _number(values[1]) if len(values) > 1 else None
    if x is None or (len(values) > 1 and arg is None):
        return False
    return near_flip(probe.meta["function"], x, arg)


def _comparison_near(probe: Probe, values: list[EngineValue]) -> bool:
    a, b = values
    if not isinstance(a, float) or not isinstance(b, float):
        return False
    return near_flip(probe.meta["operator"], a, b)


def _text_cells(text: EngineValue, sheet: str) -> list[CellRef] | None:
    """Cells of an ``INDIRECT`` text address (A1 style); None when not a plain range."""
    if not isinstance(text, str) or _is_error(text):
        return None
    m = _TEXT_REF.match(text.strip())
    if m is None:
        return None
    target = _unquote(m.group("sheet")) if m.group("sheet") else sheet
    r1, c1 = int(m.group("r1")), _col_to_number(m.group("c1"))
    r2 = int(m.group("r2")) if m.group("r2") else r1
    c2 = _col_to_number(m.group("c2")) if m.group("c2") else c1
    return _rect(target, min(r1, r2), min(c1, c2), abs(r2 - r1) + 1, abs(c2 - c1) + 1)


def _rect(sheet: str, row: int, col: int, rows: int, cols: int) -> list[CellRef]:
    return [CellRef(sheet, r, c) for r in range(row, row + rows) for c in range(col, col + cols)]


def _dynamic_cells(probe: Probe, values: list[EngineValue]) -> list[CellRef] | None:
    """The cells a dynamic reference landed on; None when the reading is unusable."""
    if probe.meta["function"] == "INDIRECT":
        return _text_cells(values[0], probe.meta["sheet"])
    nums = [v for v in values if isinstance(v, float)]
    if len(nums) != 4 or nums[2] < 1 or nums[3] < 1:
        return None
    row, col, rows, cols = (int(v) for v in nums)
    return _rect(probe.meta["sheet"], row, col, rows, cols)


def read_probes(engine: Engine, plan: ProbePlan) -> ProbeReadings:
    """Read the placed probes after ``engine.evaluate()``."""
    readings = ProbeReadings({}, {}, set(), set(), {})
    dynamic: dict[CellRef, set[CellRef]] = {}
    for i, probe in enumerate(plan.probes):
        values = [engine.value(_slot(i, j)) for j in range(len(probe.formulas))]
        key = (probe.cell, probe.ordinal)
        if probe.kind is ProbeKind.branch:
            readings.branch[key] = _branch_outcome(probe, values)
        elif probe.kind is ProbeKind.lookup:
            readings.lookup_edge[key] = _lookup_edge(values)
        elif probe.kind is ProbeKind.rounding:
            if _rounding_near(probe, values):
                readings.near_rounding.add(probe.cell)
        elif probe.kind is ProbeKind.comparison:
            if _comparison_near(probe, values):
                readings.near_comparison.add(probe.cell)
        else:
            landed = _dynamic_cells(probe, values)
            dynamic.setdefault(probe.cell, set()).update(landed or ())
            if landed is None:
                readings.dynamic_unread.add(probe.cell)
    readings.dynamic_refs = {cell: sorted(cells) for cell, cells in dynamic.items()}
    return readings
