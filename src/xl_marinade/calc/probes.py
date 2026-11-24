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

A probe computes whether or not the cell's formula ever evaluates the part it was
lifted from. Each probe therefore records its *guards*: the enclosing branches and
which of their outcomes evaluate it. A reading whose guard was not met is dropped (it
says nothing about the value); one whose guarding branch has no reading is kept but
marked unsure, so it never counts as evidence from the trust run.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from datetime import date
from enum import StrEnum
from typing import Protocol

from loguru import logger

from xl_marinade.calc.engine import Engine, EngineValue
from xl_marinade.calc.formula_ast import (
    ALWAYS_ARRAY,
    END,
    ERROR_CODES,
    EXTERNAL,
    SHEET,
    function_name,
    literal,
    parse_formula,
    quote_sheet,
    sign,
)
from xl_marinade.calc.graph import NameDef, WorkbookGraph
from xl_marinade.calc.models import CellRef, _col_to_number
from xl_marinade.calc.prepare import rewrite_clock
from xl_marinade.core.parser import (
    ASTNode,
    BinaryNode,
    ConstNode,
    FunctionNode,
    RefNode,
    UnaryNode,
)

PROBE_SHEET = "_marinade_probes"
NEAR_TOLERANCE = 1e-12
_MAX_ROW = 1_048_576
_MAX_COL = 16_384

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
_ERROR = "error"
_NONE = frozenset[str]()
_LOOKUPS = frozenset({"VLOOKUP", "HLOOKUP", "MATCH", "LOOKUP", "XLOOKUP", "XMATCH"})
# XLOOKUP / XMATCH: positions of match_mode and search_mode.
_X_MODES: dict[str, tuple[int, int]] = {"XLOOKUP": (4, 5), "XMATCH": (2, 3)}
# XLOOKUP's if_not_found arm, guarded like IFERROR's error arm.
_NOT_FOUND = "not_found"
_DYNAMIC = frozenset({"OFFSET", "INDIRECT"})

_REF = re.compile(rf"^(?:(?P<s1>{SHEET})!)?(?P<a>{END})(?::(?:(?P<s2>{SHEET})!)?(?P<b>{END}))?$")
_CELL_END = re.compile(r"^\$?(?P<col>[A-Za-z]{1,3})\$?(?P<row>[0-9]+)$")
_COL_END = re.compile(r"^\$?(?P<col>[A-Za-z]{1,3})$")
_ROW_END = re.compile(r"^\$?(?P<row>[0-9]+)$")
_TEXT_REF = re.compile(
    rf"^(?:(?P<sheet>{SHEET})!)?\$?(?P<c1>[A-Za-z]{{1,3}})\$?(?P<r1>[0-9]+)"
    r"(?::\$?(?P<c2>[A-Za-z]{1,3})\$?(?P<r2>[0-9]+))?$"
)


Guard = tuple[int, frozenset[str], frozenset[str]]
"""An enclosing branch: its branch ordinal in the same cell, the outcomes that evaluate
the guarded part, and the outcomes after which it may or may not have been evaluated."""


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
    lookup's value against its matched key). ``guards`` lists the enclosing
    branches, outermost first.
    """

    kind: ProbeKind
    cell: CellRef
    family: str
    ordinal: int
    formulas: tuple[str, ...]
    meta: dict[str, str]
    guards: tuple[Guard, ...] = ()


@dataclass(slots=True)
class ProbePlan:
    """The probes to place, in order, and per cell the kinds that could not be placed.

    ``sheets`` maps each lower-cased sheet name to the workbook's spelling, so the
    sheet in an ``INDIRECT`` text is read as Excel reads it, in any case.
    """

    probes: list[Probe]
    unplaced: dict[CellRef, list[ProbeKind]]
    sheets: dict[str, str] = field(default_factory=dict)


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
    unsure: set[tuple[ProbeKind, CellRef, int]] = field(default_factory=set)
    """Readings kept although an enclosing branch had no reading: not evidence."""
    engine_errors: dict[CellRef, str] = field(default_factory=dict)
    """Cells whose probe computed an error code Excel does not have."""
    type_mismatch: set[CellRef] = field(default_factory=set)
    """Cells with an approximate lookup whose value and matched key differ in type."""


class NameScopes(Protocol):
    """What :func:`anchor` needs to know about defined names."""

    def lookup_name(self, ref: str, sheet: str | None) -> NameDef | None: ...


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
    None when absent) or a comparison operator / ``lookup`` (``arg`` the other side;
    sides exactly equal are not near).

    Raises:
        ValueError: for an unknown ``fn``.
    """
    name = fn.upper()
    if name.lower() in _NEAR_COMPARE:
        if arg is None or not (math.isfinite(x) and math.isfinite(arg)):
            return False
        # Exactly equal sides are not near: the comparison is decided, not on a knife
        # edge (``IF(t=0, ...)`` on a whole-number period is the common case).
        distance = abs(x - arg)
        return 0 < distance <= NEAR_TOLERANCE * max(1.0, abs(x), abs(arg))
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
    if EXTERNAL.match(ref):
        # The engine copy holds the cached value only in the cell's own formula; on
        # the probe sheet the reference would not resolve.
        raise ValueError(f"cannot lift external reference {ref!r}")
    m = _REF.match(ref)
    if m is None:
        raise ValueError(f"cannot lift reference {ref!r}")
    s1, s2, b = m.group("s1"), m.group("s2"), m.group("b")
    kind, a = _end(m.group("a"))
    prefix = s1 if s1 is not None else quote_sheet(sheet)
    if b is None:
        if kind == "cell":
            return f"{prefix}!{a}"
        name = m.group("a")
        if name.lower().startswith("_xlpm."):
            raise ValueError(f"cannot lift LET/LAMBDA parameter {name!r}")
        return _lift_name(ref, sheet, names)
    kind_b, b_abs = _end(b)
    if kind != kind_b or kind == "name" or (s2 is not None and s2 != prefix):
        raise ValueError(f"cannot lift reference {ref!r}")
    return f"{prefix}!{a}:{b_abs}"


def _lift_name(ref: str, sheet: str, names: NameScopes | None) -> str:
    """A defined name as the probe sheet must read it.

    Only a workbook-level name written bare can be lifted as it is. The probe sheet
    would read the workbook's definition of a name the formula's own sheet defines, the
    engine reads no ``Sheet!Name``, and a name's reference may be relative to the cell
    that uses it (the IR keeps no ``$``), so none of those is lifted.
    """
    found = names.lookup_name(ref, sheet) if names is not None else None
    if found is None:
        return ref
    if found.kind != "cells":
        raise ValueError(f"cannot lift formula name {ref!r}")
    if found.relative:
        raise ValueError(f"cannot lift name with a relative reference {ref!r}")
    if found.scope is not None or "!" in ref:
        raise ValueError(f"cannot lift sheet-scoped name {ref!r}")
    return ref


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
        prefix = _text_literal(f"{quote_sheet(sheet)}!")
        args[0] = f'IF(ISERROR(FIND("!",{text})),{prefix}&({text}),{text})'
    return f"{node.name}({','.join(args)})"


def anchor(
    node: ASTNode,
    sheet: str,
    *,
    names: NameScopes | None = None,
    as_of: date | None = None,
) -> str:
    """Formula text for ``node`` that computes the same on the probe sheet.

    Unqualified references are qualified with ``sheet`` (quoted by
    :func:`quote_sheet`, so ``'2024'!``); every cell, column and row reference
    becomes absolute; workbook-level names stay names. ``INDIRECT`` text without a
    sheet is prefixed with ``sheet`` at run time. With ``as_of``, ``TODAY()`` and
    ``NOW()`` become that date, as they are in the engine copy. Operands of binary
    operators are always parenthesised when they are operations themselves, so the
    text means what the parsed tree means.

    Raises:
        ValueError: (when ``names`` is given) for a formula name, a name with a
            relative reference, a name ``sheet`` defines for itself, or a name written
            ``Sheet!Name``; and for an array constant or unrecognised reference,
            ``LET``/``LAMBDA`` and friends, a bare ``ROW()``/``COLUMN()``/``SHEET()``,
            ``CELL`` without a reference, or R1C1-style ``INDIRECT``.
    """
    text = _emit(node, sheet, names)
    return rewrite_clock(text, as_of) if as_of is not None else text


# -- planning ------------------------------------------------------------------


Builder = Callable[[], tuple[tuple[str, ...], dict[str, str]]]


@dataclass(slots=True)
class _Planner:
    graph: WorkbookGraph
    cap: int
    as_of: date | None = None
    probes: list[Probe] = field(default_factory=list)
    unplaced: dict[CellRef, list[ProbeKind]] = field(default_factory=dict)
    cell: CellRef = CellRef("", 0, 0)
    family: str = ""
    counters: dict[ProbeKind, int] = field(default_factory=dict)
    guards: tuple[Guard, ...] = ()

    def lift(self, node: ASTNode) -> str:
        return anchor(node, self.cell.sheet, names=self.graph, as_of=self.as_of)

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
        self.probes.append(
            Probe(kind, self.cell, self.family, ordinal, formulas, meta, self.guards)
        )

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
        branch = self.counters.get(ProbeKind.branch, 0)
        self._probe_function(node, name, blocked)
        inner = blocked or name in ALWAYS_ARRAY
        conds = _condition_args(name, len(node.args))
        arms = _arm_outcomes(name, len(node.args)) if _is_branch(node, name) else {}
        outer = self.guards
        for i, arg in enumerate(node.args):
            arm = arms.get(i)
            if arm is not None:
                self.guards = (*outer, (branch, *arm))
            self.visit(arg, in_cond or i in conds, inner)
            self.guards = outer

    def _probe_function(self, node: FunctionNode, name: str, scoped: bool) -> None:
        if name == "XLOOKUP" and _is_branch(node, name):
            # Found or not decides whether the if_not_found arm is used.
            self.add(ProbeKind.branch, scoped, lambda: self._not_found(node))
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
        # The value itself when it errors, so an engine-only error code (which the
        # cell would swallow) is still seen; FALSE when it does not.
        return (f"IF(ISERROR({first}),{first},FALSE)",), meta

    def _xmatch(self, node: FunctionNode, name: str) -> str:
        """XMATCH text for an XLOOKUP / XMATCH: the position its own modes find."""
        match_at, search_at = _X_MODES[name]
        args = [self.lift(node.args[0]), self.lift(node.args[1])]
        if len(node.args) > search_at:
            mode = self.lift(node.args[match_at]) if len(node.args) > match_at else "0"
            args += [mode, self.lift(node.args[search_at])]
        elif len(node.args) > match_at:
            args.append(self.lift(node.args[match_at]))
        return f"XMATCH({','.join(args)})"

    def _not_found(self, node: FunctionNode) -> tuple[tuple[str, ...], dict[str, str]]:
        # #N/A is "nothing found"; any other error is the cell's error, read as such.
        pos = self._xmatch(node, "XLOOKUP")
        text = f'IF(ISNA({pos}),"{_NOT_FOUND}",IF(ISERROR({pos}),{pos},"found"))'
        return (text,), {"function": "XLOOKUP"}

    def _keys(self, node: FunctionNode, name: str) -> tuple[str, str]:
        """Key vector and its length, as formula text."""
        table = self.lift(node.args[1])
        if name in _X_MODES:
            # The lookup array is one row or one column.
            return table, f"ROWS({table})*COLUMNS({table})"
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
        if name in _X_MODES:
            return self._xmatch(node, name)
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
            if name in _X_MODES:
                # Only next smaller / next larger can match a key that differs.
                mode = self.lift(node.args[_X_MODES[name][0]])
                key = f"IF(ABS({mode})=1,{key},NA())"
            elif name in {"VLOOKUP", "HLOOKUP"}:
                key = f"IF({self.lift(node.args[3])},{key},NA())"
            else:
                key = f"IF(({self.lift(node.args[2])})=0,NA(),{key})"
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
            found = self.graph.lookup_name(node.ref, self.cell.sheet)
            sheets = {r[0] for r in found.rects} if found is not None else set()
            if found is None or found.kind != "cells" or found.whole or len(sheets) != 1:
                # A whole column or row would land on up to a million cells per reading.
                raise ValueError(f"name {node.ref!r} is not a bounded range on one sheet")
            return sheets.pop()
        s1 = m.group("s1")
        if s1 is None:
            return self.cell.sheet
        written = _unquote(s1)
        return self.graph.sheet(written) or written


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


def _is_branch(node: FunctionNode, name: str) -> bool:
    """A function with a branch probe: the IF family, or an XLOOKUP with if_not_found."""
    if name == "XLOOKUP":
        return len(node.args) >= 4
    return name in _BRANCHES and bool(node.args)


def _arm_outcomes(name: str, n: int) -> dict[int, tuple[frozenset[str], frozenset[str]]]:
    """Per guarded argument of a branch: the outcomes that evaluate it, and those after
    which it may have been evaluated (an error from a condition before or at it).

    Arguments not listed are always evaluated (the first condition, the switch value).
    """
    if name == "XLOOKUP":
        return {3: (frozenset({_NOT_FOUND}), _NONE)}
    if name == "IF":
        return {1: (frozenset({"TRUE"}), _NONE), 2: (frozenset({"FALSE"}), _NONE)}
    if name == "CHOOSE":
        return {i: (frozenset({str(i)}), _NONE) for i in range(1, n)}
    if name in {"IFERROR", "IFNA"}:
        return {1: (frozenset({_ERROR}), _NONE)}
    arms: dict[int, tuple[frozenset[str], frozenset[str]]] = {}
    if name == "IFS":
        pairs = n // 2
        for k in range(1, pairs + 1):
            cond, value = 2 * (k - 1), 2 * k - 1
            if k > 1:
                later = frozenset({*(str(j) for j in range(k, pairs + 1)), "none"})
                arms[cond] = (later, frozenset({_ERROR}))
            arms[value] = (frozenset({str(k)}), _NONE)
        return arms
    if name == "SWITCH":
        cases = (n - 1) // 2
        for k in range(1, cases + 1):
            case, result = 2 * k - 1, 2 * k
            if k > 1:
                later = frozenset({*(str(j) for j in range(k, cases + 1)), "default"})
                arms[case] = (later, frozenset({_ERROR}))
            arms[result] = (frozenset({str(k)}), _NONE)
        if n % 2 == 0:
            arms[n - 1] = (frozenset({"default"}), _NONE)
    return arms


def _lookup_mode(node: FunctionNode, name: str) -> str | None:
    """MATCH mode text: ``"1"``/``"-1"`` approximate, ``"0"`` exact, None when not literal.

    For XLOOKUP / XMATCH: ``"0"`` for exact and wildcard (2) matches, ``"1"``/``"-1"``
    for next larger / next smaller.
    """
    if name == "LOOKUP":
        return "1"
    if name in _X_MODES:
        at = _X_MODES[name][0]
        if len(node.args) <= at:
            return "0"
        value = literal(node.args[at])
        if value in (2, 2.0):
            return "0"
        s = sign(value)
        return None if s is None else str(s)
    at = 3 if name in {"VLOOKUP", "HLOOKUP"} else 2
    if len(node.args) <= at:
        return "1"
    s = sign(literal(node.args[at]))
    if s is None:
        return None
    if name in {"VLOOKUP", "HLOOKUP"}:
        return "0" if s == 0 else "1"
    return str(s)


def plan_probes(
    graph: WorkbookGraph, cells: Iterable[CellRef], *, cap: int, as_of: date | None = None
) -> ProbePlan:
    """Plan probes for the branch, lookup, rounding, comparison and dynamic-ref
    nodes of each cell's formula; cells are taken in sorted order. ``as_of`` is the
    run's ``--as-of`` date, which lifted ``TODAY()``/``NOW()`` read.

    A node that :func:`anchor` cannot lift, any node beneath ``LET``/``LAMBDA``,
    every node of an array formula or inside an always-array function's arguments
    (``SUMPRODUCT``...), and every probe beyond ``cap`` is recorded in ``unplaced`` (one entry per probe).
    Unparseable formulas produce no probes (the pattern scan refuses them).
    """
    planner = _Planner(graph, cap, as_of)
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
    return ProbePlan(
        planner.probes, planner.unplaced, {name.lower(): name for name in graph.sheets()}
    )


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
    if function == "IFERROR":
        return "ok" if value is False else _ERROR
    if function == "IFNA":
        return _ERROR if value == "#N/A" else "ok"
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


def _value_type(value: EngineValue) -> str | None:
    """Excel's comparison class of a lookup value; None for empty or an error."""
    if value is None or _is_error(value):
        return None
    if isinstance(value, bool):
        return "bool"
    return "number" if isinstance(value, float) else "text"


def _types_differ(probe: Probe, values: list[EngineValue]) -> bool:
    """True when an approximate lookup matched a key of another type than its value
    (Excel finds nothing there; the engine may return a row)."""
    if probe.meta.get("operator") != "lookup":
        return False
    a, b = _value_type(values[0]), _value_type(values[1])
    return a is not None and b is not None and a != b


def _text_cells(text: EngineValue, sheet: str, sheets: dict[str, str]) -> list[CellRef] | None:
    """Cells of an ``INDIRECT`` text address (A1 style); None when not a plain range.

    An address off the grid (row 0, past row 1,048,576 or column XFD) is None too:
    Excel's ``INDIRECT`` gives ``#REF!`` there, while IronCalc 0.9.1 reads a
    sheet-qualified row 0 as 0 (or a range from it as starting on row 1).
    """
    if not isinstance(text, str) or _is_error(text):
        return None
    m = _TEXT_REF.match(text.strip())
    if m is None:
        return None
    written = _unquote(m.group("sheet")) if m.group("sheet") else sheet
    target = sheets.get(written.lower(), written)
    r1, c1 = int(m.group("r1")), _col_to_number(m.group("c1"))
    r2 = int(m.group("r2")) if m.group("r2") else r1
    c2 = _col_to_number(m.group("c2")) if m.group("c2") else c1
    if not (1 <= min(r1, r2) and max(r1, r2) <= _MAX_ROW and max(c1, c2) <= _MAX_COL):
        return None
    return _rect(target, min(r1, r2), min(c1, c2), abs(r2 - r1) + 1, abs(c2 - c1) + 1)


def _rect(sheet: str, row: int, col: int, rows: int, cols: int) -> list[CellRef]:
    return [CellRef(sheet, r, c) for r in range(row, row + rows) for c in range(col, col + cols)]


def _dynamic_cells(
    probe: Probe, values: list[EngineValue], sheets: dict[str, str]
) -> list[CellRef] | None:
    """The cells a dynamic reference landed on; None when the reading is unusable."""
    if probe.meta["function"] == "INDIRECT":
        return _text_cells(values[0], probe.meta["sheet"], sheets)
    nums = [v for v in values if isinstance(v, float)]
    if len(nums) != 4 or nums[2] < 1 or nums[3] < 1:
        return None
    row, col, rows, cols = (int(v) for v in nums)
    return _rect(probe.meta["sheet"], row, col, rows, cols)


def _live(probe: Probe, outcomes: dict[tuple[CellRef, int], str]) -> bool | None:
    """Whether the cell evaluates the probe's part: True, False, or None (unknown)."""
    live: bool | None = True
    for ordinal, taken, maybe in probe.guards:
        outcome = outcomes.get((probe.cell, ordinal))
        if outcome is None or outcome in maybe:
            live = None
        elif outcome not in taken:
            return False
    return live


def engine_error_code(engine: Engine, cell: CellRef, value: EngineValue) -> str | None:
    """The error code when ``cell`` holds an error Excel does not have."""
    if isinstance(value, str) and value.startswith("#") and value not in ERROR_CODES:
        return value if engine.is_error(cell) else None
    return None


def read_probes(engine: Engine, plan: ProbePlan) -> ProbeReadings:
    """Read the placed probes after ``engine.evaluate()``.

    Probes inside a branch arm the cell did not evaluate are skipped; those whose
    enclosing branch has no reading are kept and listed in ``unsure``.
    """
    readings = ProbeReadings({}, {}, set(), set(), {})
    dynamic: dict[CellRef, set[CellRef]] = {}
    values_of = [
        [engine.value(_slot(i, j)) for j in range(len(probe.formulas))]
        for i, probe in enumerate(plan.probes)
    ]
    outcomes = {
        (probe.cell, probe.ordinal): _branch_outcome(probe, values_of[i])
        for i, probe in enumerate(plan.probes)
        if probe.kind is ProbeKind.branch
    }
    for i, probe in enumerate(plan.probes):
        live = _live(probe, outcomes)
        if live is False:
            continue
        values = values_of[i]
        key = (probe.cell, probe.ordinal)
        if live is None:
            readings.unsure.add((probe.kind, probe.cell, probe.ordinal))
        for j, value in enumerate(values):
            code = engine_error_code(engine, _slot(i, j), value)
            if code is not None:
                readings.engine_errors.setdefault(probe.cell, code)
        if probe.kind is ProbeKind.branch:
            readings.branch[key] = outcomes[key]
        elif probe.kind is ProbeKind.lookup:
            readings.lookup_edge[key] = _lookup_edge(values)
        elif probe.kind is ProbeKind.rounding:
            if _rounding_near(probe, values):
                readings.near_rounding.add(probe.cell)
        elif probe.kind is ProbeKind.comparison:
            if _comparison_near(probe, values):
                readings.near_comparison.add(probe.cell)
            if _types_differ(probe, values):
                readings.type_mismatch.add(probe.cell)
        else:
            landed = _dynamic_cells(probe, values, plan.sheets)
            dynamic.setdefault(probe.cell, set()).update(landed or ())
            if landed is None:
                readings.dynamic_unread.add(probe.cell)
    readings.dynamic_refs = {cell: sorted(cells) for cell, cells in dynamic.items()}
    return readings
