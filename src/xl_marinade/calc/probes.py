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
from decimal import Decimal
from enum import StrEnum
from typing import Protocol

from loguru import logger

from xl_marinade.calc.engine import Engine, EngineValue
from xl_marinade.calc.formula_ast import (
    ALWAYS_ARRAY,
    DYNAMIC_HINT,
    END,
    ERROR_CODES,
    EXTERNAL,
    SHEET,
    UNKNOWN,
    built_range,
    function_name,
    iter_nodes,
    literal,
    parse_formula,
    quote_sheet,
    sign,
)
from xl_marinade.calc.graph import NameDef, WorkbookGraph
from xl_marinade.calc.models import CellRef, _col_to_number, _number_to_col
from xl_marinade.calc.prepare import rewrite_clock
from xl_marinade.core.parser import (
    ASTNode,
    BinaryNode,
    ConstNode,
    FunctionNode,
    RefNode,
    UnaryNode,
)
from xl_marinade.errors import CalcInputError

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
# A dynamic-ref probe of a range built by a call (``B1:OFFSET(...)``).
_RANGE = "range"

_REF = re.compile(rf"^(?:(?P<s1>{SHEET})!)?(?P<a>{END})(?::(?:(?P<s2>{SHEET})!)?(?P<b>{END}))?$")
_CELL_END = re.compile(r"^\$?(?P<col>[A-Za-z]{1,3})\$?(?P<row>[0-9]+)$")
_COL_END = re.compile(r"^\$?(?P<col>[A-Za-z]{1,3})$")
_ROW_END = re.compile(r"^\$?(?P<row>[0-9]+)$")

# -- number to text (ruling D6) --
NUMBER_TEXT_TIE = "number-to-text at a decimal tie"
NUMBER_TEXT_DIGITS = "number-to-text where Excel's 15-digit text differs"
NUMBER_TEXT_CRITERION = "criterion built with & from a non-integer number"
NUMBER_TEXT_NOT_CHECKED = "number-to-text not checked"
NUMBER_TEXT_NEGATIVE = "number-to-text of a negative value that rounds to zero or one unit"
# Functions that round a number to a format's decimals: TEXT(x, format),
# FIXED(x, [decimals=2], ...), DOLLAR(x, [decimals=2]).
_FORMATTERS = frozenset({"TEXT", "FIXED", "DOLLAR"})
# Joins, and the first argument each joins as text (TEXTJOIN's delimiter and
# ignore_empty come first).
_JOINS: dict[str, int] = {"CONCAT": 0, "CONCATENATE": 0, "TEXTJOIN": 2}
# Text functions, and the arguments at which a number given to them becomes text (the
# engine's text, with up to 17 significant digits). The one list of these sites;
# tests/calc/test_number_text_sites.py pins it against the installed IronCalc.
TEXT_ARGUMENTS: dict[str, tuple[int, ...]] = {
    "LEN": (0,),
    "LENB": (0,),
    "LEFT": (0,),
    "LEFTB": (0,),
    "RIGHT": (0,),
    "RIGHTB": (0,),
    "MID": (0,),
    "MIDB": (0,),
    "UPPER": (0,),
    "LOWER": (0,),
    "PROPER": (0,),
    "TRIM": (0,),
    "CLEAN": (0,),
    "CODE": (0,),
    "UNICODE": (0,),
    "REPT": (0,),
    "SUBSTITUTE": (0, 1, 2),
    "REPLACE": (0, 3),
    "EXACT": (0, 1),
    "FIND": (0, 1),
    "FINDB": (0, 1),
    "SEARCH": (0, 1),
    "SEARCHB": (0, 1),
    "TEXTBEFORE": (0, 1),
    "TEXTAFTER": (0, 1),
    "VALUETOTEXT": (0,),
    "ARRAYTOTEXT": (0,),
    "REGEXTEST": (0,),
    "REGEXEXTRACT": (0,),
    "REGEXREPLACE": (0,),
    "TEXTSPLIT": (0,),
}
# Functions whose result is text: an operand that is one of them converts no number.
_TEXT_RESULT = frozenset(
    {"TEXT", "FIXED", "DOLLAR", "CONCAT", "CONCATENATE", "TEXTJOIN", "LEFT", "LEFTB"}
    | {"RIGHT", "RIGHTB", "MID", "MIDB", "UPPER", "LOWER", "PROPER", "TRIM", "CLEAN"}
    | {"REPT", "SUBSTITUTE", "REPLACE", "TEXTBEFORE", "TEXTAFTER", "CHAR", "UNICHAR", "T"}
    | {"ADDRESS", "FORMULATEXT", "VALUETOTEXT", "ARRAYTOTEXT", "REGEXEXTRACT"}
    | {"REGEXREPLACE", "TEXTSPLIT"}
)
# Functions that can return a reference or an array: joined, more than one value.
_MULTI_RESULT = frozenset(
    {"OFFSET", "INDIRECT", "INDEX", "XLOOKUP", "FILTER", "SORT", "SORTBY", "UNIQUE"}
    | {"SEQUENCE", "TRANSPOSE", "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS", "VSTACK"}
    | {"HSTACK", "TOCOL", "TOROW", "EXPAND"}
)
# Branches return one of their arguments: more than one value only if an argument is.
_PASS_THROUGH = frozenset({"IF", "IFS", "SWITCH", "CHOOSE", "IFERROR", "IFNA"})
# Excel writes a number as text positionally, with at most 15 significant digits, at
# least in this range; outside it, it may switch to scientific notation.
_GENERAL_MIN = 1e-4
_GENERAL_MAX = 1e15
_FORMAT_TOKEN = re.compile(r'"[^"]*"|\\.|_.|\*.|\[[^\]]*\]|;|[^"\\_*\[;]+', re.S)
_BRACKET_OK = re.compile(r"^\[(?:\$[^\]]*|[A-Za-z]+|Color[0-9]+)\]$", re.I)
_ELAPSED = re.compile(r"^\[(?:h+|m+|s+)\]$", re.I)
_DATE_CODE = re.compile(r"AM/PM|A/P|[yYmMdDhHsS]")
_PLACEHOLDER = re.compile(r"[0#?]")
_TRAILING_COMMAS = re.compile(r"[0#?](,+)[^0#?]*$")
_POINT_COMMAS = re.compile(r"[0#?](,+)\.")
# A range joined as text is read one probe per cell that can hold a number Excel
# writes differently; past this many, the join is refused for the policies that
# evaluate it.
JOIN_RANGE_PROBES = 1000

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
    number_text = "number_text"


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
    join_uses: list[JoinUse] = field(default_factory=list)
    """Joins that read a shared source probe (one per source cell, planned last)."""
    over_cap: int = 0
    """Probes left unplaced only because the cap was reached: a cap of
    ``len(probes) + over_cap`` holds them all."""


@dataclass(slots=True, frozen=True)
class JoinUse:
    """A join reading one source cell: the joining cell, the arms that guard the join,
    the check (``general`` or ``criterion``) and the index of the source's probe."""

    cell: CellRef
    guards: tuple[Guard, ...]
    check: str
    source: int


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
    number_text: dict[CellRef, set[str]] = field(default_factory=dict)
    """Cells that turned a number into text where Excel's text may differ, and why."""


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


# -- number to text --------------------------------------------------------------


def _format_sections(fmt: str) -> list[str]:
    """The sections of a number format, with quoted text, escapes, fills, colours and
    locales blanked out: what is left decides how a number is shown.

    Raises:
        ValueError: for a condition (``[>100]``), more than four sections, or text the
            scan does not read.
    """
    tokens = _FORMAT_TOKEN.findall(fmt)
    if "".join(tokens) != fmt:
        raise ValueError(f"cannot read number format {fmt!r}")
    sections: list[list[str]] = [[]]
    for token in tokens:
        if token == ";":
            sections.append([])
        elif token[0] in '"\\_*':
            sections[-1].append(" ")
        elif token.startswith("["):
            if _ELAPSED.match(token):
                sections[-1].append("h")
            elif not _BRACKET_OK.match(token):
                raise ValueError(f"format condition {token!r}")
        else:
            sections[-1].append(token)
    if len(sections) > 4:
        raise ValueError(f"cannot read number format {fmt!r}")
    return ["".join(s) for s in sections]


def number_format_rounding(fmt: str) -> tuple[int, float] | None:
    """``(decimals, scale)`` for a ``TEXT`` format: the number shown is ``x * scale``
    rounded to ``decimals`` places (``%`` multiplies by 100, a comma after the last
    digit placeholder divides by 1,000). None when the format shows no number: a date,
    a time, or only literal text.

    Raises:
        ValueError: for a format whose rounding the scan does not read: ``General``,
            ``@``, scientific, fractions, fractional seconds, conditions, sections that
            round differently, or letters that are not format codes.
    """
    found: set[tuple[int, float]] = set()
    # The third section shows an exact zero only, which sits on no tie: it does not
    # take part (the fourth is for text).
    for section in _format_sections(fmt)[:2]:
        upper = section.upper()
        if "GENERAL" in upper or "@" in section:
            raise ValueError(f"cannot read number format {fmt!r}")
        if "E+" in upper or "E-" in upper:
            raise ValueError(f"scientific format {fmt!r}")
        if _DATE_CODE.search(section):
            if re.search(r"[sS]\.0", section):
                raise ValueError(f"fractional seconds in {fmt!r}")
            continue
        if re.search(r"[A-Za-z]", section):
            raise ValueError(f"cannot read number format {fmt!r}")
        if not _PLACEHOLDER.search(section):
            continue
        if "/" in section:
            raise ValueError(f"fraction format {fmt!r}")
        point = section.find(".")
        decimals = len(_PLACEHOLDER.findall(section[point + 1 :])) if point >= 0 else 0
        m = _TRAILING_COMMAS.search(section) or _POINT_COMMAS.search(section)
        commas = len(m.group(1)) if m is not None else 0
        found.add((decimals, 100.0 ** section.count("%") / 1000.0**commas))
    if len(found) > 1:
        raise ValueError(f"sections of {fmt!r} round differently")
    return found.pop() if found else None


def excel_general_text(x: float) -> str | None:
    """The text Excel makes of ``x`` (``&``, ``CONCAT``, ``LEN``): at most 15
    significant digits. None outside 0.0001 to 1e15, where Excel may write scientific
    notation."""
    if x == 0:
        return "0"
    if not (math.isfinite(x) and _GENERAL_MIN <= abs(x) < _GENERAL_MAX):
        return None
    text = format(Decimal(format(x, ".15g")), "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def text_drops_sign(x: float, decimals: int, scale: float) -> bool:
    """True when IronCalc 0.9.1's ``TEXT`` may write ``x`` without its minus sign under
    a one-section format: ``x`` is negative and, shown at ``decimals`` places after
    ``scale``, rounds to zero or one unit of the last place (``TEXT(-1,"0")`` is
    ``"1"``, ``TEXT(-0.01,"0%")`` is ``"1%"``). Excel keeps the sign (``"-1"``)."""
    if not x < 0:
        return False
    return abs(x * scale) * 10.0 ** min(decimals, 300) < 1.5


def plain_number(x: float) -> bool:
    """True when Excel's text for ``x`` (15 significant digits) is the engine's text:
    ``x`` needs no more digits, sits in the positional range and is not ``-0``."""
    if x == 0:
        return math.copysign(1.0, x) > 0
    text = excel_general_text(x)
    return text is not None and float(text) == x


def _shows_digits(fmt: ASTNode) -> bool:
    """False for a literal ``TEXT`` format that shows no number (a date, a time, only
    text); True otherwise, unreadable and non-literal formats included."""
    text = literal(fmt)
    if not isinstance(text, str):
        return True
    try:
        return number_format_rounding(text) is not None
    except ValueError:
        return True


def _maybe_number(node: ASTNode) -> bool:
    """False when ``node`` cannot be a computed number: a literal, or text."""
    if literal(node) is not UNKNOWN:
        return False
    if isinstance(node, BinaryNode) and node.operator == "&":
        return False
    return not (isinstance(node, FunctionNode) and function_name(node.name) in _TEXT_RESULT)


def _criteria_args(name: str, n: int) -> range:
    """Positions of the criteria of a SUMIF-family function with ``n`` arguments."""
    if name in {"SUMIF", "COUNTIF", "AVERAGEIF"}:
        return range(1, min(n, 2))
    if name == "COUNTIFS":
        return range(1, n, 2)
    if name in {"SUMIFS", "AVERAGEIFS", "MAXIFS", "MINIFS"}:
        return range(2, n, 2)
    return range(0)


def _joined(node: ASTNode) -> list[ASTNode]:
    """The parts a criterion joins into text: the operands of ``&`` and ``CONCAT``."""
    if isinstance(node, BinaryNode) and node.operator == "&":
        return [*_joined(node.left), *_joined(node.right)]
    if isinstance(node, FunctionNode) and function_name(node.name) in _JOINS:
        return [part for arg in node.args for part in _joined(arg)]
    return [node]


def _number_text_reason(probe: Probe, values: list[EngineValue]) -> str | None:
    """Why the engine's text for this conversion may not be Excel's; None when it is."""
    if probe.meta["check"] == "always":
        # A join calc cannot read cell by cell: refused wherever it is evaluated.
        return NUMBER_TEXT_NOT_CHECKED
    if probe.meta["check"] != "tie" and len(values) > 2:
        # A joined function result: read like a number only when it is one value.
        size = values[2]
        if not isinstance(size, float):
            return None
        if size != 1:
            return NUMBER_TEXT_NOT_CHECKED
    x = values[0]
    if not isinstance(x, float):
        return None
    if probe.meta["check"] == "tie":
        if "decimals" in probe.meta:
            decimals, scale = int(probe.meta["decimals"]), float(probe.meta["scale"])
        else:
            given = values[1] if len(values) > 1 else 2.0
            if not isinstance(given, float):
                return None
            decimals, scale = math.trunc(given), 1.0
        y = x * scale
        if not math.isfinite(y) or abs(y) * 10.0 ** min(decimals, 300) >= _GENERAL_MAX:
            # Excel shows 15 significant digits and pads with zeros; the engine does not.
            return NUMBER_TEXT_DIGITS
        if probe.meta.get("sections") == "1" and text_drops_sign(x, decimals, scale):
            return NUMBER_TEXT_NEGATIVE
        return NUMBER_TEXT_TIE if near_flip("ROUND", y, decimals) else None
    return _general_reason(values, probe.meta["check"])


def _general_reason(values: list[EngineValue], check: str) -> str | None:
    """For ``(x, x&"")`` readings: why Excel's text for ``x`` may differ; None if not."""
    x, text = values[0], values[1]
    if not isinstance(x, float) or not isinstance(text, str) or text == "":
        # An empty cell reads 0 but joins as "" in Excel and the engine alike.
        return None
    differs = excel_general_text(x) != text
    if check == "criterion" and (differs or x != math.floor(x)):
        return NUMBER_TEXT_CRITERION
    return NUMBER_TEXT_DIGITS if differs else None


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


def _fixed(formulas: tuple[str, ...], meta: dict[str, str]) -> Builder:
    """A builder for formulas known in advance."""
    return lambda: (formulas, dict(meta))


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
    inputs: frozenset[CellRef] = frozenset()
    """Mapped input cells: constants that change per policy."""
    sources: dict[CellRef, list[tuple[CellRef, tuple[Guard, ...], str]]] = field(
        default_factory=dict
    )
    """Join source cell -> the joins reading it: (joining cell, guards, check)."""
    ranges: dict[tuple[str, str], list[CellRef] | None] = field(default_factory=dict)
    """(formula sheet, reference) -> the cells a joined range holds that need a probe."""
    join_uses: list[JoinUse] = field(default_factory=list)
    over_cap: int = 0
    criteria: set[int] = field(default_factory=set)
    """``id`` of the nodes a SUMIF-family criterion joins as text, in this cell."""

    def lift(self, node: ASTNode) -> str:
        return anchor(node, self.cell.sheet, names=self.graph, as_of=self.as_of)

    def add(self, kind: ProbeKind, scoped: bool, build: Builder) -> None:
        ordinal = self.counters.get(kind, 0)
        self.counters[kind] = ordinal + 1
        if scoped or len(self.probes) >= self.cap:
            self.over_cap += not scoped
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
            if node.operator == "&":
                for operand in (node.left, node.right):
                    self._to_text(operand, blocked, joins=False)
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
        if built_range(node.name) is not None:
            # ``B1:OFFSET(...)``: a reference from B1 to where the call lands.
            self.add(ProbeKind.dynamic_ref, scoped, lambda: self._built_range(node))
            return
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
        self._probe_text(node, name, scoped)

    def _probe_text(self, node: FunctionNode, name: str, scoped: bool) -> None:
        """Number-to-text probes: a formatter's rounding, a join's or a text function's
        conversion; and, first, which parts a SUMIF-family criterion joins."""
        args = node.args
        for i in _criteria_args(name, len(args)):
            self.criteria.update(id(part) for part in _joined(args[i]))
        if name in _FORMATTERS and args and _maybe_number(args[0]):
            if name != "TEXT" or (len(args) >= 2 and _shows_digits(args[1])):
                self.add(ProbeKind.number_text, scoped, lambda: self._formatter(node, name))
        elif name in _JOINS:
            for arg in args[_JOINS[name] :]:
                if self._multi(arg, joins=True):
                    self._join_many(arg, scoped)
                else:
                    self._to_text(arg, scoped, joins=True)
        elif name in TEXT_ARGUMENTS:
            for i in TEXT_ARGUMENTS[name]:
                if i < len(args):
                    self._to_text(args[i], scoped, joins=False)

    def _to_text(self, node: ASTNode, scoped: bool, *, joins: bool) -> None:
        """Probe ``node``'s conversion to text, unless it cannot be a number."""
        if not _maybe_number(node):
            return
        check = "criterion" if id(node) in self.criteria else "general"
        self.add(ProbeKind.number_text, scoped, lambda: self._general(node, check, joins))

    def _join_many(self, node: ASTNode, scoped: bool) -> None:
        """A range or array joined as text.

        Constants Excel writes as the engine does need nothing. Every other cell of a
        range is a *source*: probed once per plan, after every cell's own probes, and
        read per policy for each join under that join's guards (:meth:`place_sources`).
        A function result is read with its size, so one value is checked like a number
        and an array refused. A join calc cannot read cell by cell (a reference it
        cannot resolve, or more than ``JOIN_RANGE_PROBES`` cells to read) is refused
        per policy, under its arm's guards.
        """
        check = "criterion" if id(node) in self.criteria else "general"
        if isinstance(node, FunctionNode):
            self.add(ProbeKind.number_text, scoped, lambda: self._sized(node, check))
            return
        todo = self._sources_of(node) if isinstance(node, RefNode) else None
        if todo is None or len(todo) > JOIN_RANGE_PROBES:
            self.add(ProbeKind.number_text, scoped, _fixed(("TRUE",), {"check": "always"}))
            return
        if not todo:
            return
        if scoped:
            # Inside an array formula or LET: the cells cannot be read for this join.
            self.add(ProbeKind.number_text, True, _fixed(("TRUE",), {"check": "always"}))
            return
        for source in todo:
            self.sources.setdefault(source, []).append((self.cell, self.guards, check))

    def _sources_of(self, node: RefNode) -> list[CellRef] | None:
        """The cells of a joined range that need a probe; cached per sheet and text, so
        a join copied down a column resolves its range once."""
        key = (self.cell.sheet, node.ref)
        if key not in self.ranges:
            cells = self._range_cells(node)
            self.ranges[key] = (
                None if cells is None else [c for c in cells if not self._plain_constant(c)]
            )
        return self.ranges[key]

    def place_sources(self) -> None:
        """Plan the join sources, after every cell's own probes so they never displace
        them. A join with a source past the cap is refused for its own cell only."""
        lost: set[CellRef] = set()
        for source in sorted(self.sources):
            uses = self.sources[source]
            if len(self.probes) >= self.cap:
                self.over_cap += 1
                lost.update(cell for cell, _guards, _check in uses)
                continue
            index = len(self.probes)
            ref = f"{quote_sheet(source.sheet)}!${_number_to_col(source.col)}${source.row}"
            formulas = (ref, f'({ref})&""')
            meta = {"check": "source"}
            self.probes.append(Probe(ProbeKind.number_text, source, "", 0, formulas, meta))
            self.join_uses.extend(
                JoinUse(cell, guards, check, index) for cell, guards, check in uses
            )
        self.sources.clear()
        for cell in sorted(lost):
            self.unplaced.setdefault(cell, []).append(ProbeKind.number_text)

    def _range_cells(self, node: RefNode) -> list[CellRef] | None:
        """Every cell a range or multi-cell name covers (a whole column or row cut to
        the used range); None when it cannot be read statically."""
        m = _REF.match(node.ref)
        if m is None:
            return None
        if m.group("b") is None:
            found = self.graph.lookup_name(node.ref, self.cell.sheet)
            if found is None or found.kind != "cells":
                return None
            return self.graph.name_cells(node.ref, self.cell.sheet)
        s1, s2 = m.group("s1"), m.group("s2")
        if s2 is not None and s2 != s1:
            return None
        sheet = _unquote(s1) if s1 is not None else self.cell.sheet
        target = f"{quote_sheet(sheet)}!{m.group('a')}:{m.group('b')}"
        try:
            return self.graph.resolve_target(target)
        except CalcInputError:
            return None

    def _plain_constant(self, cell: CellRef) -> bool:
        """A cell no policy changes whose text Excel writes as the engine does: blank,
        text, a boolean, or a number that needs at most 15 significant digits. A cell of
        an array formula's range is computed, however its stored value looks, and so is
        an empty cell a dynamic array can spill into."""
        if cell in self.inputs or self.graph.formula(cell) is not None:
            return False
        if self.graph.array_anchor(cell) is not None:
            return False
        if not self.graph.has_stored(cell):
            # Blank, unless a dynamic array's spill can grow into it for some policy.
            return not self.graph.may_spill_into(cell)
        value = self.graph.stored(cell)
        return not isinstance(value, float) or plain_number(value)

    # -- builders --

    def _formatter(self, node: FunctionNode, name: str) -> tuple[tuple[str, ...], dict[str, str]]:
        x = self.lift(node.args[0])
        if name == "TEXT":
            fmt = literal(node.args[1])
            if not isinstance(fmt, str):
                raise ValueError("TEXT format is not literal text")
            rounding = number_format_rounding(fmt)  # raises for an unreadable format
            if rounding is None:
                raise ValueError(f"format {fmt!r} shows no number")
            decimals, scale = rounding
            sections = str(len(_format_sections(fmt)))
            meta = {"check": "tie", "function": name, "decimals": str(decimals)}
            return (x,), {**meta, "scale": repr(scale), "sections": sections}
        meta = {"check": "tie", "function": name}
        if len(node.args) < 2:
            return (x,), meta
        return (x, self.lift(node.args[1])), meta

    def _sized(self, node: ASTNode, check: str) -> tuple[tuple[str, ...], dict[str, str]]:
        x = self.lift(node)
        return (x, f'({x})&""', f"ROWS({x})*COLUMNS({x})"), {"check": check}

    def _general(
        self, node: ASTNode, check: str, joins: bool
    ) -> tuple[tuple[str, ...], dict[str, str]]:
        if self._multi(node, joins):
            raise ValueError("a range or array converted to text")
        x = self.lift(node)
        return (x, f'({x})&""'), {"check": check}

    def _multi(self, node: ASTNode, joins: bool) -> bool:
        """True when ``node`` may hold more than one value: a range, a multi-cell name
        or, as a joined argument, a function that can return a reference or an array."""
        if isinstance(node, FunctionNode):
            if built_range(node.name) is not None:
                return True
            name = function_name(node.name)
            if name in _PASS_THROUGH:
                return any(self._multi(arg, joins) for arg in node.args[1:])
            return joins and name in _MULTI_RESULT
        if not isinstance(node, RefNode):
            return False
        m = _REF.match(node.ref)
        if m is None or m.group("b") is not None:
            return True
        if _end(m.group("a"))[0] != "name":
            return False
        found = self.graph.lookup_name(node.ref, self.cell.sheet)
        if found is None or found.kind != "cells":
            return found is not None
        cells = sum((r2 - r1 + 1) * (c2 - c1 + 1) for _s, r1, c1, r2, c2 in found.rects)
        return found.whole or cells > 1

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
        """XMATCH text for an XLOOKUP / XMATCH: the position its own modes find.

        The position counts from the start of the lookup array whichever way the search
        runs, so with search_mode -1 "first" is still the array's first row and "last"
        its last (a recorded decision: both edges must be exercised on their own either
        way, so swapping the labels would change no status).
        """
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
        return _static_sheet(self.graph, self.cell, node)

    def _built_range(self, node: FunctionNode) -> tuple[tuple[str, ...], dict[str, str]]:
        """Both corners of a range built by a call (``B1:OFFSET(...)``): the start's
        and the call's first cell and size. The range is the smallest rectangle holding
        both, as Excel's ``:`` makes it."""
        parts = built_range(node.name)
        if parts is None or "!" in parts[1]:
            raise ValueError(f"cannot read range end {node.name!r}")
        start_text, function = parts
        start = RefNode(start_text)
        sheet = self._static_sheet(start)
        end = FunctionNode(function, node.args)
        end_sheet: str | None = None
        if function_name(function) in {"OFFSET", "INDEX"} and node.args:
            end_sheet = self._static_sheet(node.args[0])
        elif function_name(function) == "INDIRECT" and node.args:
            end_sheet = _text_sheet(self.graph, node.args[0])
        if end_sheet is not None and end_sheet != sheet:
            # Excel's ``:`` cannot join references on two sheets (``#VALUE!``).
            raise ValueError(f"range {node.name!r} spans two sheets")
        formulas: list[str] = []
        for ref in (self.lift(start), self.lift(end)):
            corner = f"OFFSET({ref},0,0,1,1)"
            formulas += [f"ROW({corner})", f"COLUMN({corner})", f"ROWS({ref})", f"COLUMNS({ref})"]
        return tuple(formulas), {"function": _RANGE, "sheet": sheet}


def _unquote(sheet: str) -> str:
    if len(sheet) >= 2 and sheet.startswith("'") and sheet.endswith("'"):
        return sheet[1:-1].replace("''", "'")
    return sheet


def _static_sheet(graph: WorkbookGraph, cell: CellRef, node: ASTNode) -> str:
    """The sheet a reference expression in ``cell``'s formula lands on, when knowable
    without running it.

    Raises:
        ValueError: when it is not (the expression is not a reference, or a name over
            several sheets or a whole column or row).
    """
    if isinstance(node, FunctionNode) and node.args:
        if function_name(node.name) in {"OFFSET", "INDEX"}:
            return _static_sheet(graph, cell, node.args[0])
    if not isinstance(node, RefNode):
        raise ValueError("OFFSET base is not a static reference")
    m = _REF.match(node.ref)
    if m is None:
        raise ValueError(f"cannot lift reference {node.ref!r}")
    if _end(m.group("a"))[0] == "name" and m.group("b") is None:
        found = graph.lookup_name(node.ref, cell.sheet)
        sheets = {r[0] for r in found.rects} if found is not None else set()
        if found is None or found.kind != "cells" or found.whole or len(sheets) != 1:
            # A whole column or row would land on up to a million cells per reading.
            raise ValueError(f"name {node.ref!r} is not a bounded range on one sheet")
        return sheets.pop()
    s1 = m.group("s1")
    if s1 is None:
        return cell.sheet
    written = _unquote(s1)
    return graph.sheet(written) or written


def _text_sheet(graph: WorkbookGraph, text: ASTNode) -> str | None:
    """The sheet an ``INDIRECT`` text names literally: its leading text literal holds
    ``Sheet!`` (``"Rates!B"&row``). None when the text names no sheet that way."""
    node = text
    while True:
        if isinstance(node, BinaryNode) and node.operator == "&":
            node = node.left
        elif (
            isinstance(node, FunctionNode)
            and function_name(node.name) in {"CONCAT", "CONCATENATE"}
            and node.args
        ):
            node = node.args[0]
        else:
            break
    value = literal(node)
    if not isinstance(value, str) or "!" not in value:
        return None
    return graph.sheet(_unquote(value.partition("!")[0]))


def reach_sheets(graph: WorkbookGraph, cell: CellRef) -> set[str]:
    """The sheets the dynamic references in ``cell``'s formula can land on, where that
    is known without running them: an ``OFFSET``'s starting sheet, the sheet
    an ``INDIRECT`` text names literally, a built range's sheet. Other ``INDIRECT``
    texts can land anywhere; they are not listed."""
    text = graph.formula(cell)
    if text is None or DYNAMIC_HINT.search(text) is None:
        return set()
    root = parse_formula(text)
    if root is None:
        return set()
    out: set[str] = set()
    for node in iter_nodes(root):
        if not isinstance(node, FunctionNode) or not node.args:
            continue
        try:
            parts = built_range(node.name)
            if parts is not None:
                out.add(_static_sheet(graph, cell, RefNode(parts[0])))
                continue
            name = function_name(node.name)
            if name == "OFFSET":
                out.add(_static_sheet(graph, cell, node.args[0]))
            elif name == "INDIRECT" and (sheet := _text_sheet(graph, node.args[0])) is not None:
                out.add(sheet)
        except ValueError:
            continue
    return {sheet for sheet in out if graph.sheet(sheet) is not None}


def probe_kinds(graph: WorkbookGraph, cell: CellRef) -> list[ProbeKind]:
    """The probes ``cell``'s formula would get, placed or not."""
    plan = plan_probes(graph, [cell], cap=1 << 30)
    return [p.kind for p in plan.probes] + plan.unplaced.get(cell, [])


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
    graph: WorkbookGraph,
    cells: Iterable[CellRef],
    *,
    cap: int,
    as_of: date | None = None,
    inputs: Iterable[CellRef] = (),
    extra: Iterable[CellRef] = (),
) -> ProbePlan:
    """Plan probes for the branch, lookup, rounding, comparison, dynamic-ref and
    number-to-text nodes of each cell's formula; cells are taken in sorted order, then
    the ``extra`` cells not among them (so they are the first to go past ``cap``).
    ``as_of`` is the run's ``--as-of`` date, which lifted ``TODAY()``/``NOW()`` read;
    ``inputs`` are the mapped input cells (constants a policy changes).

    A node that :func:`anchor` cannot lift, any node beneath ``LET``/``LAMBDA``,
    every node of an array formula or inside an always-array function's arguments
    (``SUMPRODUCT``...), and every probe beyond ``cap`` is recorded in ``unplaced`` (one entry per probe).
    Unparseable formulas produce no probes (the pattern scan refuses them).
    """
    planner = _Planner(graph, cap, as_of, inputs=frozenset(inputs))
    first = set(cells)
    for group in (sorted(first), sorted(set(extra) - first)):
        for cell in group:
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
            planner.criteria = set()
            planner.visit(root, False, graph.is_array_formula(cell))
        # Each group's join sources go after its cells' own probes.
        planner.place_sources()
    sheets = {name.lower(): name for name in graph.sheets()}
    return ProbePlan(planner.probes, planner.unplaced, sheets, planner.join_uses, planner.over_cap)


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
    if probe.meta["function"] == _RANGE:
        return _range_cells(probe.meta["sheet"], values)
    nums = [v for v in values if isinstance(v, float)]
    if len(nums) != 4 or nums[2] < 1 or nums[3] < 1:
        return None
    row, col, rows, cols = (int(v) for v in nums)
    return _rect(probe.meta["sheet"], row, col, rows, cols)


def _range_cells(sheet: str, values: list[EngineValue]) -> list[CellRef] | None:
    """The rectangle spanning a built range's two corners (first cell and size each)."""
    nums = [v for v in values if isinstance(v, float)]
    if len(nums) != 8 or min(nums) < 1:
        return None
    tops, lefts, bottoms, rights = [], [], [], []
    for row, col, rows, cols in (nums[:4], nums[4:]):
        tops.append(int(row))
        lefts.append(int(col))
        bottoms.append(int(row + rows) - 1)
        rights.append(int(col + cols) - 1)
    top, left = min(tops), min(lefts)
    return _rect(sheet, top, left, max(bottoms) - top + 1, max(rights) - left + 1)


def _live(probe: Probe, outcomes: dict[tuple[CellRef, int], str]) -> bool | None:
    """Whether the cell evaluates the probe's part: True, False, or None (unknown)."""
    return _live_at(probe.cell, probe.guards, outcomes)


def _live_at(
    cell: CellRef, guards: tuple[Guard, ...], outcomes: dict[tuple[CellRef, int], str]
) -> bool | None:
    """Whether ``cell`` evaluates a part under ``guards``: True, False or None."""
    live: bool | None = True
    for ordinal, taken, maybe in guards:
        outcome = outcomes.get((cell, ordinal))
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
        elif probe.kind is ProbeKind.number_text:
            if probe.meta["check"] == "source":
                continue  # read below, once for each join that reads it
            detail = _number_text_reason(probe, values)
            if detail is not None:
                readings.number_text.setdefault(probe.cell, set()).add(detail)
        else:
            landed = _dynamic_cells(probe, values, plan.sheets)
            dynamic.setdefault(probe.cell, set()).update(landed or ())
            if landed is None:
                readings.dynamic_unread.add(probe.cell)
    for use in plan.join_uses:
        if _live_at(use.cell, use.guards, outcomes) is False:
            continue
        detail = _general_reason(values_of[use.source], use.check)
        if detail is not None:
            readings.number_text.setdefault(use.cell, set()).add(detail)
    readings.dynamic_refs = {cell: sorted(cells) for cell, cells in dynamic.items()}
    return readings
