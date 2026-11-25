"""Shared helpers over the core formula parser's AST: parsing, names and literals."""

from __future__ import annotations

import re
from collections.abc import Iterator

from xl_marinade.core.parser import (
    ASTNode,
    BinaryNode,
    ConstNode,
    FormulaParser,
    FunctionNode,
    RefNode,
    UnaryNode,
)

FN_PREFIXES = ("_XLFN.", "_XLWS.", "_XLL.", "_XLUDF.")

SHEET = r"(?:'(?:[^']|'')+'|[^'!:\s{}();,\"#\[\]]+)"
"""A sheet prefix as written in a reference (quoted or bare), without the ``!``."""
CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]+"
NAME = r"[A-Za-z_\\][\w.\\?]*"
END = rf"(?:{CELL}|\$?[A-Za-z]{{1,3}}|\$?[0-9]+|{NAME})"
"""One end of a reference: a cell, a column, a row or a name."""
# What a well-formed RefNode may look like; anything else is a mangled parse.
VALID_REF = re.compile(rf"^(?:{SHEET}!)?{END}(?::(?:{SHEET}!)?{END})?$")
STRUCTURED = re.compile(r"^[A-Za-z_\\][\w.\\]*\[.*\]$")
EXTERNAL = re.compile(r"^(?:'\[[0-9]+\]|\[[0-9]+\])")
"""A reference into another workbook (``[1]Sheet!A1``, ``'[1]My Sheet'!A1``)."""
_EXTERNAL_REF = re.compile(rf"^\[[0-9]+\]{SHEET}!{END}(?::{END})?$")

# Functions whose arguments are array-evaluated in every Excel version.
ALWAYS_ARRAY = frozenset(
    {
        "SUMPRODUCT",
        "MMULT",
        "MDETERM",
        "MINVERSE",
        "TRANSPOSE",
        "FREQUENCY",
        "LINEST",
        "LOGEST",
        "TREND",
        "GROWTH",
    }
)
ERROR_CODES = frozenset(
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

# Excel worksheet functions IronCalc 0.9.1 does not know: it answers #NAME?, an Excel
# error code, so an IFERROR / ISERROR / ISNUMBER around the call would turn the gap into
# a plausible value. Built by calling every Excel function bare in the engine (a known
# one gives an argument error); tests/calc/test_engine_gaps.py pins it to the installed
# IronCalc. At run time calc also asks the installed engine about every function the
# workbook calls (engine.functions_engine_lacks), which adds VBA, add-in and newer
# functions and those that answer #N/IMPL! (FORECAST.ETS); patterns._shape_gap covers
# the calls that answer #N/IMPL! only for some arguments.
ENGINE_LACKS = frozenset(
    {
        "AGGREGATE",
        "BAHTTEXT",
        "CALL",
        "CUBEKPIMEMBER",
        "CUBEMEMBER",
        "CUBEMEMBERPROPERTY",
        "CUBERANKEDMEMBER",
        "CUBESET",
        "CUBESETCOUNT",
        "CUBEVALUE",
        "DBCS",
        "DETECTLANGUAGE",
        "EUROCONVERT",
        "FIELDVALUE",
        "FILTERXML",
        "GETPIVOTDATA",
        "GROUPBY",
        "IMAGE",
        "JIS",
        "PHONETIC",
        "PIVOTBY",
        "REGISTER.ID",
        "RTD",
        "SQL.REQUEST",
        "STOCKHISTORY",
        "TRANSLATE",
        "WEBSERVICE",
    }
)

# Error literals the core parser splits at ``/`` (``#N/A`` -> ``"#N" / A``).
_SPLIT_ERROR = re.compile(r"#(?:N/A|DIV/0!)", re.IGNORECASE)
# Text where ``#`` is not an error literal: strings, quoted sheets, structured refs.
_SKIPPED = re.compile(r"\"(?:[^\"]|\"\")*\"|'(?:[^']|'')*'|\[[^\]]*\]")

# Text that may hold a dynamic reference: OFFSET, INDIRECT, or a ``:`` before a call
# (a range built by one, ``B1:INDEX(...)``). A cheap filter before parsing.
DYNAMIC_HINT = re.compile(r"OFFSET|INDIRECT|:[^:(),]*\(", re.IGNORECASE)

UNKNOWN = object()
"""Sentinel returned by :func:`literal` for a non-literal argument."""

# A sheet name that needs no quotes: an identifier that reads as neither a cell (A1,
# XFD9), an R1C1 reference (R, C, RC, R1C1, R2C, RC3) nor a boolean.
_PLAIN_SHEET = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_CELL_LIKE = re.compile(r"^[A-Za-z]{1,3}[0-9]+$")
_R1C1_LIKE = re.compile(r"^(?:[Rr][0-9]*(?:[Cc][0-9]*)?|[Cc][0-9]*)$")


class ExcelFormulaParser(FormulaParser):
    """The core parser with Excel's operator precedence.

    Excel binds ``&`` tighter than the comparisons (``A&B="x"`` is ``(A&B)="x"``,
    ``A=B&C`` is ``A=(B&C)``); the core parser has them the other way round. Lifted
    probe text must compute what its cell computes, so calc parses with this one.
    """

    PRECEDENCE = {
        "=": 1,
        "<>": 1,
        "<": 1,
        ">": 1,
        "<=": 1,
        ">=": 1,
        "&": 2,
        "+": 3,
        "-": 3,
        "*": 4,
        "/": 4,
        "^": 5,
    }


def quote_sheet(sheet: str) -> str:
    """``sheet`` as a formula must write it before ``!``, quoted when Excel needs it.

    Unlike core's ``quote_sheet_name`` this also quotes names that start with a digit
    (``'2024'``) or read as a cell, an R1C1 reference or a boolean.
    """
    plain = (
        _PLAIN_SHEET.match(sheet)
        and not _CELL_LIKE.match(sheet)
        and not _R1C1_LIKE.match(sheet)
        and sheet.upper() not in {"TRUE", "FALSE"}
    )
    if plain:
        return sheet
    escaped = sheet.replace("'", "''")
    return f"'{escaped}'"


def function_name(raw: str) -> str:
    """Upper-case name with ``_xlfn.`` / ``_xlws.`` / ``_xll.`` prefixes removed."""
    name = raw.upper()
    stripped = True
    while stripped:
        stripped = False
        for prefix in FN_PREFIXES:
            if name.startswith(prefix):
                name = name[len(prefix) :]
                stripped = True
    return name


def built_range(raw: str) -> tuple[str, str] | None:
    """``(start, function)`` for a range whose end is a function call.

    The core parser reads ``B1:OFFSET(B1,n,0)`` as one function named ``B1:OFFSET``:
    a reference from ``B1`` to wherever the call lands. ``function`` is the call's name
    as written (it may carry a sheet, ``Tab!OFFSET``, which Excel does not allow). None
    for an ordinary function name.
    """
    start, sep, function = raw.rpartition(":")
    if not sep or not start or not function:
        return None
    return start, function


def call_name(raw: str) -> str:
    """:func:`function_name` of a call, or of the end of a range built by one."""
    built = built_range(raw)
    return function_name(built[1] if built is not None else raw)


def iter_nodes(root: ASTNode) -> Iterator[ASTNode]:
    """Every node of the tree in pre-order (a node before its children, left to right)."""
    stack: list[ASTNode] = [root]
    while stack:
        node = stack.pop()
        yield node
        if isinstance(node, FunctionNode):
            stack.extend(reversed(node.args))
        elif isinstance(node, BinaryNode):
            stack.extend([node.right, node.left])
        elif isinstance(node, UnaryNode):
            stack.append(node.operand)


def plausible_ref(ref: str) -> bool:
    """True when a RefNode's text looks like a real reference, name or structured ref."""
    if VALID_REF.match(ref) or _EXTERNAL_REF.match(ref):
        return True
    return bool(STRUCTURED.match(ref)) and ref.count("[") == ref.count("]")


def parse_formula(text: str) -> ASTNode | None:
    """Parse the formula; None when the parser raised or the parse looks mangled.

    The core parser silently stops at text it cannot place (an intersection space,
    ``INDEX(...):B3``), turns array constants into garbage references and splits
    ``#N/A`` / ``#DIV/0!`` at the slash, so a parse that left text behind, holds an
    implausible reference or contains such an error literal is rejected too.
    """
    if _SPLIT_ERROR.search(_SKIPPED.sub("", text)):
        return None
    parser = ExcelFormulaParser()
    try:
        root = parser.parse(text)
    except (ValueError, IndexError, RecursionError):
        return None
    if parser.formula[parser.pos :].strip():
        return None
    for node in iter_nodes(root):
        if isinstance(node, RefNode) and not plausible_ref(node.ref):
            return None
    return root


def literal(node: ASTNode) -> object:
    """The value of a literal argument (``TRUE``, ``0``, ``-1``), else ``UNKNOWN``."""
    if isinstance(node, ConstNode):
        return node.value
    if isinstance(node, UnaryNode) and node.operator in {"+", "-"}:
        inner = literal(node.operand)
        if isinstance(inner, int | float) and not isinstance(inner, bool):
            return -inner if node.operator == "-" else inner
    return UNKNOWN


def sign(value: object) -> int | None:
    """Sign of a numeric/boolean literal; None for text or a non-literal."""
    if isinstance(value, bool):
        return 1 if value else 0
    if isinstance(value, int | float):
        return (value > 0) - (value < 0)
    return None
