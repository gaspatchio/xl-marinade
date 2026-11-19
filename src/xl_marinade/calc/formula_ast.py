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

FN_PREFIXES = ("_XLFN.", "_XLWS.", "_XLL.")

SHEET = r"(?:'(?:[^']|'')+'|[^'!:\s{}();,\"#\[\]]+)"
"""A sheet prefix as written in a reference (quoted or bare), without the ``!``."""
CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]+"
NAME = r"[A-Za-z_\\][\w.\\?]*"
END = rf"(?:{CELL}|\$?[A-Za-z]{{1,3}}|\$?[0-9]+|{NAME})"
"""One end of a reference: a cell, a column, a row or a name."""
# What a well-formed RefNode may look like; anything else is a mangled parse.
VALID_REF = re.compile(rf"^(?:{SHEET}!)?{END}(?::(?:{SHEET}!)?{END})?$")
STRUCTURED = re.compile(r"^[A-Za-z_\\][\w.\\]*\[.*\]$")

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

# Error literals the core parser splits at ``/`` (``#N/A`` -> ``"#N" / A``).
_SPLIT_ERROR = re.compile(r"#(?:N/A|DIV/0!)", re.IGNORECASE)
# Text where ``#`` is not an error literal: strings, quoted sheets, structured refs.
_SKIPPED = re.compile(r"\"(?:[^\"]|\"\")*\"|'(?:[^']|'')*'|\[[^\]]*\]")

UNKNOWN = object()
"""Sentinel returned by :func:`literal` for a non-literal argument."""


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
    if VALID_REF.match(ref):
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
    parser = FormulaParser()
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
