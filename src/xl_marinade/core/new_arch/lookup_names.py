# ABOUTME: Rewrites defined names and LET variables in a parsed formula into the
# ABOUTME: references they stand for, so lookup resolution sees plain ranges.

"""Name expansion for semantic lookup resolution.

The lookup resolver (``ResolutionEngine`` and the INDEX strategy chain) reads
cell values by address. It cannot dereference a defined name, so a lookup
written against names --

    INDEX(MP_Age, MATCH(PointID, MP_ID, 0))

-- fell back to ``partial_resolved`` with the name itself as the "resolved"
ref, dropped the key drivers, and emitted ``UNRESOLVED:MP_Age``, while the same
lookup written with ranges resolved to the exact row. Every lookup in the
lifelib example workbooks is written against names.

``expand_names`` returns a copy of the AST in which:

- a ``Ref`` naming a defined name that refers to exactly one contiguous range
  is replaced by that range (``MP_Age`` -> ``Model_Points!B2:B10001``);
- a ``LET`` is replaced by its calculation with each bound variable inlined,
  so ``LET(k, B1, INDEX(MP_Age, k))`` resolves as ``INDEX(MP_Age, B1)``.

Formula-valued (dynamic), external and multi-area names are left untouched for
the existing fallback. The rewrite only feeds lookup resolution; the stored
formula and the static reference edges are unaffected.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from typing import Any

# (name, sheet the formula is on) -> the name's A1 ranges, or None.
NameResolver = Callable[[str, str], list[str] | None]

# A reference as the parser leaves a name: an optional sheet qualifier and an
# identifier. Excel forbids names that look like cells, so a cell address that
# matches this (A1, R1C1) simply fails to resolve as a name and is kept.
_NAME_REF = re.compile(
    r"^(?:(?P<sheet>'(?:[^']|'')+'|[^'!]+)!)?(?P<name>[A-Za-z_\\][A-Za-z0-9_.\\]*)$"
)


def expand_names(node: Any, sheet_name: str, resolve: NameResolver) -> Any:
    """Return `node` with defined names and LET variables expanded (see module doc)."""
    return _expand(node, sheet_name, resolve, {})


def _expand(node: Any, sheet_name: str, resolve: NameResolver, env: dict[str, Any]) -> Any:
    if isinstance(node, list):
        return [_expand(item, sheet_name, resolve, env) for item in node]
    if not isinstance(node, dict):
        return node
    if node.get("type") == "Ref":
        return _expand_ref(node, sheet_name, resolve, env)
    if node.get("type") == "Function" and _function_name(node) == "LET":
        inlined = _inline_let(node, sheet_name, resolve, env)
        if inlined is not None:
            return inlined
    return {key: _expand(value, sheet_name, resolve, env) for key, value in node.items()}


def _function_name(node: dict[str, Any]) -> str:
    return str(node.get("name", "")).upper().removeprefix("_XLFN.")


def _expand_ref(
    node: dict[str, Any], sheet_name: str, resolve: NameResolver, env: dict[str, Any]
) -> Any:
    match = _NAME_REF.match(str(node.get("ref", "")))
    if not match:
        return node
    name, sheet = match.group("name"), match.group("sheet")
    if sheet is None and name.lower() in env:
        return env[name.lower()]  # a LET variable shadows any defined name
    scope = sheet.strip("'").replace("''", "'") if sheet else sheet_name
    ranges = resolve(name, scope)
    if not ranges or len(ranges) != 1:
        return node
    return {**node, "ref": ranges[0]}


def _inline_let(
    node: dict[str, Any], sheet_name: str, resolve: NameResolver, env: dict[str, Any]
) -> Any | None:
    """LET(n1, v1, ..., calc) -> calc with each n_i replaced by its expanded v_i."""
    args = node.get("args") or []
    if len(args) < 3 or len(args) % 2 == 0:
        return None
    scope = dict(env)
    for name_node, value_node in zip(args[:-1:2], args[1:-1:2], strict=True):
        if not (isinstance(name_node, dict) and name_node.get("type") == "Ref"):
            return None
        # Bindings are sequential: each value may use the variables before it.
        scope[str(name_node.get("ref", "")).lower()] = _expand(
            value_node, sheet_name, resolve, scope
        )
    return _expand(args[-1], sheet_name, resolve, scope)
