# ABOUTME: Canonical-A1 formula form for cross-position family grouping
# ABOUTME: Replaces cell/range refs with placeholders, uppercases, gates with meaningfulness

"""
Canonical-A1 Formula Grouping Primitive (Cycle 17 #312)

Produces a position-invariant canonical form of a formula by replacing
cell/range references with placeholders, uppercasing, and normalizing
whitespace. Used for family-grouping decisions in downstream analytics.

Meaningfulness gate: returns None when the canonical form contains no
parentheses — those are either pure references (=A1), pure literals
(=1), or simple operator chains (=A1+B1) which create no meaningful
"family" beyond trivial coincidence.

Design reference: docs/future_phases/phase2/calibration/
                  cycle17-canonical-a1-primitive-scoping-2026-05-11.md §A Rule B
"""

import re

# Order matters: structured table refs first (longest), then cell ranges,
# then whole-column / whole-row, then single cells.
#
# Each pattern matches a complete reference unit (optionally sheet-qualified)
# and is replaced with the placeholder. Sheet prefixes are stripped because
# canonical-A1 is sheet-agnostic by design (cross-sheet template replicas
# should share a canonical).
_SHEET_PREFIX = r"(?:(?:'[^']*'|[A-Za-z_][\w\.]*)!)"

# Structured / table refs:
#   Table1[Col]                   — simple column ref
#   Table1[[#Headers],[Col]]      — qualifier + column (nested brackets)
#   Table1[#Headers]              — full-qualifier ref
#   [@Col]                        — implicit row, current table
#
# The outer brackets can contain nested [...] groups, so the body is matched
# as a sequence of either non-bracket chars or balanced [...] inner groups.
_BRACKET_BODY = r"(?:[^\[\]]|\[[^\[\]]*\])*"
_STRUCTURED_REF = (
    r"[A-Za-z_][\w\.]*\[" + _BRACKET_BODY + r"\]"
    r"|\[@[^\]]+\]"
    r"|\[#[^\]]+\]"
)

# Cell range: A1:B10, $A$1:$B$10 — must come before single cell + whole-col/row
_CELL_RANGE = r"\$?[A-Za-z]{1,3}\$?\d{1,7}:\$?[A-Za-z]{1,3}\$?\d{1,7}"

# Whole-column range: A:A, $D:$D
_COL_RANGE = r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"

# Whole-row range: 1:1, $5:$10
_ROW_RANGE = r"\$?\d{1,7}:\$?\d{1,7}"

# Single cell: A1, $A$1
_CELL = r"\$?[A-Za-z]{1,3}\$?\d{1,7}"

# Combined pattern — alternations ordered longest-first so the regex engine
# picks the most specific match. Sheet prefix is optional.
_REF_PATTERN = re.compile(
    r"(?P<sp>" + _SHEET_PREFIX + r")?"
    r"(?:"
    r"(?P<sref>" + _STRUCTURED_REF + r")"
    r"|(?P<crange>" + _CELL_RANGE + r")"
    r"|(?P<colrange>" + _COL_RANGE + r")"
    r"|(?P<rowrange>" + _ROW_RANGE + r")"
    r"|(?P<cell>" + _CELL + r")"
    r")"
)

# String literals: "..." with "" as escape — must be skipped, not canonicalized.
_STRING_PATTERN = re.compile(r'"(?:[^"]|"")*"')


def _placeholder_for(match: re.Match) -> str:
    if match.group("sref") is not None:
        # Structured refs collapse to CELLRANGE (they denote a column slice).
        return "CELLRANGE"
    if match.group("crange") is not None:
        return "CELLRANGE"
    if match.group("colrange") is not None:
        return "COLRANGE"
    if match.group("rowrange") is not None:
        return "ROWRANGE"
    if match.group("cell") is not None:
        return "CELL"
    return match.group(0)


def _replace_refs_outside_strings(formula: str) -> str:
    """Walk the formula, leaving string-literal contents intact while
    canonicalizing all references in non-string regions."""
    result: list[str] = []
    i = 0
    n = len(formula)
    while i < n:
        ch = formula[i]
        if ch == '"':
            m = _STRING_PATTERN.match(formula, i)
            if m:
                # Strings are dropped from canonical form: their content has no
                # bearing on family-grouping (two formulas differing only by a
                # literal label should still cluster). Replace with a stable
                # placeholder.
                result.append("STR")
                i = m.end()
                continue
            # Lone unmatched quote: pass through and advance one char.
            result.append(ch)
            i += 1
            continue
        # Try a ref match at this position.
        m = _REF_PATTERN.match(formula, i)
        if m:
            result.append(_placeholder_for(m))
            i = m.end()
            continue
        result.append(ch)
        i += 1
    return "".join(result)


def compute_canonical_a1(formula_text: str) -> str | None:
    """
    Compute a canonical, position-invariant form of an A1 formula.

    Replaces cell references with CELL, cell ranges with CELLRANGE,
    whole-column ranges with COLRANGE, whole-row ranges with ROWRANGE,
    string literals with STR. Uppercases identifiers / function names,
    normalizes whitespace.

    Returns None when the canonical form contains no parentheses — the
    "meaningfulness gate" that excludes pure references / pure literals /
    simple operator chains from family grouping (see #310 §D for the FP
    class this closes).

    Args:
        formula_text: A1-notation formula, with or without leading '='.

    Returns:
        Canonical form string (e.g., '=SUM(CELLRANGE)/COUNT(CELLRANGE)'),
        or None if the formula fails the meaningfulness gate.
    """
    if not formula_text:
        return None

    # Preserve the leading '=' if present, canonicalize the body.
    has_eq = formula_text.startswith("=")
    body = formula_text[1:] if has_eq else formula_text

    # Replace references (outside string literals).
    canon = _replace_refs_outside_strings(body)

    # Uppercase: function names, identifiers, and our placeholders are all
    # ASCII-safe; numeric literals are unaffected by case.
    canon = canon.upper()

    # Normalize whitespace: collapse runs to a single space, strip ends.
    canon = re.sub(r"\s+", " ", canon).strip()

    if has_eq:
        canon = "=" + canon

    # Meaningfulness gate: must contain a function call (parenthesis) to
    # constitute a meaningful family. Pure refs, pure literals, and operator
    # chains (=CELL+CELL) produce too-weak clusters.
    if "(" not in canon:
        return None

    return canon
