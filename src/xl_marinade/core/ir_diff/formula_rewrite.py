# ABOUTME: R1C1 formula tokenizer and coordinate transformer for canonicalization.
# ABOUTME: Rewrites absolute row/col refs through rho/kappa maps, preserves relative offsets.

from __future__ import annotations

import re

# Token types
_TOK_REF = "REF"  # R1C1 reference (possibly sheet-qualified)
_TOK_STRING = "STRING"  # String literal "..." or '...'
_TOK_OTHER = "OTHER"  # Everything else (operators, function names, numbers, etc.)

# Pattern for R1C1 references, possibly sheet-qualified
# Matches: R1C1, R[2]C[-1], R10C[-2], Sheet1!R1C1, 'Sheet Name'!R1C1, [Book.xlsx]Sheet!R1C1
_SHEET_PREFIX = r"""(?:
    (?:\[.*?\])? # Optional external workbook [Book.xlsx]
    (?:
        (?:'[^']*') |  # Sheet name in single quotes
        (?:\w+)        # Simple sheet name
    )
    !                  # Separator
)?"""

_ROW_PART = r"R(?:\[(-?\d+)\]|(\d+))"  # R[rel] or Rabs
_COL_PART = r"C(?:\[(-?\d+)\]|(\d+))"  # C[rel] or Cabs

_REF_PATTERN = re.compile(
    rf"({_SHEET_PREFIX})({_ROW_PART}{_COL_PART})",
    re.VERBOSE,
)

# Simpler tokenizer: split formula into refs, strings, and other tokens
_TOKENIZE_PATTERN = re.compile(
    r"""
    (?P<string>"[^"]*"|'[^']*')             # String literals
    | (?P<ref>
        (?:                                   # Optional sheet prefix
            (?:\[.*?\])?                      #   Optional external [Book.xlsx]
            (?:'[^']*'|\w+)                  #   Sheet name
            !                                 #   Separator
        )?
        R(?:\[-?\d+\]|\d+)                   # Row part
        C(?:\[-?\d+\]|\d+)                   # Col part
    )
    | (?P<other>[^"']+?)                     # Everything else (non-greedy to not eat refs)
    """,
    re.VERBOSE,
)


# More robust approach: scan character by character
def _tokenize_r1c1(formula: str) -> list[tuple[str, str]]:
    """Tokenize an R1C1 formula into (type, text) pairs.

    Returns a list of tokens where each is:
    - ("STRING", '"hello"')
    - ("REF", "Sheet1!R1C2" or "R[1]C[-2]")
    - ("OTHER", "+", "SUM(", etc.)
    """
    tokens = []
    i = 0
    n = len(formula)

    while i < n:
        ch = formula[i]

        # String literals
        if ch == '"':
            j = i + 1
            while j < n and formula[j] != '"':
                j += 1
            tokens.append((_TOK_STRING, formula[i : j + 1]))
            i = j + 1
            continue

        # Check for R1C1 reference (possibly with sheet prefix)
        ref_match = _try_match_ref(formula, i)
        if ref_match:
            ref_text, end = ref_match
            tokens.append((_TOK_REF, ref_text))
            i = end
            continue

        # Accumulate other characters
        j = i
        while j < n:
            if formula[j] == '"':
                break
            if _try_match_ref(formula, j):
                break
            j += 1
        if j > i:
            tokens.append((_TOK_OTHER, formula[i:j]))
        i = j

    return tokens


def _try_match_ref(formula: str, pos: int) -> tuple[str, int] | None:
    """Try to match an R1C1 reference at position pos.

    Returns (matched_text, end_position) or None.
    """
    n = len(formula)
    i = pos

    # Try to match optional sheet prefix
    prefix_end = i

    # External workbook prefix [Book.xlsx]
    if i < n and formula[i] == "[":
        j = formula.find("]", i)
        if j == -1:
            pass  # Not a valid external ref
        else:
            prefix_end = j + 1

    # Sheet name prefix (quoted or simple)
    sheet_start = prefix_end
    if prefix_end < n and formula[prefix_end] == "'":
        j = formula.find("'", prefix_end + 1)
        if j != -1 and j + 1 < n and formula[j + 1] == "!":
            prefix_end = j + 2  # Past the '!'
        else:
            prefix_end = sheet_start  # Reset — not a valid sheet prefix
    elif prefix_end < n and formula[prefix_end].isalpha():
        j = prefix_end
        while j < n and (formula[j].isalnum() or formula[j] == "_"):
            j += 1
        if j < n and formula[j] == "!":
            prefix_end = j + 1
        else:
            prefix_end = sheet_start  # Reset

    # Now try to match R[n]Cn or RnC[n] etc.
    r_start = prefix_end
    if r_start >= n or formula[r_start] != "R":
        return None

    # Must not be preceded by a letter (would be a function name like "ROUND")
    if prefix_end == pos and pos > 0:
        prev = formula[pos - 1]
        if prev.isalpha() or prev == "_":
            return None

    j = r_start + 1
    if j >= n:
        return None

    # Row part
    if formula[j] == "[":
        # Relative row R[n]
        k = formula.find("]", j)
        if k == -1:
            return None
        try:
            int(formula[j + 1 : k])
        except ValueError:
            return None
        j = k + 1
    elif formula[j].isdigit():
        # Absolute row Rn
        while j < n and formula[j].isdigit():
            j += 1
    else:
        return None

    # Column part: must start with C
    if j >= n or formula[j] != "C":
        return None
    j += 1
    if j >= n:
        return None

    if formula[j] == "[":
        # Relative col C[n]
        k = formula.find("]", j)
        if k == -1:
            return None
        try:
            int(formula[j + 1 : k])
        except ValueError:
            return None
        j = k + 1
    elif formula[j].isdigit():
        # Absolute col Cn
        while j < n and formula[j].isdigit():
            j += 1
    else:
        return None

    # Must not be followed by a letter/digit that would make this part of a longer token
    if j < n and (formula[j].isalpha() or formula[j] == "_"):
        return None

    return formula[pos:j], j


def _parse_ref(ref_text: str) -> dict:
    """Parse an R1C1 reference into components.

    Returns dict with:
    - prefix: sheet/external prefix (empty string if none)
    - row_abs: int or None (absolute row number)
    - row_rel: int or None (relative row offset)
    - col_abs: int or None (absolute col number)
    - col_rel: int or None (relative col offset)
    """
    # Split off prefix (everything before R)
    r_idx = ref_text.rfind("R")
    # Find the R that starts the row reference (not inside a sheet name)
    for idx in range(len(ref_text)):
        if ref_text[idx] == "R" and (idx == 0 or ref_text[idx - 1] == "!"):
            r_idx = idx
            break
        # Also check if this R starts after prefix
        if ref_text[idx] == "!" and idx + 1 < len(ref_text) and ref_text[idx + 1] == "R":
            r_idx = idx + 1
            break

    prefix = ref_text[:r_idx]
    rest = ref_text[r_idx:]  # e.g., "R10C[-2]" or "R[1]C5"

    row_abs = None
    row_rel = None
    col_abs = None
    col_rel = None

    # Parse row
    m = re.match(r"R\[(-?\d+)\]", rest)
    if m:
        row_rel = int(m.group(1))
        rest = rest[m.end() :]
    else:
        m = re.match(r"R(\d+)", rest)
        if m:
            row_abs = int(m.group(1))
            rest = rest[m.end() :]

    # Parse col
    m = re.match(r"C\[(-?\d+)\]", rest)
    if m:
        col_rel = int(m.group(1))
    else:
        m = re.match(r"C(\d+)", rest)
        if m:
            col_abs = int(m.group(1))

    return {
        "prefix": prefix,
        "row_abs": row_abs,
        "row_rel": row_rel,
        "col_abs": col_abs,
        "col_rel": col_rel,
    }


def _rebuild_ref(parsed: dict) -> str:
    """Rebuild an R1C1 reference from parsed components."""
    parts = [parsed["prefix"]]

    if parsed["row_rel"] is not None:
        parts.append(f"R[{parsed['row_rel']}]")
    elif parsed["row_abs"] is not None:
        parts.append(f"R{parsed['row_abs']}")

    if parsed["col_rel"] is not None:
        parts.append(f"C[{parsed['col_rel']}]")
    elif parsed["col_abs"] is not None:
        parts.append(f"C{parsed['col_abs']}")

    return "".join(parts)


def rewrite_formula_r1c1(
    formula: str,
    sheet_map: dict[str, str],
    row_map: dict[int, int | None],
    col_map: dict[int, int | None],
) -> str | None:
    """Rewrite an R1C1 formula through coordinate maps.

    - Absolute row/col references are transformed through rho/kappa.
    - Relative offsets are preserved as-is.
    - Sheet-qualified references have their sheet names rewritten via sheet_map.
    - String literals are preserved verbatim.
    - External workbook references are preserved verbatim.

    Args:
        formula: R1C1 formula string (e.g., "=R10C[-2]+R[1]C5").
        sheet_map: Mapping from A sheet names to B sheet names.
        row_map: Mapping from A row numbers to B row numbers.
        col_map: Mapping from A col numbers to B col numbers.

    Returns:
        Rewritten formula, or None if any absolute reference maps to a deleted row/col.
    """
    if not formula:
        return formula

    tokens = _tokenize_r1c1(formula)
    result_parts = []

    for tok_type, tok_text in tokens:
        if tok_type == _TOK_STRING:
            result_parts.append(tok_text)
        elif tok_type == _TOK_REF:
            rewritten = _rewrite_ref(tok_text, sheet_map, row_map, col_map)
            if rewritten is None:
                return None  # Reference to deleted row/col
            result_parts.append(rewritten)
        else:
            result_parts.append(tok_text)

    return "".join(result_parts)


def _rewrite_ref(
    ref_text: str,
    sheet_map: dict[str, str],
    row_map: dict[int, int | None],
    col_map: dict[int, int | None],
) -> str | None:
    """Rewrite a single R1C1 reference."""
    parsed = _parse_ref(ref_text)

    # Rewrite sheet prefix
    prefix = parsed["prefix"]
    if prefix:
        # Check for external ref — preserve verbatim
        if "[" in prefix:
            pass  # Keep as-is
        elif prefix.endswith("!"):
            sheet_name = prefix[:-1]
            # Remove quotes if present
            if sheet_name.startswith("'") and sheet_name.endswith("'"):
                sheet_name = sheet_name[1:-1]
            if sheet_name in sheet_map:
                new_name = sheet_map[sheet_name]
                # Re-quote if needed
                if " " in new_name or any(c in new_name for c in "!@#$%^&*()-+"):
                    parsed["prefix"] = f"'{new_name}'!"
                else:
                    parsed["prefix"] = f"{new_name}!"

    # Rewrite absolute row
    if parsed["row_abs"] is not None:
        if parsed["row_abs"] in row_map:
            new_row = row_map[parsed["row_abs"]]
            if new_row is None:
                return None  # Row was explicitly deleted
            parsed["row_abs"] = new_row
        # else: row not tracked (boundary/empty row like 1048576) — preserve as-is

    # Rewrite absolute col
    if parsed["col_abs"] is not None:
        if parsed["col_abs"] in col_map:
            new_col = col_map[parsed["col_abs"]]
            if new_col is None:
                return None  # Col was explicitly deleted
            parsed["col_abs"] = new_col
        # else: col not tracked — preserve as-is

    # Relative offsets preserved as-is
    return _rebuild_ref(parsed)
