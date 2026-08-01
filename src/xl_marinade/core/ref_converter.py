# ABOUTME: Utilities for converting between A1 and R1C1 cell reference notations
# ABOUTME: Handles absolute, relative, and mixed references deterministically

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xl_marinade.core.parser import FormulaParser


@dataclass
class CellReference:
    """
    Parsed cell reference with absolute/relative flags.

    Attributes:
        row: Row number (1-indexed)
        col: Column number (1-indexed, A=1, B=2, etc.)
        row_absolute: True if row is absolute ($A1)
        col_absolute: True if column is absolute (A$1)
        sheet: Optional sheet name
    """

    row: int
    col: int
    row_absolute: bool
    col_absolute: bool
    sheet: str | None = None


def col_letter_to_num(col_letter: str) -> int:
    """
    Convert column letter(s) to column number.

    Args:
        col_letter: Column letter (A, B, ..., Z, AA, AB, ...)

    Returns:
        Column number (1-indexed, A=1, B=2, Z=26, AA=27, etc.)

    Example:
        >>> col_letter_to_num("A")
        1
        >>> col_letter_to_num("Z")
        26
        >>> col_letter_to_num("AA")
        27
    """
    col_letter = col_letter.upper()
    num = 0
    for char in col_letter:
        num = num * 26 + (ord(char) - ord("A") + 1)
    return num


# OPTIMIZATION FIX 3: Cache for common columns (A-ZZ = columns 1-702)
# This covers 99%+ of real-world Excel usage and eliminates string building overhead
# for the most frequently accessed columns during range expansion.
_COL_NUM_TO_LETTER_CACHE: dict[int, str] = {}


def _build_col_cache() -> None:
    """Build cache for columns 1-702 (A-ZZ) on module load."""
    for i in range(1, 703):
        letters = []
        num = i
        while num > 0:
            num -= 1
            letters.append(chr(num % 26 + ord("A")))
            num //= 26
        _COL_NUM_TO_LETTER_CACHE[i] = "".join(reversed(letters))


# Build cache on module import (one-time cost)
_build_col_cache()


def col_num_to_letter(col_num: int) -> str:
    """
    Convert column number to column letter(s).

    PERFORMANCE OPTIMIZATION (Sprint 5 Phase A):
    Columns 1-702 (A-ZZ) are cached to avoid repeated string building.
    This covers 99%+ of real-world usage and provides significant speedup
    during range expansion.

    Args:
        col_num: Column number (1-indexed, A=1, B=2, Z=26, AA=27, etc.)

    Returns:
        Column letter(s)

    Example:
        >>> col_num_to_letter(1)
        'A'
        >>> col_num_to_letter(26)
        'Z'
        >>> col_num_to_letter(27)
        'AA'
    """
    # Check cache first (covers A-ZZ, which is 99%+ of real usage)
    if col_num in _COL_NUM_TO_LETTER_CACHE:
        return _COL_NUM_TO_LETTER_CACHE[col_num]

    # Fallback for columns beyond ZZ (rare)
    letters = []
    while col_num > 0:
        col_num -= 1
        letters.append(chr(col_num % 26 + ord("A")))
        col_num //= 26
    return "".join(reversed(letters))


def quote_sheet_name(sheet_name: str) -> str:
    """
    Quote sheet name if necessary for A1 references.

    Args:
        sheet_name: Raw sheet name

    Returns:
        Sheet name enclosed in single quotes if it contains spaces or special characters,
        with internal single quotes escaped. Returns original if no quoting needed.

    Example:
        >>> quote_sheet_name("Sheet1")
        'Sheet1'
        >>> quote_sheet_name("Interest Rate")
        "'Interest Rate'"
        >>> quote_sheet_name("O'Reilly")
        "'O''Reilly'"
    """
    if not sheet_name:
        return ""

    # Check if quoting is needed
    # Excel requires quoting for spaces and most special characters
    # Simple heuristic: Quote if not just alphanumeric/underscore/period
    if re.match(r"^[A-Za-z0-9_.]+$", sheet_name):
        return sheet_name

    # Escape single quotes
    escaped = sheet_name.replace("'", "''")
    return f"'{escaped}'"


def parse_a1_reference(ref: str) -> CellReference | None:
    """
    Parse A1-style cell reference.

    Handles spill references (e.g., "B2#") by stripping the # operator and parsing
    the base cell reference. The # operator means "the entire spilled range from this cell"
    but for parsing purposes we just need the cell coordinates.

    Args:
        ref: A1-style reference (e.g., "A1", "$A$1", "A$1", "$A1", "Sheet1!A1", "B2#")

    Returns:
        Parsed CellReference or None if invalid

    Example:
        >>> ref = parse_a1_reference("$A$1")
        >>> (ref.row, ref.col, ref.row_absolute, ref.col_absolute)
        (1, 1, True, True)
        >>> ref = parse_a1_reference("B2#")
        >>> (ref.row, ref.col)
        (2, 2)
    """
    # Handle sheet-qualified references
    sheet = None
    if "!" in ref:
        parts = ref.split("!", 1)
        sheet = parts[0].strip("'")
        ref = parts[1]

    # Strip spill operator (#) if present
    # The # means "entire spilled range" but we parse the base cell
    if ref.endswith("#"):
        ref = ref[:-1]

    # Match A1-style reference with optional $ for absolute
    # Pattern: optional $ + letters + optional $ + digits
    match = re.match(r"^(\$?)([A-Z]+)(\$?)(\d+)$", ref, re.IGNORECASE)
    if not match:
        return None

    col_absolute = bool(match.group(1))
    col_letter = match.group(2).upper()
    row_absolute = bool(match.group(3))
    row_num = int(match.group(4))

    col_num = col_letter_to_num(col_letter)

    return CellReference(
        row=row_num, col=col_num, row_absolute=row_absolute, col_absolute=col_absolute, sheet=sheet
    )


def a1_to_r1c1(a1_ref: str, base_row: int, base_col: int) -> str:
    """
    Convert A1 reference to R1C1 notation relative to a base cell.

    Args:
        a1_ref: A1-style reference (e.g., "A1", "$A$1", "Sheet1!B2")
        base_row: Row of the cell containing this reference (1-indexed)
        base_col: Column of the cell containing this reference (1-indexed)

    Returns:
        R1C1-style reference

    Raises:
        ValueError: If a1_ref is invalid

    Examples:
        >>> a1_to_r1c1("A1", 1, 1)  # Relative ref from same cell
        'R[0]C[0]'
        >>> a1_to_r1c1("$A$1", 2, 2)  # Absolute ref
        'R1C1'
        >>> a1_to_r1c1("A$1", 2, 2)  # Mixed: col relative, row absolute
        'R1C[0]'
        >>> a1_to_r1c1("$A1", 2, 2)  # Mixed: col absolute, row relative
        'R[0]C1'
    """
    parsed = parse_a1_reference(a1_ref)
    if parsed is None:
        raise ValueError(f"Invalid A1 reference: {a1_ref}")

    # Build R1C1 notation
    parts = []

    # Sheet prefix if present
    if parsed.sheet:
        parts.append(f"{quote_sheet_name(parsed.sheet)}!")

    # Row part
    if parsed.row_absolute:
        parts.append(f"R{parsed.row}")
    else:
        offset = parsed.row - base_row
        if offset == 0:
            parts.append("R")
        else:
            parts.append(f"R[{offset}]")

    # Column part
    if parsed.col_absolute:
        parts.append(f"C{parsed.col}")
    else:
        offset = parsed.col - base_col
        if offset == 0:
            parts.append("C")
        else:
            parts.append(f"C[{offset}]")

    return "".join(parts)


def r1c1_to_a1(r1c1_ref: str, base_row: int, base_col: int) -> str:
    """
    Convert R1C1 reference to A1 notation relative to a base cell.

    Args:
        r1c1_ref: R1C1-style reference (e.g., "R1C1", "R[1]C[2]")
        base_row: Row of the cell containing this reference (1-indexed)
        base_col: Column of the cell containing this reference (1-indexed)

    Returns:
        A1-style reference

    Raises:
        ValueError: If r1c1_ref is invalid

    Examples:
        >>> r1c1_to_a1("R[0]C[0]", 1, 1)  # Relative to same cell
        'A1'
        >>> r1c1_to_a1("R1C1", 2, 2)  # Absolute
        '$A$1'
        >>> r1c1_to_a1("R1C[0]", 2, 2)  # Mixed: row absolute, col relative
        'B$1'
    """
    # Handle sheet-qualified references
    sheet = None
    if "!" in r1c1_ref:
        parts = r1c1_ref.split("!", 1)
        sheet = parts[0].strip("'")
        r1c1_ref = parts[1]

    # Parse R1C1 notation
    # Pattern: R (absolute) or R[offset] + C (absolute) or C[offset]
    match = re.match(r"^R(\d+|(?:\[[-+]?\d+\])?)?C(\d+|(?:\[[-+]?\d+\])?)$", r1c1_ref)
    if not match:
        raise ValueError(f"Invalid R1C1 reference: {r1c1_ref}")

    row_part = match.group(1) or ""
    col_part = match.group(2) or ""

    # Parse row
    if not row_part:
        # R (no offset) - relative to base row
        row_num = base_row
        row_absolute = False
    elif row_part.startswith("["):
        # R[offset] - relative
        offset = int(row_part[1:-1])
        row_num = base_row + offset
        row_absolute = False
    else:
        # R<num> - absolute
        row_num = int(row_part)
        row_absolute = True

    # Parse column
    if not col_part:
        # C (no offset) - relative to base col
        col_num = base_col
        col_absolute = False
    elif col_part.startswith("["):
        # C[offset] - relative
        offset = int(col_part[1:-1])
        col_num = base_col + offset
        col_absolute = False
    else:
        # C<num> - absolute
        col_num = int(col_part)
        col_absolute = True

    # Build A1 notation
    parts = []

    # Sheet prefix if present
    if sheet:
        parts.append(f"{quote_sheet_name(sheet)}!")

    # Column
    if col_absolute:
        parts.append("$")
    parts.append(col_num_to_letter(col_num))

    # Row
    if row_absolute:
        parts.append("$")
    parts.append(str(row_num))

    return "".join(parts)


def normalize_a1_range(range_ref: str) -> str:
    """
    Normalize A1 range reference to canonical form.

    Args:
        range_ref: A1-style range (e.g., "A1:B10", "A1", "Sheet1!A1:B10")

    Returns:
        Normalized range string

    Example:
        >>> normalize_a1_range("B10:A1")
        'A1:B10'
    """
    # Handle sheet-qualified references
    sheet_prefix = ""
    if "!" in range_ref:
        parts = range_ref.split("!", 1)
        sheet_raw = parts[0].strip("'")
        sheet_prefix = quote_sheet_name(sheet_raw) + "!"
        range_ref = parts[1]

    # Split range if present
    if ":" in range_ref:
        start, end = range_ref.split(":", 1)
        start_ref = parse_a1_reference(start)
        end_ref = parse_a1_reference(end)

        if start_ref is None or end_ref is None:
            return sheet_prefix + range_ref  # Return as-is if invalid

        # Normalize order (top-left to bottom-right)
        min_row = min(start_ref.row, end_ref.row)
        max_row = max(start_ref.row, end_ref.row)
        min_col = min(start_ref.col, end_ref.col)
        max_col = max(start_ref.col, end_ref.col)

        # Build normalized range
        start_str = col_num_to_letter(min_col) + str(min_row)
        end_str = col_num_to_letter(max_col) + str(max_row)

        return f"{sheet_prefix}{start_str}:{end_str}"

    return sheet_prefix + range_ref


def format_cell_address(sheet: str, row: int, col: int, absolute: bool = False) -> str:
    """
    Format cell address from components.

    Args:
        sheet: Sheet name (empty string for no sheet qualification)
        row: Row number (1-indexed)
        col: Column number (1-indexed)
        absolute: Whether to use absolute references ($A$1)

    Returns:
        Formatted cell address

    Example:
        >>> format_cell_address("Sheet1", 5, 2)
        'Sheet1!B5'
        >>> format_cell_address("Interest Rate", 5, 2)
        "'Interest Rate'!B5"
        >>> format_cell_address("", 1, 1)
        'A1'
        >>> format_cell_address("Sheet1", 1, 1, absolute=True)
        'Sheet1!$A$1'
    """
    col_letter = col_num_to_letter(col)

    cell_part = f"${col_letter}${row}" if absolute else f"{col_letter}{row}"

    if sheet:
        quoted_sheet = quote_sheet_name(sheet)
        return f"{quoted_sheet}!{cell_part}"
    else:
        return cell_part


_r1c1_parser: FormulaParser | None = None


def convert_formula_to_r1c1(formula: str, base_row: int, base_col: int) -> str:
    """
    Convert formula from A1 notation to R1C1 notation using AST parsing.

    Uses AST-based conversion to correctly preserve string literals.
    Falls back to regex if parsing fails.

    Args:
        formula: Formula string (with or without leading =)
        base_row: Row of cell containing formula (1-indexed)
        base_col: Column of cell containing formula (1-indexed)

    Returns:
        Formula in R1C1 notation
    """
    global _r1c1_parser
    import logging

    from xl_marinade.core.parser import FormulaParser

    has_equals = formula.startswith("=")
    formula_body = formula[1:] if has_equals else formula

    try:
        if _r1c1_parser is None:
            _r1c1_parser = FormulaParser()
        ast = _r1c1_parser.parse(formula_body)
        result = ast.to_r1c1_string(base_row, base_col)
        return f"={result}" if has_equals else result
    except Exception as e:
        # Fallback to regex approach for robustness
        logging.debug(f"AST parsing failed for formula, using regex fallback: {e}")
        return _convert_formula_to_r1c1_regex(formula, base_row, base_col)


def _convert_formula_to_r1c1_regex(formula: str, base_row: int, base_col: int) -> str:
    """
    Convert formula from A1 notation to R1C1 notation using regex (fallback).

    This is the legacy implementation used as a fallback when AST parsing fails.
    It may incorrectly convert cell references inside string literals.

    Args:
        formula: Formula string (with or without leading =)
        base_row: Row of cell containing formula (1-indexed)
        base_col: Column of cell containing formula (1-indexed)

    Returns:
        Formula in R1C1 notation
    """
    # Strip leading = if present
    has_equals = formula.startswith("=")
    if has_equals:
        formula = formula[1:]

    # Pattern to match A1-style cell references (including sheet-qualified, ranges, and spill refs)
    # Matches: A1, $A$1, Sheet1!A1, A1:B2, A1#, etc.
    # This is a simplified pattern - full Excel formula parsing is complex
    pattern = r"(?:(?:'[^']*'|[A-Za-z_][\w\.]*)\!)?(?:\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?#?)"

    def replace_ref(match: re.Match) -> str:
        ref = match.group(0)
        try:
            # Handle ranges
            if ":" in ref:
                # For ranges, convert start and end separately
                sheet_prefix = ""
                range_part = ref
                if "!" in ref:
                    parts = ref.split("!", 1)
                    sheet_prefix = parts[0] + "!"
                    range_part = parts[1]

                start, end = range_part.split(":", 1)
                start_r1c1 = a1_to_r1c1(start, base_row, base_col)
                end_r1c1 = a1_to_r1c1(end, base_row, base_col)

                # Remove sheet prefix from second ref if present
                if sheet_prefix:
                    start_r1c1 = sheet_prefix + start_r1c1.split("!")[-1]
                    end_r1c1 = end_r1c1.split("!")[-1]

                return f"{start_r1c1}:{end_r1c1}"
            else:
                # Single cell reference
                return a1_to_r1c1(ref, base_row, base_col)
        except (ValueError, AttributeError):
            # If conversion fails, return original ref
            return ref

    # Replace all cell references in formula
    result = re.sub(pattern, replace_ref, formula, flags=re.IGNORECASE)

    # Add back leading = if it was there
    if has_equals:
        result = "=" + result

    return result


# Excel's maximum row count (2^20 for .xlsx format)
EXCEL_MAX_ROWS = 1048576


def parse_column_only_range(ref_part: str) -> tuple[int, int] | None:
    """
    Parse full-column range references like "$A:$O" or "A:DM".

    Args:
        ref_part: The range portion without sheet prefix (e.g., "$A:$O")

    Returns:
        Tuple of (start_col, end_col) or None if not a column-only range

    Example:
        >>> parse_column_only_range("$A:$O")
        (1, 15)
        >>> parse_column_only_range("A:DM")
        (1, 117)
        >>> parse_column_only_range("A1:O10")
        None  # Has row numbers, not column-only
    """
    import re

    if ":" not in ref_part:
        return None

    parts = ref_part.split(":", 1)

    # Pattern for column-only reference (optional $ + letters only, no digits)
    col_only_pattern = r"^(\$?)([A-Z]+)$"

    start_match = re.match(col_only_pattern, parts[0], re.IGNORECASE)
    end_match = re.match(col_only_pattern, parts[1], re.IGNORECASE)

    if start_match and end_match:
        start_col = col_letter_to_num(start_match.group(2).upper())
        end_col = col_letter_to_num(end_match.group(2).upper())
        return (start_col, end_col)

    return None


def parse_row_only_range(ref_part: str) -> tuple[int, int] | None:
    """
    Parse full-row range references like "$1:$50" or "1:100".

    Args:
        ref_part: The range portion without sheet prefix (e.g., "$1:$50")

    Returns:
        Tuple of (start_row, end_row) or None if not a row-only range

    Example:
        >>> parse_row_only_range("$1:$50")
        (1, 50)
        >>> parse_row_only_range("1:100")
        (1, 100)
        >>> parse_row_only_range("A1:A50")
        None  # Has column letters, not row-only
    """
    import re

    if ":" not in ref_part:
        return None

    parts = ref_part.split(":", 1)

    # Pattern for row-only reference (optional $ + digits only, no letters)
    row_only_pattern = r"^(\$?)(\d+)$"

    start_match = re.match(row_only_pattern, parts[0])
    end_match = re.match(row_only_pattern, parts[1])

    if start_match and end_match:
        start_row = int(start_match.group(2))
        end_row = int(end_match.group(2))
        return (start_row, end_row)

    return None


def parse_cell_address(cell: str) -> dict[str, str | int]:
    """
    Parse cell address into components for sorting/analysis.

    Args:
        cell: Cell address in format "Sheet!A1" or "A1" or "A1:B2"
              Also handles full-column refs like "Sheet!$A:$O"
              and full-row refs like "Sheet!$1:$50"

    Returns:
        Dictionary with keys: sheet (str), row (int), col (int), address (str)
        For ranges, also includes: height (int), width (int)
        For full-column ranges: row=1, height=EXCEL_MAX_ROWS
        For full-row ranges: col=1, width=16384 (Excel max columns)
        For unparseable addresses (external refs, malformed), returns row=0, col=0
        to ensure deterministic sorting (unparseable addresses sort before valid cells).

    Example:
        >>> parse_cell_address("Sheet1!B5")
        {'sheet': 'Sheet1', 'row': 5, 'col': 2, 'address': 'Sheet1!B5'}
        >>> parse_cell_address("A1")
        {'sheet': '', 'row': 1, 'col': 1, 'address': 'A1'}
        >>> parse_cell_address("A1:B2")
        {'sheet': '', 'row': 1, 'col': 1, 'height': 2, 'width': 2, 'address': 'A1:B2'}
        >>> parse_cell_address("[External.xlsx]Sheet1!A1")
        {'sheet': '[External.xlsx]Sheet1', 'row': 0, 'col': 0,
         'address': '[External.xlsx]Sheet1!A1'}

    Note:
        External references (e.g., '[Book.xlsx]Sheet!A1') and malformed addresses
        cannot be parsed by parse_a1_reference(), so they receive row=0, col=0
        for deterministic sorting. This ensures they sort before valid cells.
    """
    sheet = ""
    ref_part = cell

    # Extract sheet if present
    if "!" in cell:
        parts = cell.split("!", 1)
        sheet = parts[0].strip("'")
        ref_part = parts[1]

    # Check if it's a range
    if ":" in ref_part:
        # Try full-column range first (e.g., "$A:$O")
        col_range = parse_column_only_range(ref_part)
        if col_range:
            start_col, end_col = col_range
            return {
                "sheet": sheet,
                "row": 1,
                "col": start_col,
                "height": EXCEL_MAX_ROWS,
                "width": abs(end_col - start_col) + 1,
                "address": cell,
            }

        # Try full-row range (e.g., "$1:$50")
        row_range = parse_row_only_range(ref_part)
        if row_range:
            start_row, end_row = row_range
            # Excel max columns is 16384 (XFD)
            return {
                "sheet": sheet,
                "row": start_row,
                "col": 1,
                "height": abs(end_row - start_row) + 1,
                "width": 16384,
                "address": cell,
            }

        # Standard cell range (e.g., "A1:B10")
        parts = ref_part.split(":", 1)
        start_parsed = parse_a1_reference(parts[0])
        end_parsed = parse_a1_reference(parts[1])

        if start_parsed is None or end_parsed is None:
            # Return defaults for unparseable ranges
            return {"sheet": sheet, "row": 0, "col": 0, "address": cell}

        # Calculate dimensions
        height = abs(end_parsed.row - start_parsed.row) + 1
        width = abs(end_parsed.col - start_parsed.col) + 1

        return {
            "sheet": sheet,
            "row": start_parsed.row,
            "col": start_parsed.col,
            "height": height,
            "width": width,
            "address": cell,
        }

    # Parse single cell reference
    parsed = parse_a1_reference(ref_part)

    if parsed is None:
        # Return defaults for unparseable addresses (external refs, malformed)
        # row=0, col=0 ensures deterministic sorting (unparseable sort before valid)
        return {"sheet": sheet, "row": 0, "col": 0, "address": cell}

    return {"sheet": sheet, "row": parsed.row, "col": parsed.col, "address": cell}
