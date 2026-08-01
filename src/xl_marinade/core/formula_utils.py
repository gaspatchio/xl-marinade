# ABOUTME: Utility functions for extracting and handling Excel formulas from openpyxl cells.
# ABOUTME: Handles special cases like ArrayFormula objects to ensure robust formula extraction.

from openpyxl.cell.cell import Cell
from openpyxl.worksheet.formula import ArrayFormula


def extract_formula_string(cell: Cell) -> str:
    """
    Extract formula string from openpyxl cell, handling ArrayFormula objects.

    Excel array formulas are represented as ArrayFormula objects in openpyxl,
    not as strings. This function handles both regular formulas and array formulas.

    Args:
        cell: openpyxl Cell object

    Returns:
        Formula string (empty if cell has no formula)

    Examples:
        Regular formula: cell.value = "=A1+B1" → returns "=A1+B1"
        Array formula:   cell.value = ArrayFormula(...) → returns "=INDEX(...)"
        Value-only:      cell.value = 123 → returns ""
        Empty cell:      cell.value = None → returns ""

    Note:
        For array formulas, the actual formula text is stored in the
        ArrayFormula.text attribute, not in the object's string representation.
    """
    if not cell.value:
        return ""

    # Handle ArrayFormula objects
    # Excel array formulas (entered with Ctrl+Shift+Enter or dynamic arrays)
    # are represented as ArrayFormula objects in openpyxl
    if isinstance(cell.value, ArrayFormula):
        # ArrayFormula.text contains the actual formula string
        if hasattr(cell.value, "text") and cell.value.text:
            return str(cell.value.text)
        # Fallback if .text is missing or None
        return ""

    # Handle regular formula strings
    # Regular formulas are stored directly as strings in cell.value
    if isinstance(cell.value, str):
        return cell.value

    # Not a formula (number, boolean, datetime, etc.)
    return ""
