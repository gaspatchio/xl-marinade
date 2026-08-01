# ABOUTME: Sheet-type inference for classification context.
# ABOUTME: Detects if a sheet is for calculations, assumptions, or inputs.

from enum import Enum


class SheetType(Enum):
    """Semantic type of an Excel sheet based on its naming pattern."""

    CALCULATION = "calculation"
    ASSUMPTION = "assumption"
    POLICYHOLDER = "policyholder"
    UNKNOWN = "unknown"


# Pattern lists for sheet type inference
CALCULATION_PATTERNS: list[str] = [
    "projection",
    "calculation",
    "engine",
    "model",
    "cashflow",
    "output",
]

ASSUMPTION_PATTERNS: list[str] = [
    "assumption",
    "parameter",
    "config",
    "mortality",
    "lapse",
    "rate",
    "table",
    "expense",
    "commission",
    "inflation",
    "interest",
    "tax",
    "premium",
]

POLICYHOLDER_PATTERNS: list[str] = ["input", "data", "policyholder", "policy"]


def infer_sheet_type(sheet_name: str | None) -> SheetType:
    """Infer the semantic type of a sheet based on its name.

    Uses pattern matching on common actuarial sheet naming conventions.
    Calculation sheets contain formulas that transform inputs/assumptions.
    Assumption sheets contain judgmental parameters about future behavior.
    Policyholder sheets contain policy-level attributes that vary by insured life.

    Args:
        sheet_name: Name of the Excel sheet (None or empty returns UNKNOWN)

    Returns:
        SheetType enum indicating the sheet's role

    Examples:
        >>> infer_sheet_type("Projection")
        SheetType.CALCULATION
        >>> infer_sheet_type("Mortality table")
        SheetType.ASSUMPTION
        >>> infer_sheet_type("Policyholder data")
        SheetType.POLICYHOLDER
        >>> infer_sheet_type("Sheet1")
        SheetType.UNKNOWN
        >>> infer_sheet_type(None)
        SheetType.UNKNOWN
    """
    if not sheet_name:
        return SheetType.UNKNOWN

    lower = sheet_name.lower()

    # Check calculation patterns first (most restrictive)
    if any(p in lower for p in CALCULATION_PATTERNS):
        return SheetType.CALCULATION

    # Then assumption patterns
    if any(p in lower for p in ASSUMPTION_PATTERNS):
        return SheetType.ASSUMPTION

    # Then policyholder patterns
    if any(p in lower for p in POLICYHOLDER_PATTERNS):
        return SheetType.POLICYHOLDER

    return SheetType.UNKNOWN
