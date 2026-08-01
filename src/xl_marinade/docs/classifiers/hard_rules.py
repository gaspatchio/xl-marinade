# ABOUTME: Hard classification rules that override heuristic decisions.
# ABOUTME: Enforces constraints like "formula cells on calc sheets ≠ Input".

import logging
from dataclasses import dataclass

from .sheet_context import SheetType, infer_sheet_type

logger = logging.getLogger(__name__)

# Confidence level for forced classifications
# Lower than high-confidence heuristics to indicate rule-based override
FORCED_CLASSIFICATION_CONFIDENCE = 0.85


@dataclass
class ClassificationRejection:
    """Records when a classification is rejected by hard rules."""

    original_class: str
    rejected_reason: str
    forced_class: str


def validate_classification(
    binding_id: str, sheet_name: str, has_formula: bool, proposed_class: str, confidence: float
) -> tuple[str, float, ClassificationRejection | None]:
    """Validate and potentially override a proposed classification.

    Applies hard rules that override heuristic classification decisions.
    Key rule: Formula cells on calculation sheets CANNOT be Input or Assumption.

    Hard rules are non-negotiable constraints based on actuarial semantics:
    - Calculation sheets contain calculated values (formulas)
    - Input/Assumption sheets contain source data (constants)
    - A formula cell on a calculation sheet must be a Calculation or Result

    Args:
        binding_id: ID of the binding being classified
        sheet_name: Sheet where binding resides
        has_formula: Whether the binding contains formulas
        proposed_class: Classification from heuristics
        confidence: Confidence score from heuristics

    Returns:
        Tuple of (final_class, final_confidence, rejection_info)
        If rejection_info is not None, the classification was overridden.

    Examples:
        >>> # Formula on calculation sheet misclassified as Policyholder Data
        >>> validate_classification(
        ...     "abc123", "Projection", True, "Policyholder Data", 0.90
        ... )
        ('Calculation', 0.85, ClassificationRejection(...))

        >>> # Constant on assumption sheet - no rejection
        >>> validate_classification(
        ...     "def456", "Mortality table", False, "Assumption", 0.95
        ... )
        ('Assumption', 0.95, None)
    """
    sheet_type = infer_sheet_type(sheet_name)

    # HARD RULE: Formula cells on calculation sheets cannot be Policyholder Data/Assumption
    if sheet_type == SheetType.CALCULATION and has_formula:
        if proposed_class in ("Policyholder Data", "Assumption"):
            rejection = ClassificationRejection(
                original_class=proposed_class,
                rejected_reason=(
                    f"Formula cell on calculation sheet '{sheet_name}' "
                    f"cannot be {proposed_class}. Calculation sheets contain "
                    f"formulas that transform inputs, not source data."
                ),
                forced_class="Calculation",
            )
            logger.warning(f"Classification rejected for {binding_id}: {rejection.rejected_reason}")
            return ("Calculation", FORCED_CLASSIFICATION_CONFIDENCE, rejection)

    # No rejection - return original
    return (proposed_class, confidence, None)
