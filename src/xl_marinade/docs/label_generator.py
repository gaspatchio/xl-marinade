# ABOUTME: Label generation utilities for bindings with poor or missing label candidates
# ABOUTME: Generates meaningful labels from formula patterns, dependencies, and actuarial context

import logging
import re

logger = logging.getLogger(__name__)


def generate_label_from_formula(
    formula: str,
    dependencies: list[str],
    ir_db_path: str,
    overlay_bindings: dict[str, tuple[str, float, str]],
) -> str | None:
    """
    Generate a meaningful label from formula pattern and dependencies.

    This function analyzes common Excel patterns (SUM, AVERAGE, etc.) and uses
    dependency labels to construct a meaningful variable name.

    Args:
        formula: The Excel formula (e.g., "=SUM(U7:U607)")
        dependencies: List of parent binding IDs this formula references
        ir_db_path: Path to IR database for looking up additional context
        overlay_bindings: Dict of binding_id -> (label, score, type) from overlay

    Returns:
        Generated label string, or None if generation not possible
    """
    if not formula or not formula.startswith("="):
        return None

    # Pattern 1: SUM of a range
    # If formula is =SUM(range) and range has a known label, generate "Total <Label>"
    sum_match = re.match(r"=SUM\(([A-Z]+\d+:[A-Z]+\d+)\)", formula, re.IGNORECASE)
    if sum_match and dependencies:
        # Get the label of the first dependency (the range being summed)
        dep_binding_id = dependencies[0]
        if dep_binding_id in overlay_bindings:
            dep_label, dep_score, _ = overlay_bindings[dep_binding_id]
            if dep_label and dep_label != "Unlabelled" and dep_score > 0.5:
                return f"Total {dep_label}"

    # Pattern 2: AVERAGE of a range
    avg_match = re.match(r"=AVERAGE\(([A-Z]+\d+:[A-Z]+\d+)\)", formula, re.IGNORECASE)
    if avg_match and dependencies:
        dep_binding_id = dependencies[0]
        if dep_binding_id in overlay_bindings:
            dep_label, dep_score, _ = overlay_bindings[dep_binding_id]
            if dep_label and dep_label != "Unlabelled" and dep_score > 0.5:
                return f"Average {dep_label}"

    # Pattern 3: NPV or PV calculations
    if re.match(r"=NPV\(", formula, re.IGNORECASE):
        if dependencies and len(dependencies) >= 2:
            # Usually NPV(rate, cashflows)
            # Use the cashflow label
            cf_binding_id = dependencies[1] if len(dependencies) > 1 else dependencies[0]
            if cf_binding_id in overlay_bindings:
                cf_label, cf_score, _ = overlay_bindings[cf_binding_id]
                if cf_label and cf_label != "Unlabelled" and cf_score > 0.5:
                    return f"NPV of {cf_label}"
        return "Net Present Value"

    # Pattern 4: Simple arithmetic with 2 operands (could expand this)
    # e.g. =A1*B1, =A1+B1
    simple_math = re.match(r"=([A-Z]+\d+)\s*([+\-*/])\s*([A-Z]+\d+)$", formula)
    if simple_math and len(dependencies) == 2:
        dep1_id, dep2_id = dependencies[0], dependencies[1]
        label1 = overlay_bindings.get(dep1_id, (None, 0, None))[0]
        label2 = overlay_bindings.get(dep2_id, (None, 0, None))[0]

        if label1 and label2 and label1 != "Unlabelled" and label2 != "Unlabelled":
            operator = simple_math.group(2)
            operators = {
                "*": f"{label1} × {label2}",
                "+": f"{label1} + {label2}",
                "-": f"{label1} − {label2}",
                "/": f"{label1} / {label2}",
            }
            if operator in operators:
                return operators[operator]

    # Could not generate a label
    return None


def should_generate_label(
    current_label: str, current_score: float, actuarial_class: str | None, candidate_type: str
) -> bool:
    """
    Determine if label generation should be attempted.

    Generation is triggered when:
    - Current label is generic (sheet name, address-based)
    - Score is below threshold for important variables
    - Variable is classified as Result or Calculation

    Args:
        current_label: Current label from candidate selection
        current_score: Confidence score of current label
        actuarial_class: Classification (Result, Calculation, Input, Assumption)
        candidate_type: Type of candidate that won (sheet_name, scan_above, etc.)

    Returns:
        True if label generation should be attempted
    """
    # Always generate for Results with generic labels
    if actuarial_class == "Result":
        # Sheet names are always generic for Results
        if candidate_type == "sheet_name":
            return True
        # Low confidence labels for Results
        if current_score < 0.95:
            return True

    # Generate for Calculations with very generic labels
    if actuarial_class == "Calculation":
        if candidate_type == "sheet_name" and current_score < 0.95:
            return True
        if current_score < 0.6:  # Very low confidence
            return True

    # Don't generate for Inputs/Assumptions (they should come from Excel)
    return False
