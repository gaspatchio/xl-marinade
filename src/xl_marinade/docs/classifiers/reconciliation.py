# ABOUTME: Classify which variables require reconciliation testing in model re-implementation
# ABOUTME: Applies rule "Actuarial formulas → reconcile; Helper plumbing → skip"

import json
import logging
from pathlib import Path
from typing import Any

from xl_marinade.core.parser import ASTNode, BinaryNode, FunctionNode

logger = logging.getLogger(__name__)

# Actuarial function patterns (always reconcile)
ACTUARIAL_FUNCTIONS = {
    # Mathematical/Statistical
    "SUM",
    "AVERAGE",
    "COUNT",
    "COUNTA",
    "COUNTIF",
    "COUNTIFS",
    "MIN",
    "MAX",
    "STDEV",
    "STDEVP",
    "VAR",
    "VARP",
    "PERCENTILE",
    # Lookups
    "VLOOKUP",
    "HLOOKUP",
    "INDEX",
    "MATCH",
    "XLOOKUP",
    "LOOKUP",
    # Financial
    "NPV",
    "IRR",
    "PV",
    "FV",
    "PMT",
    "RATE",
    "NPER",
    "IPMT",
    "PPMT",
    # Conditional
    "IF",
    "IFS",
    "SWITCH",
    "CHOOSE",
    # Date calculations
    "DATE",
    "EDATE",
    "YEARFRAC",
    "DATEDIF",
    "EOMONTH",
    "NETWORKDAYS",
    # Other actuarial
    "SUMIF",
    "SUMIFS",
    "AVERAGEIF",
    "AVERAGEIFS",
    "SUMPRODUCT",
}

# Helper function patterns (skip reconciliation)
HELPER_FUNCTIONS = {
    # Formatting
    "TEXT",
    "CONCATENATE",
    "CONCAT",
    "TRIM",
    "UPPER",
    "LOWER",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "SUBSTITUTE",
    "REPLACE",
    # Structural
    "ROW",
    "COLUMN",
    "CELL",
    "INDIRECT",
    "OFFSET",
    "ADDRESS",
    # Type conversion
    "VALUE",
    "N",
    "T",
}

# Arithmetic operators that imply calculation (reconcile)
ARITHMETIC_OPERATORS = {"+", "-", "*", "/", "^"}


def classify_reconciliation_requirement(
    variable: dict[str, Any],
    formula_ast: ASTNode | None,
    role: str,
    overrides: dict[str, Any] | None = None,
) -> tuple[bool, str]:
    """
    Determine if variable requires reconciliation testing.

    Applies rule: "Actuarial formulas → reconcile; Helper plumbing → skip"

    Args:
        variable: Variable metadata dict (must include variable_id)
        formula_ast: Parsed formula AST (None for constants)
        role: Actuarial classification role (Result, Calculation, Input, etc.)
        overrides: Optional user override dict (variable_id → reconciliation decision)

    Returns:
        Tuple of (reconciliation_required: bool, rationale: str)
    """
    variable_id = variable.get("variable_id")

    # 1. Check user override first (always wins)
    if overrides and variable_id in overrides:
        override = overrides[variable_id]
        return override["reconciliation_required"], override["rationale"]

    # 2. Results always reconcile
    if role == "Result":
        return True, "Key model output requiring verification"

    # 3. Helpers never reconcile
    if role == "Helper":
        return False, "Structural or formatting helper, not a calculation"

    # 4. Constants (no formula) don't reconcile
    if formula_ast is None:
        return False, "Constant value, no calculation to verify"

    # 5. Analyze formula patterns
    try:
        functions = _extract_function_names(formula_ast)
        has_arithmetic = _has_arithmetic_operations(formula_ast)

        # Check arithmetic operations FIRST (even if wrapped in helper functions)
        if has_arithmetic:
            return True, "Contains mathematical calculations"

        # Check if any functions are actuarial (reconcile)
        actuarial_funcs = functions & ACTUARIAL_FUNCTIONS
        if actuarial_funcs:
            func_list = ", ".join(sorted(actuarial_funcs)[:3])
            return True, f"Uses actuarial formula pattern ({func_list})"

        # Check if all functions are helpers (skip)
        if functions and all(f in HELPER_FUNCTIONS for f in functions):
            func_list = ", ".join(sorted(functions)[:3])
            return False, f"Formatting/structural helper functions only ({func_list})"

    except Exception as e:
        logger.warning(f"Failed to analyze formula patterns for {variable_id}: {e}")
        # Fall through to default

    # 6. Default based on role
    if role == "Calculation":
        return True, "Intermediate calculation in model flow"
    elif role in ["Policyholder Data", "Assumption"]:
        return False, "Input data, no calculation to verify"
    else:
        return False, "No clear reconciliation requirement detected"


def _extract_function_names(ast: ASTNode) -> set[str]:
    """
    Extract all Excel function names from AST.

    Recursively walks AST and collects FunctionNode names.

    Returns:
        Set of uppercase function names (e.g., {'SUM', 'VLOOKUP'})
    """
    return set(_yield_function_names(ast))


def _yield_function_names(ast: ASTNode):
    """Generator yielding function names from AST."""
    if isinstance(ast, FunctionNode):
        yield ast.name.upper()
        # Recurse into arguments
        for arg in ast.args:
            yield from _yield_function_names(arg)
    elif isinstance(ast, BinaryNode):
        # Recurse into left and right
        yield from _yield_function_names(ast.left)
        yield from _yield_function_names(ast.right)


def _has_arithmetic_operations(ast: ASTNode) -> bool:
    """
    Check if AST contains arithmetic operators.

    Looks for +, -, *, /, ^ operations in BinaryNode.

    Returns:
        True if arithmetic operations found
    """
    if isinstance(ast, BinaryNode):
        if ast.operator in ARITHMETIC_OPERATORS:
            return True
        # Recurse into children
        return _has_arithmetic_operations(ast.left) or _has_arithmetic_operations(ast.right)

    if isinstance(ast, FunctionNode):
        # Recurse into function arguments
        return any(_has_arithmetic_operations(arg) for arg in ast.args)

    return False


def load_reconciliation_overrides(override_path: str) -> dict[str, dict[str, Any]]:
    """
    Load reconciliation overrides from JSON file.

    Expected format:
    {
      "overrides": [
        {
          "variable_id": "var_005",
          "reconciliation_required": false,
          "rationale": "Row counter, not a calculation"
        }
      ]
    }

    Returns:
        Dict mapping variable_id to override (reconciliation_required + rationale)
    """
    if not Path(override_path).exists():
        logger.info(f"No reconciliation overrides file found at {override_path}")
        return {}

    try:
        with open(override_path) as f:
            data = json.load(f)

        overrides = {}
        for override in data.get("overrides", []):
            var_id = override.get("variable_id")
            if var_id:
                overrides[var_id] = {
                    "reconciliation_required": override.get("reconciliation_required", False),
                    "rationale": override.get("rationale", "User override"),
                }

        logger.info(f"Loaded {len(overrides)} reconciliation overrides")
        return overrides

    except Exception as e:
        logger.error(f"Failed to load reconciliation overrides from {override_path}: {e}")
        return {}
