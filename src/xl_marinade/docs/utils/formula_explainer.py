# ABOUTME: Generate natural language descriptions of Excel formulas for documentation
# ABOUTME: Uses AST pattern matching to create human-readable formula explanations

"""
Formula Explainer Module

Generates plain English descriptions of Excel formulas for documentation using
AST-based pattern matching. Provides deterministic, human-readable explanations
for common formula patterns (arithmetic, lookups, aggregations, conditionals).
"""

import logging
from functools import lru_cache

from xl_marinade.core.parser import (
    ASTNode,
    BinaryNode,
    ConstNode,
    FormulaParser,
    FunctionNode,
    RefNode,
)

logger = logging.getLogger(__name__)


# Operator descriptions for natural language
OP_DESCRIPTIONS = {
    "+": "adds",
    "-": "subtracts",
    "*": "multiplies",
    "/": "divides",
    "^": "raises to power",
    "&": "concatenates",
}


def _describe_arithmetic(op: str) -> str:
    """Get verb for arithmetic operator"""
    return OP_DESCRIPTIONS.get(op, f"applies {op}")


def _describe_operand(
    node: ASTNode, variable_map: dict[str, str], current_sheet: str | None = None
) -> str:
    """Get human-readable name for operand"""
    if isinstance(node, ConstNode):
        return str(node.value)
    elif isinstance(node, RefNode):
        # Try to get label from variable_map
        addr = node.to_string()
        label = variable_map.get(addr)

        # Try with sheet prefix if not found and no sheet in address
        if not label and current_sheet and "!" not in addr:
            label = variable_map.get(f"{current_sheet}!{addr}")

        # Try without sheet prefix if not found and sheet in address
        if not label and "!" in addr:
            label = variable_map.get(addr.split("!")[-1])

        return label if label else addr
    elif isinstance(node, BinaryNode):
        # Nested operation - describe recursively
        return f"({_describe_from_ast(node, variable_map, current_sheet)})"
    elif isinstance(node, FunctionNode):
        # Nested function
        return f"({_describe_function(node, variable_map, current_sheet)})"
    return "value"


def _describe_function(
    node: FunctionNode, variable_map: dict[str, str], current_sheet: str | None = None
) -> str:
    """Generate description for function call"""
    func_name = node.name.upper()

    if func_name == "VLOOKUP":
        # VLOOKUP(lookup_value, table_array, col_index, [range_lookup])
        if len(node.args) >= 3:
            lookup_val = _describe_operand(node.args[0], variable_map, current_sheet)
            table = _describe_operand(node.args[1], variable_map, current_sheet)
            col_idx = node.args[2].value if isinstance(node.args[2], ConstNode) else "column"
            return f"Looks up {lookup_val} in {table} and returns column {col_idx}"

    elif func_name == "INDEX":
        # INDEX(array, row_num, [col_num])
        if len(node.args) >= 2:
            array = _describe_operand(node.args[0], variable_map, current_sheet)
            return f"Returns value from {array}"

    elif func_name == "MATCH":
        # MATCH(lookup_value, lookup_array, [match_type])
        if len(node.args) >= 2:
            lookup_val = _describe_operand(node.args[0], variable_map, current_sheet)
            array = _describe_operand(node.args[1], variable_map, current_sheet)
            return f"Finds position of {lookup_val} in {array}"

    elif func_name in ("SUM", "AVERAGE", "COUNT", "MIN", "MAX"):
        # Aggregation functions
        if len(node.args) > 0:
            range_desc = _describe_operand(node.args[0], variable_map, current_sheet)
            verb = "calculates average of" if func_name == "AVERAGE" else func_name.lower() + "s"
            return f"{verb.capitalize()} {range_desc}"

    elif func_name == "IF":
        # IF(condition, value_if_true, value_if_false)
        if len(node.args) >= 3:
            true_val = _describe_operand(node.args[1], variable_map, current_sheet)
            false_val = _describe_operand(node.args[2], variable_map, current_sheet)
            return f"Returns {true_val} if condition met, otherwise {false_val}"

    elif func_name == "IFS" and len(node.args) >= 2:
        # IFS(condition1, value1, condition2, value2, ...)
        return "Returns value based on multiple conditions"

    # Generic fallback
    return f"Applies {func_name} function"


def _describe_from_ast(
    ast: ASTNode, variable_map: dict[str, str], current_sheet: str | None = None
) -> str:
    """Generate description from AST node"""
    if isinstance(ast, BinaryNode):
        left_desc = _describe_operand(ast.left, variable_map, current_sheet)
        right_desc = _describe_operand(ast.right, variable_map, current_sheet)
        verb = _describe_arithmetic(ast.operator)
        return f"{verb.capitalize()} {left_desc} and {right_desc}"
    elif isinstance(ast, FunctionNode):
        return _describe_function(ast, variable_map, current_sheet)
    elif isinstance(ast, ConstNode):
        return f"Value is constant {ast.value}"
    elif isinstance(ast, RefNode):
        addr = ast.to_string()
        label = variable_map.get(addr)
        if not label and current_sheet and "!" not in addr:
            label = variable_map.get(f"{current_sheet}!{addr}")
        return f"References {label if label else addr}"
    return "Calculation"  # Fallback


@lru_cache(maxsize=1000)
def _parse_formula_cached(formula: str) -> tuple[bool, ASTNode | None]:
    """
    Parse formula with caching for performance.

    Returns:
        Tuple of (success, ast) where success is False if parsing fails
    """
    try:
        parser = FormulaParser()
        ast = parser.parse(formula)
        return (True, ast)
    except (ValueError, AttributeError, TypeError) as e:
        logger.debug(f"Failed to parse formula {formula}: {e}")
        return (False, None)


def generate_explanation(
    formula: str, semantic_formula: str, variable_map: dict[str, str], sheet: str
) -> str:
    """
    Generate natural language description of Excel formula.

    Uses AST pattern matching with LRU caching for performance optimization.
    Common formulas (repeated across rows) are cached to avoid re-parsing.

    Args:
        formula: Original Excel formula (e.g., "=A1+B1")
        semantic_formula: Formula with labels (e.g., "=Revenue+Cost")
        variable_map: Maps cell addresses to variable labels
        sheet: Current sheet name for resolving references

    Returns:
        Human-readable description (e.g., "Adds Revenue and Cost")
        Falls back to semantic_formula if description generation fails

    Performance:
        - Uses LRU cache (1000 formulas) for repeated patterns
        - Target: <5ms per formula for typical actuarial models

    Examples:
        >>> generate_explanation("=A1+B1", "=Revenue+Cost",
        ...                      {"Sheet!A1": "Revenue", "Sheet!B1": "Cost"}, "Sheet")
        'Adds Revenue and Cost'

        >>> generate_explanation("=SUM(A1:A10)", "=SUM(Premiums)",
        ...                      {"Sheet!A1:A10": "Premiums"}, "Sheet")
        'Sums Premiums'
    """
    # Validation
    if not formula or formula == "None":
        return ""

    try:
        # Parse formula to AST (with caching)
        success, ast = _parse_formula_cached(formula)
        if not success or not ast:
            return semantic_formula  # Fallback

        # Generate description from AST
        description = _describe_from_ast(ast, variable_map, sheet)
        return description

    except (ValueError, AttributeError, TypeError, KeyError) as e:
        logger.warning(f"Failed to generate description for {formula}: {e}")
        return semantic_formula  # Fallback to semantic formula
    except Exception as e:
        # Catch-all for unexpected errors (should not happen in production)
        logger.error(f"Unexpected error generating description for {formula}: {e}", exc_info=True)
        return semantic_formula
