# ABOUTME: VLOOKUP-specific resolution strategies for column index resolution
# ABOUTME: Handles VLOOKUP patterns with literal, cell reference, and expression column indices

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xl_marinade.core.resolution import ResolutionEngine, ResolutionResult
    from xl_marinade.core.resolution_strategies import ResolutionContext


class VLookupLiteralColIndexStrategy:
    """
    Resolve VLOOKUP when col_index_num is a literal constant.

    Example: =VLOOKUP(A1, B:D, 2, FALSE) → resolves to column C
    """

    def __init__(self, resolution_engine: ResolutionEngine) -> None:
        """Initialize with resolution engine."""
        self.engine = resolution_engine

    def can_handle(self, func_name: str) -> bool:
        """Check if this strategy handles the function."""
        return func_name.upper() == "VLOOKUP"

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Attempt to resolve VLOOKUP with literal col_index.

        Args:
            context: Resolution context with AST and workbook

        Returns:
            ResolutionResult if col_index is literal, None to delegate
        """

        ast = context.ast
        func_name = ast.get("name", "").upper()
        if func_name != "VLOOKUP":
            return None

        args = ast.get("args", [])
        if len(args) < 3:
            return None

        col_index_arg = args[2]

        # Check if col_index is a literal constant
        if col_index_arg.get("type") != "Const":
            return None  # Delegate to next strategy

        col_index_value = col_index_arg.get("value")
        if not isinstance(col_index_value, (int, float)):
            return None

        # Use existing resolve_vlookup_semantic with resolved col_index
        result = self.engine.resolve_vlookup_semantic(
            ast, current_sheet=context.current_sheet, cell_address=context.cell_address
        )

        return result


class VLookupCellRefColIndexStrategy:
    """
    Resolve VLOOKUP when col_index_num is a cell reference.

    Example: =VLOOKUP(A1, B:D, E1, FALSE) where E1=2 → resolves to column C
    """

    def __init__(self, resolution_engine: ResolutionEngine) -> None:
        """Initialize with resolution engine."""
        self.engine = resolution_engine

    def can_handle(self, func_name: str) -> bool:
        """Check if this strategy handles the function."""
        return func_name.upper() == "VLOOKUP"

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Attempt to resolve VLOOKUP with cell reference col_index.

        Args:
            context: Resolution context with AST and workbook

        Returns:
            ResolutionResult if col_index resolves, None to delegate
        """

        ast = context.ast
        func_name = ast.get("name", "").upper()
        if func_name != "VLOOKUP":
            return None

        args = ast.get("args", [])
        if len(args) < 3:
            return None

        col_index_arg = args[2]

        # Check if col_index is a cell reference
        if col_index_arg.get("type") != "Ref":
            return None  # Delegate to next strategy

        # Try to resolve the cell reference using _resolve_argument
        col_index_result = self.engine._resolve_argument(
            col_index_arg, current_sheet=context.current_sheet
        )

        if not col_index_result.success:
            return None  # Delegate if can't resolve

        # Use existing resolve_vlookup_semantic
        result = self.engine.resolve_vlookup_semantic(
            ast, current_sheet=context.current_sheet, cell_address=context.cell_address
        )

        return result


class VLookupExpressionColIndexStrategy:
    """
    Resolve VLOOKUP when col_index_num is a simple expression.

    Example: =VLOOKUP(A1, B:D, 1+1, FALSE) → resolves to column C
    """

    def __init__(self, resolution_engine: ResolutionEngine) -> None:
        """Initialize with resolution engine."""
        self.engine = resolution_engine

    def can_handle(self, func_name: str) -> bool:
        """Check if this strategy handles the function."""
        return func_name.upper() == "VLOOKUP"

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Attempt to resolve VLOOKUP with expression col_index.

        Args:
            context: Resolution context with AST and workbook

        Returns:
            ResolutionResult if col_index expression resolves, None to delegate
        """

        ast = context.ast
        func_name = ast.get("name", "").upper()
        if func_name != "VLOOKUP":
            return None

        args = ast.get("args", [])
        if len(args) < 3:
            return None

        col_index_arg = args[2]

        # Check if col_index is a binary expression
        if col_index_arg.get("type") != "Binary":
            return None  # Delegate to next strategy

        # Try to resolve the expression using _resolve_argument
        col_index_result = self.engine._resolve_argument(
            col_index_arg, current_sheet=context.current_sheet
        )

        if not col_index_result.success:
            return None  # Delegate if can't resolve

        # Use existing resolve_vlookup_semantic
        result = self.engine.resolve_vlookup_semantic(
            ast, current_sheet=context.current_sheet, cell_address=context.cell_address
        )

        return result
