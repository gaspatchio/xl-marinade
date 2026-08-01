# ABOUTME: HLOOKUP-specific resolution strategies for row index resolution
# ABOUTME: Handles HLOOKUP patterns with literal, cell reference, and expression row indices

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xl_marinade.core.resolution import ResolutionEngine, ResolutionResult
    from xl_marinade.core.resolution_strategies import ResolutionContext


class HLookupLiteralRowIndexStrategy:
    """
    Resolve HLOOKUP when row_index_num is a literal constant.

    Example: =HLOOKUP(A1, B1:D10, 2, FALSE) → resolves to row 2
    """

    def __init__(self, resolution_engine: ResolutionEngine) -> None:
        """Initialize with resolution engine."""
        self.engine = resolution_engine

    def can_handle(self, func_name: str) -> bool:
        """Check if this strategy handles the function."""
        return func_name.upper() == "HLOOKUP"

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Attempt to resolve HLOOKUP with literal row_index.

        Args:
            context: Resolution context with AST and workbook

        Returns:
            ResolutionResult if row_index is literal, None to delegate
        """

        ast = context.ast
        func_name = ast.get("name", "").upper()
        if func_name != "HLOOKUP":
            return None

        args = ast.get("args", [])
        if len(args) < 3:
            return None

        row_index_arg = args[2]

        # Check if row_index is a literal constant
        if row_index_arg.get("type") != "Const":
            return None  # Delegate to next strategy

        row_index_value = row_index_arg.get("value")
        if not isinstance(row_index_value, (int, float)):
            return None

        # Use existing resolve_hlookup_semantic with resolved row_index
        result = self.engine.resolve_hlookup_semantic(
            ast, current_sheet=context.current_sheet, cell_address=context.cell_address
        )

        return result


class HLookupCellRefRowIndexStrategy:
    """
    Resolve HLOOKUP when row_index_num is a cell reference.

    Example: =HLOOKUP(A1, B1:D10, E1, FALSE) where E1=3 → resolves to row 3
    """

    def __init__(self, resolution_engine: ResolutionEngine) -> None:
        """Initialize with resolution engine."""
        self.engine = resolution_engine

    def can_handle(self, func_name: str) -> bool:
        """Check if this strategy handles the function."""
        return func_name.upper() == "HLOOKUP"

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Attempt to resolve HLOOKUP with cell reference row_index.

        Args:
            context: Resolution context with AST and workbook

        Returns:
            ResolutionResult if row_index resolves, None to delegate
        """

        ast = context.ast
        func_name = ast.get("name", "").upper()
        if func_name != "HLOOKUP":
            return None

        args = ast.get("args", [])
        if len(args) < 3:
            return None

        row_index_arg = args[2]

        # Check if row_index is a cell reference
        if row_index_arg.get("type") != "Ref":
            return None  # Delegate to next strategy

        # Try to resolve the cell reference using _resolve_argument
        row_index_result = self.engine._resolve_argument(
            row_index_arg, current_sheet=context.current_sheet
        )

        if not row_index_result.success:
            return None  # Delegate if can't resolve

        # Use existing resolve_hlookup_semantic
        result = self.engine.resolve_hlookup_semantic(
            ast, current_sheet=context.current_sheet, cell_address=context.cell_address
        )

        return result


class HLookupExpressionRowIndexStrategy:
    """
    Resolve HLOOKUP when row_index_num is a simple expression.

    Example: =HLOOKUP(A1, B1:D10, 1+1, FALSE) → resolves to row 2
    """

    def __init__(self, resolution_engine: ResolutionEngine) -> None:
        """Initialize with resolution engine."""
        self.engine = resolution_engine

    def can_handle(self, func_name: str) -> bool:
        """Check if this strategy handles the function."""
        return func_name.upper() == "HLOOKUP"

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Attempt to resolve HLOOKUP with expression row_index.

        Args:
            context: Resolution context with AST and workbook

        Returns:
            ResolutionResult if row_index expression resolves, None to delegate
        """

        ast = context.ast
        func_name = ast.get("name", "").upper()
        if func_name != "HLOOKUP":
            return None

        args = ast.get("args", [])
        if len(args) < 3:
            return None

        row_index_arg = args[2]

        # Check if row_index is a binary expression
        if row_index_arg.get("type") != "Binary":
            return None  # Delegate to next strategy

        # Try to resolve the expression using _resolve_argument
        row_index_result = self.engine._resolve_argument(
            row_index_arg, current_sheet=context.current_sheet
        )

        if not row_index_result.success:
            return None  # Delegate if can't resolve

        # Use existing resolve_hlookup_semantic
        result = self.engine.resolve_hlookup_semantic(
            ast, current_sheet=context.current_sheet, cell_address=context.cell_address
        )

        return result
