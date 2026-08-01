# ABOUTME: INDEX-specific resolution strategies for partial and full resolution
# ABOUTME: Handles INDEX-MATCH-MATCH patterns with static column/row searches

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from xl_marinade.core.resolution import ResolutionEngine, ResolutionResult
    from xl_marinade.core.resolution_strategies import ResolutionContext

from xl_marinade.core.ref_converter import col_num_to_letter, parse_cell_address

# ============================================================================
# Shared Helper Functions (DRY - used by all INDEX strategies)
# ============================================================================


def _extract_array_dependencies(array_arg: dict[str, Any]) -> tuple[str, list[str]]:
    """
    Extract dependencies from INDEX array argument.

    Handles all array argument types:
    - Simple Ref: "A1:A10" → ("A1:A10", ["A1:A10"])
    - Named range: "MyRange" → ("MyRange", ["MyRange"])
    - OFFSET expr: OFFSET(A1:A10, B1, 0) → ("", ["A1:A10", "B1"])
    - IF expr: IF(C1, A1:A10, B1:B10) → ("", ["C1", "A1:A10", "B1:B10"])
    - Function result: INDIRECT("A1:A10") → ("", ["A1:A10"]) if literal arg

    Returns:
        Tuple of (resolved_ref, lookup_drivers):
        - resolved_ref: Primary reference string for resolved_lookup_ref (may be empty)
        - lookup_drivers: List of ALL cell/range references in expression

    Architecture:
        Uses established ref extraction pattern from ref_extractor.py and pipeline.py
        Recursively walks AST to find all Ref nodes (cell refs and named ranges)
    """
    if not isinstance(array_arg, dict):
        return "", []

    # Simple case: Direct reference (cell/range or named range)
    if array_arg.get("type") == "Ref":
        ref = array_arg.get("ref", "")
        return ref, [ref] if ref else []

    # Complex case: Expression containing references
    # Extract ALL references recursively (matches _walk_ast pattern)
    refs = _extract_refs_from_node(array_arg)

    # resolved_ref is first ref (or empty if none)
    # lookup_drivers is all refs
    resolved_ref = refs[0] if refs else ""
    return resolved_ref, refs


def _extract_refs_from_node(node: dict[str, Any]) -> list[str]:
    """
    Recursively extract all Ref nodes from AST subtree.

    Matches established pattern from ref_extractor.py:_walk_ast()
    and pipeline.py:_extract_references_from_ast()

    Args:
        node: AST node dictionary

    Returns:
        List of reference strings (deterministic order)
    """
    refs: list[str] = []

    if not isinstance(node, dict):
        return refs

    node_type = node.get("type")

    # Direct reference (cell/range or named range)
    if node_type == "Ref":
        ref = node.get("ref", "")
        if ref:
            refs.append(ref)

    # Function - recurse through arguments
    elif node_type == "Function":
        for arg in node.get("args", []):
            refs.extend(_extract_refs_from_node(arg))

    # Unary operator - recurse through operand
    elif node_type == "Unary":
        operand = node.get("operand")
        if operand:
            refs.extend(_extract_refs_from_node(operand))

    # Binary operator - recurse through both sides
    elif node_type == "Binary":
        left = node.get("left")
        right = node.get("right")
        if left:
            refs.extend(_extract_refs_from_node(left))
        if right:
            refs.extend(_extract_refs_from_node(right))

    # Const nodes have no references

    return refs


# ============================================================================
# Shared Helper Functions for MATCH resolution (DRY - used by Column/Row strategies)
# ============================================================================


def _try_resolve_match(
    engine: ResolutionEngine, arg: dict[str, Any], context: ResolutionContext
) -> tuple[int | None, list[str]]:
    """
    Attempt to resolve MATCH argument to a position using ResolutionEngine.

    This replaces the old _is_static_string_match check by using the engine's
    resolve_match_semantic method, which handles:
    - Literal strings: MATCH("Label", ...)
    - Cell references: MATCH(A1, ...) where A1 contains a constant
    - Expressions: MATCH(CONCATENATE(B3,B4), ...) where B3/B4 are constants
    - Arithmetic: MATCH(1+2, ...) resolving to 3

    Args:
        engine: Resolution engine with MATCH resolution logic
        arg: AST node (should be MATCH function)
        context: Resolution context with workbook snapshot

    Returns:
        Tuple of (position, dependencies) if resolved, (None, []) otherwise
    """
    if arg.get("type") != "Function":
        return None, []
    if arg.get("name", "").upper() != "MATCH":
        return None, []

    # Use public resolve_match_semantic
    result = engine.resolve_match_semantic(
        arg, current_sheet=context.current_sheet, cell_address=context.cell_address
    )

    if result.status == "resolved" and result.resolved_value is not None:
        try:
            position = int(result.resolved_value)
            return position, result.lookup_drivers
        except (ValueError, TypeError):
            return None, []

    return None, []


def _is_dynamic_match(arg: dict[str, Any]) -> bool:
    """
    Check if arg is MATCH with cell reference lookup_value.

    Args:
        arg: AST node

    Returns:
        True if this is MATCH(cell_ref, ...)
    """
    if arg.get("type") != "Function":
        return False
    if arg.get("name", "").upper() != "MATCH":
        return False

    match_args = arg.get("args", [])
    if not match_args:
        return False

    lookup_value = match_args[0]
    return bool(lookup_value.get("type") == "Ref")


class IndexFullResolutionStrategy:
    """
    Resolve INDEX when both row_num and column_num are deterministic.

    Wraps existing resolve_index_semantic logic for full resolution case.
    Returns exact cell when both row and column can be statically determined.
    """

    def __init__(self, resolution_engine: ResolutionEngine) -> None:
        """
        Initialize with resolution engine.

        Args:
            resolution_engine: Engine with existing INDEX resolution logic
        """
        self.engine = resolution_engine

    def can_handle(self, func_name: str) -> bool:
        """Return True if this is an INDEX function."""
        return func_name.upper() == "INDEX"

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Attempt full resolution using existing logic.

        Args:
            context: Resolution context

        Returns:
            ResolutionResult with exact cell if fully resolved, None otherwise
        """
        # Delegate to existing logic, only return if fully resolved
        result = self.engine.resolve_index_semantic(
            context.ast, current_sheet=context.current_sheet, cell_address=context.cell_address
        )

        if result.status == "resolved":
            result.partial_info["resolution_level"] = "exact_cell"
            return result

        # Not fully resolved - delegate to next strategy
        return None


class IndexPartialColumnStrategy:
    """
    Resolve INDEX to column range when column MATCH uses static/constant expression
    but row MATCH uses cell reference (dynamic).

    This handles the INDEX-MATCH-MATCH pattern from UC-S3-07:
    - Column MATCH has constant expression (e.g., "Payment Frequency", CONCATENATE("A","B"))
    - Row MATCH has cell reference (e.g., Calculations!$B55)

    We resolve the constant column MATCH and return the entire column
    as the dependency range.
    """

    def __init__(self, resolution_engine: ResolutionEngine) -> None:
        """
        Initialize with resolution engine.

        Args:
            resolution_engine: Engine with MATCH resolution logic
        """
        self.engine = resolution_engine

    def can_handle(self, func_name: str) -> bool:
        """Return True if this is an INDEX function."""
        return func_name.upper() == "INDEX"

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Attempt partial column resolution.

        Args:
            context: Resolution context

        Returns:
            ResolutionResult with column range if partially resolved, None otherwise
        """
        # Import here to avoid circular dependency
        from xl_marinade.core.resolution import ResolutionResult

        ast = context.ast
        args = ast.get("args", [])

        if len(args) < 3:
            return None  # Need table, row, column args

        table_arg = args[0]
        row_arg = args[1]
        col_arg = args[2]

        # Try to resolve column MATCH using engine
        col_position, col_drivers = _try_resolve_match(self.engine, col_arg, context)
        if col_position is None:
            return None  # Can't resolve column - delegate to next strategy

        # Check if row arg is dynamic (otherwise FullResolution handles it)
        if not _is_dynamic_match(row_arg):
            return None

        # Build column range
        table_range = table_arg.get("ref", "")
        if not table_range:
            return None

        column_range = self._extract_column(table_range, col_position, context.current_sheet)

        row_driver = self._get_row_driver(row_arg)

        return ResolutionResult(
            status="partial_resolved",
            resolved_lookup_ref=column_range,
            lookup_drivers=[table_range, row_driver] + col_drivers,
            notes=f"INDEX column resolved to position {col_position}; row depends on {row_driver}",
            partial_info={
                "resolution_level": "column_range",
                "table_range": table_range,
                "column_position": col_position,
                "row_driver": row_driver,
            },
        )

    def _extract_column(self, table_range: str, col_position: int, current_sheet: str) -> str:
        """
        Extract column N from table range as A1 reference.

        Args:
            table_range: Table range string (e.g., "'Asset Register'!$B$12:$AL$91")
            col_position: 1-based column position within table
            current_sheet: Current sheet name for context

        Returns:
            Column range as A1 reference
        """
        parsed = parse_cell_address(table_range)
        if not parsed:
            return table_range  # Fallback to full range

        sheet_name = parsed.get("sheet", current_sheet)
        start_row = parsed.get("row", 0)
        start_col = parsed.get("col", 0)
        height = parsed.get("height", 1)

        # Type narrowing - these should always be appropriate types from parse_cell_address
        if not isinstance(sheet_name, str):
            return table_range  # Fallback
        if not isinstance(start_row, int) or not isinstance(start_col, int):
            return table_range  # Fallback
        if not isinstance(height, int):
            return table_range  # Fallback

        # Compute target column
        target_col = start_col + (col_position - 1)
        end_row = start_row + height - 1
        col_letter = col_num_to_letter(target_col)

        # Build column range with proper quoting
        if sheet_name:
            # Quote sheet name if it contains spaces
            if " " in sheet_name and not sheet_name.startswith("'"):
                sheet_name = f"'{sheet_name}'"
            return f"{sheet_name}!${col_letter}${start_row}:${col_letter}${end_row}"
        else:
            return f"${col_letter}${start_row}:${col_letter}${end_row}"

    def _get_row_driver(self, row_arg: dict[str, Any]) -> str:
        """
        Extract cell reference that drives row lookup.

        Args:
            row_arg: Row argument AST node

        Returns:
            Cell reference string that drives the row lookup
        """
        match_args = row_arg.get("args", [])
        if match_args and match_args[0].get("type") == "Ref":
            ref = match_args[0].get("ref", "unknown")
            return str(ref) if ref is not None else "unknown"
        return "unknown"


class IndexPartialRowStrategy:
    """
    Resolve INDEX to row range when row MATCH uses constant expression
    but column MATCH uses cell reference (dynamic).

    Mirror of IndexPartialColumnStrategy for the opposite case.
    """

    def __init__(self, resolution_engine: ResolutionEngine) -> None:
        """
        Initialize with resolution engine.

        Args:
            resolution_engine: Engine with MATCH resolution logic
        """
        self.engine = resolution_engine

    def can_handle(self, func_name: str) -> bool:
        """Return True if this is an INDEX function."""
        return func_name.upper() == "INDEX"

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Attempt partial row resolution.

        Args:
            context: Resolution context

        Returns:
            ResolutionResult with row range if partially resolved, None otherwise
        """
        # Import here to avoid circular dependency
        from xl_marinade.core.resolution import ResolutionResult

        ast = context.ast
        args = ast.get("args", [])

        if len(args) < 3:
            return None  # Need table, row, column args

        table_arg = args[0]
        row_arg = args[1]
        col_arg = args[2]

        # Try to resolve row MATCH using engine
        row_position, row_drivers = _try_resolve_match(self.engine, row_arg, context)
        if row_position is None:
            return None  # Can't resolve row - delegate to next strategy

        # Check if column arg is dynamic (otherwise FullResolution handles it)
        if not _is_dynamic_match(col_arg):
            return None

        # Build row range
        table_range = table_arg.get("ref", "")
        if not table_range:
            return None

        row_range = self._extract_row(table_range, row_position, context.current_sheet)

        col_driver = self._get_col_driver(col_arg)

        return ResolutionResult(
            status="partial_resolved",
            resolved_lookup_ref=row_range,
            lookup_drivers=[table_range, col_driver] + row_drivers,
            notes=f"INDEX row resolved to position {row_position}; column depends on {col_driver}",
            partial_info={
                "resolution_level": "row_range",
                "table_range": table_range,
                "row_position": row_position,
                "column_driver": col_driver,
            },
        )

    def _extract_row(self, table_range: str, row_position: int, current_sheet: str) -> str:
        """
        Extract row N from table range as A1 reference.

        Args:
            table_range: Table range string
            row_position: 1-based row position within table
            current_sheet: Current sheet name for context

        Returns:
            Row range as A1 reference
        """
        parsed = parse_cell_address(table_range)
        if not parsed:
            return table_range  # Fallback to full range

        sheet_name = parsed.get("sheet", current_sheet)
        start_row = parsed.get("row", 0)
        start_col = parsed.get("col", 0)
        width = parsed.get("width", 1)

        # Type narrowing - these should always be appropriate types from parse_cell_address
        if not isinstance(sheet_name, str):
            return table_range  # Fallback
        if not isinstance(start_row, int) or not isinstance(start_col, int):
            return table_range  # Fallback
        if not isinstance(width, int):
            return table_range  # Fallback

        # Compute target row
        target_row = start_row + (row_position - 1)
        end_col = start_col + width - 1

        start_col_letter = col_num_to_letter(start_col)
        end_col_letter = col_num_to_letter(end_col)

        # Build row range with proper quoting
        if sheet_name:
            # Quote sheet name if it contains spaces
            if " " in sheet_name and not sheet_name.startswith("'"):
                sheet_name = f"'{sheet_name}'"
            return f"{sheet_name}!${start_col_letter}${target_row}:${end_col_letter}${target_row}"
        else:
            return f"${start_col_letter}${target_row}:${end_col_letter}${target_row}"

    def _get_col_driver(self, col_arg: dict[str, Any]) -> str:
        """
        Extract cell reference that drives column lookup.

        Args:
            col_arg: Column argument AST node

        Returns:
            Cell reference string that drives the column lookup
        """
        match_args = col_arg.get("args", [])
        if match_args and match_args[0].get("type") == "Ref":
            ref = match_args[0].get("ref", "unknown")
            return str(ref) if ref is not None else "unknown"
        return "unknown"


class Index2ArgStrategy:
    """
    Resolve 2-arg INDEX patterns (no column_num argument).

    This is the MISSING strategy that handles 97% of INDEX failures in a large model.

    Handles ALL 2-arg patterns by leveraging existing resolution:
    - MATCH row_num: =INDEX(array, MATCH(...))
    - Cell ref row_num: =INDEX(array, B1)
    - Expression row_num: =INDEX(array, ROW()-5)
    - Literal row_num: =INDEX(array, 5)

    Resolution algorithm:
    1. Detect 2-arg INDEX: len(args) == 2
    2. Delegate to engine.resolve_index_semantic() which has complete 2-arg logic
    3. If result is "resolved" or "partial_resolved", return it
    4. If result is "unresolved", convert to "partial_resolved" with full array dependency
       (per ADR-041: no conservative_fallback allowed)

    IMPORTANT: This strategy returns a result for EVERY 2-arg INDEX,
    never returning None (which would trigger ConservativeFallback).
    """

    def __init__(self, resolution_engine: ResolutionEngine) -> None:
        """
        Initialize with resolution engine.

        Args:
            resolution_engine: Engine with existing INDEX resolution logic
        """
        self.engine = resolution_engine

    def can_handle(self, func_name: str) -> bool:
        """Return True if this is an INDEX function."""
        return func_name.upper() == "INDEX"

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Attempt 2-arg INDEX resolution.

        Args:
            context: Resolution context

        Returns:
            ResolutionResult for all 2-arg INDEX patterns,
            None for 3-arg (delegate to other strategies)
        """
        # Import here to avoid circular dependency
        from xl_marinade.core.resolution import ResolutionResult

        ast = context.ast
        args = ast.get("args", [])

        # Only handle 2-arg INDEX (array, row_num)
        if len(args) != 2:
            return None  # Let other strategies handle 3-arg

        # Use existing resolution logic
        result = self.engine.resolve_index_semantic(
            ast, current_sheet=context.current_sheet, cell_address=context.cell_address
        )

        # CRITICAL: Never return None for 2-arg INDEX
        # If unresolved, convert to partial_resolved per ADR-041
        if result.status == "unresolved":
            # Extract dependencies from array argument (handles Ref, named ranges, expressions)
            resolved_ref, drivers = _extract_array_dependencies(args[0])
            return ResolutionResult(
                status="partial_resolved",
                resolved_lookup_ref=resolved_ref,
                lookup_drivers=drivers,
                notes="2-arg INDEX row_num unresolvable; using full array dependencies",
                partial_info={"resolution_level": "full_array"},
            )

        return result
