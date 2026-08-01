# ABOUTME: Resolution strategy framework for pluggable lookup resolution
# ABOUTME: Implements strategy chain pattern per design doc resolution_strategy_chain.md

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Protocol

from openpyxl import Workbook

if TYPE_CHECKING:
    from xl_marinade.core.manual_resolution import ManualResolutionProvider
    from xl_marinade.core.resolution import ResolutionEngine, ResolutionResult

logger = logging.getLogger(__name__)


@dataclass
class ResolutionContext:
    """
    Context provided to resolution strategies.

    Contains all information needed for a strategy to attempt resolution,
    including the parsed AST, value source (workbook or LazyValueFetcher),
    cell location, and manual overrides.
    """

    ast: dict[str, Any]
    workbook: Workbook | Any  # Workbook or ValueSource (LazyValueFetcher)
    cell_address: str
    current_sheet: str
    manual_provider: ManualResolutionProvider | None


class ResolutionStrategy(Protocol):
    """
    Interface for pluggable resolution strategies.

    Strategies are chained in priority order. Each strategy can either:
    - Return a ResolutionResult if it successfully resolves (even partially)
    - Return None to delegate to the next strategy in the chain
    """

    def can_handle(self, func_name: str) -> bool:
        """
        Return True if this strategy handles this function type.

        Args:
            func_name: Function name (e.g., "INDEX", "VLOOKUP")

        Returns:
            True if this strategy should attempt resolution for this function
        """
        ...

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Attempt resolution. Return None to delegate to next strategy.

        Args:
            context: All information needed for resolution

        Returns:
            ResolutionResult if this strategy can resolve (even partially)
            None to delegate to the next strategy in chain
        """
        ...


class ResolutionChain:
    """
    Chain of resolution strategies tried in order.

    Iterates through strategies in priority order until one returns a non-None result.
    Order is deterministic (ADR-000) - same input always produces same output.
    """

    def __init__(self, strategies: list[ResolutionStrategy]) -> None:
        """
        Initialize chain with ordered list of strategies.

        Args:
            strategies: List of strategies in priority order (highest first)
        """
        self.strategies = strategies

    def resolve(self, func_name: str, context: ResolutionContext) -> ResolutionResult:
        """
        Resolve using first successful strategy.

        Args:
            func_name: Function name (e.g., "INDEX", "VLOOKUP")
            context: Resolution context with all needed information

        Returns:
            ResolutionResult from first successful strategy, or unresolved status
        """
        # Import here to avoid circular dependency
        from xl_marinade.core.resolution import ResolutionResult

        for strategy in self.strategies:
            if strategy.can_handle(func_name):
                result = strategy.try_resolve(context)
                if result is not None:
                    logger.debug(
                        f"[{context.cell_address}] {func_name} resolved by "
                        f"{strategy.__class__.__name__}: {result.status}"
                    )
                    return result

        logger.warning(f"[{context.cell_address}] {func_name} unresolved by all strategies")
        return ResolutionResult(status="unresolved", notes=f"No strategy succeeded for {func_name}")


class ManualResolutionStrategy:
    """
    Check manual_resolutions.json for user-provided overrides.

    Priority: Highest (always checked first)
    Applies to: All lookup functions
    """

    SUPPORTED_FUNCTIONS = {"INDEX", "VLOOKUP", "HLOOKUP", "XLOOKUP", "MATCH", "CHOOSE", "ADDRESS"}

    def __init__(self, manual_provider: ManualResolutionProvider | None) -> None:
        """
        Initialize with manual resolution provider.

        Args:
            manual_provider: Provider for manual overrides, or None if not available
        """
        self.manual_provider = manual_provider

    def can_handle(self, func_name: str) -> bool:
        """Return True if this is a lookup function that supports manual resolution."""
        return func_name.upper() in self.SUPPORTED_FUNCTIONS

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Check for manual override and return it if found.

        Args:
            context: Resolution context

        Returns:
            ResolutionResult with manual override if found, None otherwise
        """
        # Import here to avoid circular dependency
        from xl_marinade.core.resolution import ResolutionResult

        if not self.manual_provider:
            return None

        manual = self.manual_provider.get_resolution(context.cell_address)
        if manual:
            return ResolutionResult(
                status="resolved",
                resolved_lookup_ref=manual.get("resolved_ref", ""),
                resolution_source="manual",
                notes=manual.get("reason", "Manual override"),
                partial_info={"resolution_level": "manual"},
            )

        return None


class ConservativeFallbackStrategy:
    """
    Last resort: depend on entire lookup table.

    This is mathematically correct (the formula DOES depend on the table)
    but imprecise (we don't know which specific cell).

    Priority: Low (near end of chain)
    Applies to: All table lookup functions EXCEPT INDEX (per ADR-041)
    """

    SUPPORTED_FUNCTIONS = {"VLOOKUP", "HLOOKUP", "XLOOKUP"}

    def can_handle(self, func_name: str) -> bool:
        """Return True if this is a table lookup function (EXCEPT INDEX)."""
        return func_name.upper() in self.SUPPORTED_FUNCTIONS

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult | None:
        """
        Extract table argument and use entire table as dependency.

        Args:
            context: Resolution context

        Returns:
            ResolutionResult with full table as dependency, or None if no table found
        """
        # Import here to avoid circular dependency
        from xl_marinade.core.resolution import ResolutionResult

        table_range = self._extract_table_arg(context.ast)
        if not table_range:
            return None

        # Ensure table_range is a string
        if not isinstance(table_range, str):
            return None

        return ResolutionResult(
            status="conservative_fallback",
            resolved_lookup_ref=table_range,
            lookup_drivers=[table_range],
            notes="Arguments could not be resolved; depending on full table",
            partial_info={"resolution_level": "full_table", "reason": "dynamic_arguments"},
        )

    def _extract_table_arg(self, ast: dict[str, Any]) -> str | None:
        """
        Extract table/array argument from lookup function AST.

        For VLOOKUP/HLOOKUP/XLOOKUP: second argument (index 1) is the table
        For INDEX: first argument (index 0) is the array

        Args:
            ast: Parsed AST of lookup function

        Returns:
            Table range string if found, None otherwise
        """
        func_name = ast.get("name", "").upper()
        args = ast.get("args", [])
        if not args:
            return None

        # Determine which argument is the table
        if func_name in ("VLOOKUP", "HLOOKUP", "XLOOKUP"):
            # Second argument is the table
            if len(args) < 2:
                return None
            table_arg = args[1]
        else:
            # INDEX and others use first argument
            table_arg = args[0]

        if table_arg.get("type") == "Ref":
            ref = table_arg.get("ref")
            return str(ref) if ref is not None else None

        return None


class IndexArrayFallbackStrategy:
    """
    Final fallback for INDEX: depend on entire array with partial_resolved status.

    CRITICAL: Returns partial_resolved, NOT conservative_fallback.
    This is mathematically correct - formula depends on array, we just
    don't know which specific cell.

    Per ADR-041: conservative_fallback is NOT acceptable for INDEX.
    All INDEX calls must return "resolved" or "partial_resolved".
    """

    def can_handle(self, func_name: str) -> bool:
        """Return True for INDEX function only."""
        return func_name.upper() == "INDEX"

    def try_resolve(self, context: ResolutionContext) -> ResolutionResult:
        """
        Return partial_resolved with array dependency.

        Args:
            context: Resolution context

        Returns:
            ResolutionResult with status="partial_resolved"
        """
        # Import here to avoid circular dependency
        from xl_marinade.core.resolution import ResolutionResult
        from xl_marinade.core.strategies.index_strategies import _extract_array_dependencies

        ast = context.ast
        args = ast.get("args", [])

        # Extract dependencies from array argument (handles Ref, named ranges, expressions)
        if args:
            resolved_ref, drivers = _extract_array_dependencies(args[0])
        else:
            resolved_ref, drivers = "", []

        return ResolutionResult(
            status="partial_resolved",
            resolved_lookup_ref=resolved_ref,
            lookup_drivers=drivers,
            notes="INDEX fallback - using full array dependencies (row/column unresolvable)",
            partial_info={"resolution_level": "full_array"},
        )


def create_index_resolution_chain(
    resolution_engine: ResolutionEngine, manual_provider: ManualResolutionProvider | None = None
) -> ResolutionChain:
    """
    Create resolution chain for INDEX function.

    Strategy order (deterministic per ADR-000):
    1. ManualResolutionStrategy - user overrides (highest priority)
    2. IndexFullResolutionStrategy - both row and column deterministic
    3. Index2ArgStrategy - ALL 2-arg INDEX patterns (NEW - handles 97% of failures)
    4. IndexPartialColumnStrategy - 3-arg with static column
    5. IndexPartialRowStrategy - 3-arg with static row
    6. IndexArrayFallbackStrategy - catch-all that returns partial_resolved
       (NEW - replaces ConservativeFallback per ADR-041)

    NOTE:
    - ConservativeFallbackStrategy REMOVED from INDEX chain per ADR-041
    - Order matters: more specific strategies first
    - Index2ArgStrategy is critical - it handles most actuarial patterns

    Args:
        resolution_engine: Engine with existing INDEX resolution logic
        manual_provider: Provider for manual overrides, or None

    Returns:
        ResolutionChain configured for INDEX lookups
    """
    # Import here to avoid circular dependency
    from xl_marinade.core.strategies.index_strategies import (
        Index2ArgStrategy,
        IndexFullResolutionStrategy,
        IndexPartialColumnStrategy,
        IndexPartialRowStrategy,
    )

    return ResolutionChain(
        [
            ManualResolutionStrategy(manual_provider),
            IndexFullResolutionStrategy(resolution_engine),
            Index2ArgStrategy(resolution_engine),  # NEW - handles 2-arg patterns
            IndexPartialColumnStrategy(resolution_engine),
            IndexPartialRowStrategy(resolution_engine),
            IndexArrayFallbackStrategy(),  # NEW - replaces ConservativeFallback
        ]
    )


def create_vlookup_resolution_chain(
    resolution_engine: ResolutionEngine, manual_provider: ManualResolutionProvider | None = None
) -> ResolutionChain:
    """
    Create resolution chain for VLOOKUP function.

    Strategy order (deterministic per ADR-000):
    1. ManualResolutionStrategy - user overrides (highest priority)
    2. VLookupLiteralColIndexStrategy - col_index is literal constant
    3. VLookupCellRefColIndexStrategy - col_index is cell reference
    4. VLookupExpressionColIndexStrategy - col_index is expression (1+1, etc.)
    5. ConservativeFallbackStrategy - use full table as dependency

    Args:
        resolution_engine: Engine with existing VLOOKUP resolution logic
        manual_provider: Provider for manual overrides, or None

    Returns:
        ResolutionChain configured for VLOOKUP lookups
    """
    # Import here to avoid circular dependency
    from xl_marinade.core.strategies.vlookup_strategies import (
        VLookupCellRefColIndexStrategy,
        VLookupExpressionColIndexStrategy,
        VLookupLiteralColIndexStrategy,
    )

    return ResolutionChain(
        [
            ManualResolutionStrategy(manual_provider),
            VLookupLiteralColIndexStrategy(resolution_engine),
            VLookupCellRefColIndexStrategy(resolution_engine),
            VLookupExpressionColIndexStrategy(resolution_engine),
            ConservativeFallbackStrategy(),
        ]
    )


def create_hlookup_resolution_chain(
    resolution_engine: ResolutionEngine, manual_provider: ManualResolutionProvider | None = None
) -> ResolutionChain:
    """
    Create resolution chain for HLOOKUP function.

    Strategy order (deterministic per ADR-000):
    1. ManualResolutionStrategy - user overrides (highest priority)
    2. HLookupLiteralRowIndexStrategy - row_index is literal constant
    3. HLookupCellRefRowIndexStrategy - row_index is cell reference
    4. HLookupExpressionRowIndexStrategy - row_index is expression (1+1, etc.)
    5. ConservativeFallbackStrategy - use full table as dependency

    Args:
        resolution_engine: Engine with existing HLOOKUP resolution logic
        manual_provider: Provider for manual overrides, or None

    Returns:
        ResolutionChain configured for HLOOKUP lookups
    """
    # Import here to avoid circular dependency
    from xl_marinade.core.strategies.hlookup_strategies import (
        HLookupCellRefRowIndexStrategy,
        HLookupExpressionRowIndexStrategy,
        HLookupLiteralRowIndexStrategy,
    )

    return ResolutionChain(
        [
            ManualResolutionStrategy(manual_provider),
            HLookupLiteralRowIndexStrategy(resolution_engine),
            HLookupCellRefRowIndexStrategy(resolution_engine),
            HLookupExpressionRowIndexStrategy(resolution_engine),
            ConservativeFallbackStrategy(),
        ]
    )
