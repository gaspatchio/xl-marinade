# ABOUTME: Extract cell and range references from parsed formula AST
# ABOUTME: Identifies all static references deterministically for dependency graph construction

import re
from typing import TYPE_CHECKING, Any

from xl_marinade.core.ref_converter import col_letter_to_num, col_num_to_letter, parse_a1_reference

if TYPE_CHECKING:
    from xl_marinade.core.lazy_workbook import PopulatedCells


def extract_references_from_ast(ast: dict[str, Any]) -> list[str]:
    """
    Extract all cell/range references from an AST.

    Args:
        ast: Dictionary representation of formula AST

    Returns:
        Sorted list of unique reference strings (deterministic order)

    Example:
        >>> ast = {'type': 'Binary', 'operator': '+',
        ...        'left': {'type': 'Ref', 'ref': 'A1'},
        ...        'right': {'type': 'Ref', 'ref': 'B1'}}
        >>> extract_references_from_ast(ast)
        ['A1', 'B1']
    """
    refs: set[str] = set()
    _walk_ast(ast, refs)
    return sorted(refs)  # Sort for deterministic ordering


def _walk_ast(node: dict[str, Any], refs: set[str]) -> None:
    """
    Recursively walk AST and collect references.

    Args:
        node: AST node dictionary
        refs: Set to accumulate references (modified in place)
    """
    if not isinstance(node, dict):
        return

    node_type = node.get("type")

    if node_type == "Ref":
        # Direct reference node
        ref = node.get("ref")
        if ref:
            refs.add(ref)

    elif node_type == "Function":
        # Function node - walk all arguments
        args = node.get("args", [])
        for arg in args:
            _walk_ast(arg, refs)

    elif node_type == "Unary":
        # Unary node - walk operand
        operand = node.get("operand")
        if operand:
            _walk_ast(operand, refs)

    elif node_type == "Binary":
        # Binary node - walk both sides
        left = node.get("left")
        right = node.get("right")
        if left:
            _walk_ast(left, refs)
        if right:
            _walk_ast(right, refs)

    # Const nodes have no references to extract


def is_structured_reference(ref: str) -> bool:
    """
    Check if reference is a structured table reference.

    Args:
        ref: Reference string

    Returns:
        True if reference contains table notation (e.g., "Table1[Column]")

    Example:
        >>> is_structured_reference("Table1[Revenue]")
        True
        >>> is_structured_reference("A1")
        False
    """
    return "[" in ref and "]" in ref


def is_defined_name(ref: str) -> bool:
    """
    Check if reference might be a defined name.

    Note: This is a heuristic. True defined name resolution requires
    workbook metadata (handled in Story 2.5).

    Args:
        ref: Reference string

    Returns:
        True if reference doesn't match A1 or structured reference patterns

    Example:
        >>> is_defined_name("MyVariable")
        True
        >>> is_defined_name("A1")
        False
        >>> is_defined_name("Table1[Col]")
        False
    """
    # Not a structured ref
    if is_structured_reference(ref):
        return False

    # Sheet-qualified A1 pattern
    if "!" in ref:
        parts = ref.split("!", 1)
        ref = parts[1]

    # A1-style pattern: optional $ + letters + optional $ + digits
    if re.match(r"^(\$?)([A-Z]+)(\$?)(\d+)$", ref, re.IGNORECASE):
        return False

    # Range pattern
    if ":" in ref:
        parts = ref.split(":", 1)
        if all(re.match(r"^(\$?)([A-Z]+)(\$?)(\d+)$", p, re.IGNORECASE) for p in parts):
            return False
        # Full-column reference (e.g., C:C, $C:$C)
        if all(re.match(r"^(\$?)([A-Z]+)$", p, re.IGNORECASE) for p in parts):
            return False
        # Full-row reference (e.g., 5:5, $5:$5)
        if all(re.match(r"^(\$?)(\d+)$", p, re.IGNORECASE) for p in parts):
            return False

    # Otherwise, likely a defined name
    return True


def _expand_range_sparse(range_ref: str, populated_cells: "PopulatedCells") -> list[str]:
    """
    Sparse range expansion using precomputed populated cells.

    Instead of expanding range to ALL cells in bounding box, intersects range
    with populated cells and returns only cells that actually exist.

    Performance: O(populated_in_range * log N) vs O(range_area) for naive expansion.
    Example: A1:ZZ10000 (260k cells) with 100 populated → returns 100 cells.

    Preserves deterministic row-major ordering (A1, B1, C1, ..., A2, B2, ...).

    NOTE: Not cached (unlike legacy path) because populated_cells varies per sheet.
    The performance win comes from sparse intersection, not caching.

    Args:
        range_ref: Range reference (e.g., "A1:B2", "Sheet1!A1:B2")
        populated_cells: PopulatedCells object with precomputed cell sets

    Returns:
        List of populated cell references within range, in row-major order
    """
    # Handle sheet-qualified references
    sheet_prefix = ""
    if "!" in range_ref:
        parts = range_ref.split("!", 1)
        sheet_prefix = parts[0] + "!"
        range_ref = parts[1]

    # Single cell (not a range)
    if ":" not in range_ref:
        # In sparse mode, PopulatedCells only tracks formula cells (performance optimization).
        # For single-cell refs we therefore only return the cell if it is a formula cell.
        # Constant/value-only single-cell precedents are handled elsewhere (e.g., constant
        # binding detection) to avoid traversing tens of thousands of input cells.
        coord = range_ref.replace("$", "").rstrip("#")  # Normalize for set membership
        if coord in populated_cells.formula_cells:
            return [sheet_prefix + range_ref]
        return []

    # Parse range bounds
    start_str, end_str = range_ref.split(":", 1)
    start_ref = parse_a1_reference(start_str)
    end_ref = parse_a1_reference(end_str)

    if start_ref is None or end_ref is None:
        # Invalid range, return empty
        return []

    # Calculate range bounds
    min_row = min(start_ref.row, end_ref.row)
    max_row = max(start_ref.row, end_ref.row)
    min_col = min(start_ref.col, end_ref.col)
    max_col = max(start_ref.col, end_ref.col)

    # Intersect with populated cells using precomputed spatial index
    # PERFORMANCE FIX: Use binary search on precomputed spatial_index to find start position,
    # then iterate only cells within range bounds. This changes complexity from
    # O(all_populated_cells) to O(log N + cells_in_range).
    #
    # For a sheet with 123k populated cells and a range like C16:C38 (28 cells),
    # this reduces from 123k iterations to ~log(123k) + 28 ≈ 45 operations.
    import bisect

    result = []

    # Binary search strategy: Find first cell >= (min_row, min_col) using precomputed spatial_index
    # spatial_index is list of (coord, (row, col)) tuples, already sorted by (row, col)

    # Binary search for start position (first cell >= min_row)
    # Use bisect_left with key function to find first row >= min_row
    start_idx = bisect.bisect_left(
        populated_cells.spatial_index, (min_row, 0), key=lambda x: (x[1][0], 0)
    )

    # Iterate from start_idx until we exceed max_row
    for i in range(start_idx, len(populated_cells.spatial_index)):
        coord, (row_idx, col_idx) = populated_cells.spatial_index[i]

        # Stop if we've exceeded max_row (since list is sorted by row)
        if row_idx > max_row:
            break

        # Check if within range bounds
        if min_row <= row_idx <= max_row and min_col <= col_idx <= max_col:
            # Add absolute markers if original range had them
            cell = coord
            if start_ref.col_absolute or end_ref.col_absolute:
                # Preserve absolute column marker
                cell = "$" + cell if not cell.startswith("$") else cell

            result.append(sheet_prefix + cell)

    return result


def _expand_range_to_cells_cached(range_ref: str, max_cells: int = 10000) -> tuple[str, ...]:
    """
    Cached implementation of expand_range_to_cells (legacy path).

    Returns tuple for hashability (required by lru_cache).
    This is the performance-critical inner function that gets memoized.

    NOTE: This is the LEGACY path used when populated_cells not available.
    Prefer sparse expansion (_expand_range_sparse) when possible.
    """
    # Handle sheet-qualified references
    sheet_prefix = ""
    if "!" in range_ref:
        parts = range_ref.split("!", 1)
        sheet_prefix = parts[0] + "!"
        range_ref = parts[1]

    # Check if it's a range
    if ":" not in range_ref:
        return (sheet_prefix + range_ref,)  # Single cell as tuple

    # OPTIMIZATION FIX 2: Early cell count check before expensive parsing
    # Calculate cell count using lightweight parsing to avoid work for large ranges
    start_str, end_str = range_ref.split(":", 1)

    # Quick bounds calculation without full parsing
    try:
        # Simple regex to extract row/col without full parse_a1_reference overhead
        start_match = re.match(r"^(\$?)([A-Z]+)(\$?)(\d+)$", start_str, re.IGNORECASE)
        end_match = re.match(r"^(\$?)([A-Z]+)(\$?)(\d+)$", end_str, re.IGNORECASE)

        if start_match and end_match:
            start_col = col_letter_to_num(start_match.group(2))
            start_row = int(start_match.group(4))
            end_col = col_letter_to_num(end_match.group(2))
            end_row = int(end_match.group(4))

            num_rows = abs(end_row - start_row) + 1
            num_cols = abs(end_col - start_col) + 1
            total_cells = num_rows * num_cols

            # Skip expansion if range is too large
            if total_cells > max_cells:
                return ()  # Empty tuple - caller will skip this range
    except (ValueError, AttributeError):
        # Fall through to full parsing if quick check fails
        pass

    # Now do full parsing for ranges we'll actually expand
    start_ref = parse_a1_reference(start_str)
    end_ref = parse_a1_reference(end_str)

    if start_ref is None or end_ref is None:
        # Invalid range, return as-is as single item
        return (sheet_prefix + range_ref,)

    # Calculate range bounds
    min_row = min(start_ref.row, end_ref.row)
    max_row = max(start_ref.row, end_ref.row)
    min_col = min(start_ref.col, end_ref.col)
    max_col = max(start_ref.col, end_ref.col)

    # Expand to individual cells
    cells = []
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            col_letter = col_num_to_letter(col)
            cell = f"{col_letter}{row}"

            # Add absolute markers if original had them
            if start_ref.col_absolute or end_ref.col_absolute:
                cell = "$" + cell

            cells.append(sheet_prefix + cell)

    return tuple(cells)


def expand_range_to_cells(
    range_ref: str,
    max_cells: int = 10000,
    populated_cells: "PopulatedCells | None" = None,
    debug: bool = False,
) -> list[str]:
    """
    Expand a range reference to individual cell references.

    PERFORMANCE OPTIMIZATION (Sprint 5 Phase A - Story 4B):
    - Memoized via _expand_range_to_cells_cached() to avoid redundant expansion
    - Supports sparse intersection: if populated_cells provided, returns only
      cells that actually exist in the sheet (not the entire bounding box)

    Sparse intersection changes worst-case from O(range_area) to O(populated_in_range).
    Example: A1:ZZ10000 with 100 populated cells returns 100 cells, not 260,000.

    Args:
        range_ref: Range reference (e.g., "A1:B2", "Sheet1!A1:B2")
        max_cells: Maximum cells to expand (default 10,000). If range is larger,
                   returns empty list to prevent memory/performance issues.
                   Ignored if populated_cells provided (sparse intersection always safe).
        populated_cells: Optional PopulatedCells object for sparse intersection.
                        If provided, returns only cells that exist in the sheet.
        debug: Enable debug logging (default False, controlled by performance_profile flag)

    Returns:
        List of individual cell references in row-major order, or empty list
        if range exceeds max_cells threshold (when populated_cells not provided).

    Example:
        >>> expand_range_to_cells("A1:B2")
        ['A1', 'B1', 'A2', 'B2']
        >>> expand_range_to_cells("A1:Z100000")  # Too large
        []
        >>> expand_range_to_cells("A1:Z100000", populated_cells=pc)  # Sparse
        ['A5', 'A10', 'B20']  # Only populated cells in range

    Note:
        This is used for traversal to discover formula cells within ranges.
        Large ranges are handled efficiently via sparse intersection when available.
    """
    import sys

    # SPARSE INTERSECTION PATH (Story 4B)
    if populated_cells is not None:
        result = _expand_range_sparse(range_ref, populated_cells)
        if debug:
            print(
                f"  DEBUG SPARSE: {range_ref} → {len(result)} cells (from {len(populated_cells.all_cells_sorted)} populated)",
                file=sys.stderr,
            )
        return result

    # LEGACY PATH (backward compatibility, used when populated_cells not available)
    if debug:
        print(f"  DEBUG LEGACY: {range_ref} (no populated_cells provided)", file=sys.stderr)

    # OPTIMIZATION FIX 1: Memoization via lru_cache
    # Delegate to cached implementation and convert tuple back to list
    return list(_expand_range_to_cells_cached(range_ref, max_cells))


def categorize_references(refs: list[str]) -> dict[str, list[str]]:
    """
    Categorize references by type for structured processing.

    Args:
        refs: List of reference strings

    Returns:
        Dictionary with keys:
        - 'cell_refs': A1-style cell/range references
        - 'structured_refs': Table references (Table1[Column])
        - 'defined_names': Potential defined names

    Example:
        >>> categorize_references(['A1', 'Table1[Rev]', 'MyVar', 'B1:C10'])
        {'cell_refs': ['A1', 'B1:C10'], 'structured_refs': ['Table1[Rev]'], 'defined_names': ['MyVar']}
    """
    result: dict[str, list[str]] = {"cell_refs": [], "structured_refs": [], "defined_names": []}

    for ref in refs:
        if is_structured_reference(ref):
            result["structured_refs"].append(ref)
        elif is_defined_name(ref):
            result["defined_names"].append(ref)
        else:
            result["cell_refs"].append(ref)

    # Sort each category for deterministic output
    for key in result:
        result[key] = sorted(result[key])

    return result


def is_range_reference(ref: str) -> bool:
    """Check if reference is a range (contains ':').

    Args:
        ref: Cell or range reference (e.g., "A1", "A1:B10", "Sheet!A1:B10")

    Returns:
        True if reference is a range, False if single cell.
    """
    return ":" in ref


def get_range_cell_count(ref: str) -> int:
    """Count cells in a range reference.

    Args:
        ref: Range reference (e.g., "A1:B10", "Sheet!A1:C5")

    Returns:
        Number of cells in range. Returns 1 for single cell references.
    """
    if not is_range_reference(ref):
        return 1

    # Handle sheet-qualified references
    range_part = ref
    if "!" in ref:
        parts = ref.split("!", 1)
        range_part = parts[1]

    try:
        start_str, end_str = range_part.split(":", 1)
        start_ref = parse_a1_reference(start_str)
        end_ref = parse_a1_reference(end_str)

        if start_ref is None or end_ref is None:
            return 1

        num_rows = abs(end_ref.row - start_ref.row) + 1
        num_cols = abs(end_ref.col - start_ref.col) + 1
        return num_rows * num_cols
    except Exception:
        # Fallback for complex/invalid ranges
        return 1
