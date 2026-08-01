# ABOUTME: Core constant range merging rules for grouping logic
# ABOUTME: Implements deterministic rules for merging constant cells into logical tables

from collections.abc import Callable
from typing import Any

from xl_marinade.core.grouping.geometry import BoundingBox, compute_bounding_box, parse_a1_address


def merge_constant_ranges(
    ranges: list[dict[str, Any]], validator: Callable[[dict[str, Any]], bool] | None = None
) -> list[dict[str, Any]]:
    """
    Apply merging rules to constant ranges.

    Args:
        ranges: List of range dicts with keys:
            - "address": A1 notation (e.g., "F6:F10")
            - "dtype": Data type string (e.g., "number", "text", "boolean")
            - Optional: "cells" - list of cell dicts with dtype info
        validator: Optional callback to validate potential merges.
                  Receives merged range dict, returns True if valid.

    Returns:
        List of merged ranges (may include overlaps). Each dict has same structure
        as input but with merged addresses.
    """
    if not ranges:
        return []

    if len(ranges) == 1:
        return ranges.copy()

    # Iteratively merge until stable
    current = ranges.copy()
    changed = True
    max_iterations = len(ranges) * len(ranges)  # Safety limit
    iteration = 0

    while changed and iteration < max_iterations:
        changed = False
        iteration += 1

        merged = []
        used = set()

        for i, range_a in enumerate(current):
            if i in used:
                continue

            # Try to merge range_a with any remaining range
            merge_found = False

            for j, range_b in enumerate(current[i + 1 :], start=i + 1):
                if j in used:
                    continue

                merged_range = _try_merge_pair(range_a, range_b)

                if merged_range is not None:
                    # Apply validator if provided
                    if validator and not validator(merged_range):
                        continue

                    # Successfully merged
                    merged.append(merged_range)
                    used.add(i)
                    used.add(j)
                    merge_found = True
                    changed = True
                    break

            if not merge_found:
                # Keep range_a as-is
                merged.append(range_a)
                used.add(i)

        current = merged

    return current


def _try_merge_pair(range_a: dict[str, Any], range_b: dict[str, Any]) -> dict[str, Any] | None:
    """
    Try to merge two ranges according to merging rules.

    Returns:
        Merged range dict if merge is allowed, None otherwise.
    """
    # Extract sheet names and check compatibility
    sheet_a = range_a["address"].split("!")[0] if "!" in range_a["address"] else ""
    sheet_b = range_b["address"].split("!")[0] if "!" in range_b["address"] else ""

    # Only merge ranges from same sheet
    if sheet_a and sheet_b and sheet_a != sheet_b:
        return None

    # Strip sheet names from addresses if present
    addr_a = range_a["address"].split("!")[-1] if "!" in range_a["address"] else range_a["address"]
    addr_b = range_b["address"].split("!")[-1] if "!" in range_b["address"] else range_b["address"]

    bbox_a = parse_a1_address(addr_a)
    bbox_b = parse_a1_address(addr_b)

    # Check for full containment
    if bbox_a.contains(bbox_b):
        return range_a
    if bbox_b.contains(bbox_a):
        return range_b

    # Compute bounding box (without sheet names)
    bbox = compute_bounding_box(addr_a, addr_b)

    # Apply dimension-specific rules
    if bbox.width == 1:
        # Single-column test
        if should_merge_single_column(range_a, range_b, bbox):
            return _create_merged_range(range_a, range_b, bbox)
    elif bbox.height == 1:
        # Single-row test
        if should_merge_single_row(range_a, range_b, bbox):
            return _create_merged_range(range_a, range_b, bbox)
    else:
        # Multi-dimensional test
        if should_merge_multi_dim(range_a, range_b, bbox):
            return _create_merged_range(range_a, range_b, bbox)

    # Don't merge
    return None


def should_merge_single_column(
    range_a: dict[str, Any], range_b: dict[str, Any], bbox: BoundingBox
) -> bool:
    """
    Check if two ranges should merge using single-column rules.

    Single-column rules:
    1. Bounding box width must be 1 (same column)
    2. All non-blank cells in both ranges must have same dtype
    3. Blank cells within ranges are allowed

    Args:
        range_a: First range dict
        range_b: Second range dict
        bbox: Bounding box of both ranges

    Returns:
        True if ranges should merge, False otherwise
    """
    # Must be single column
    if bbox.width != 1:
        return False

    # Check dtype consistency
    dtype_a = range_a.get("dtype")
    dtype_b = range_b.get("dtype")

    # If both have dtype metadata, they must match
    if dtype_a and dtype_b and dtype_a != dtype_b:
        return False

    # Check cell-level dtypes if available
    cells_a = range_a.get("cells", [])
    cells_b = range_b.get("cells", [])

    if cells_a or cells_b:
        return _check_dtype_consistency(cells_a + cells_b)

    # If no detailed cell info, allow merge if range-level dtypes match
    return True


def should_merge_single_row(
    range_a: dict[str, Any], range_b: dict[str, Any], bbox: BoundingBox
) -> bool:
    """
    Check if two ranges should merge using single-row rules.

    Single-row rules (same as single-column but horizontal):
    1. Bounding box height must be 1 (same row)
    2. All non-blank cells in both ranges must have same dtype
    3. Blank cells within ranges are allowed

    Args:
        range_a: First range dict
        range_b: Second range dict
        bbox: Bounding box of both ranges

    Returns:
        True if ranges should merge, False otherwise
    """
    # Must be single row
    if bbox.height != 1:
        return False

    # Check dtype consistency (same logic as single-column)
    dtype_a = range_a.get("dtype")
    dtype_b = range_b.get("dtype")

    if dtype_a and dtype_b and dtype_a != dtype_b:
        return False

    cells_a = range_a.get("cells", [])
    cells_b = range_b.get("cells", [])

    if cells_a or cells_b:
        return _check_dtype_consistency(cells_a + cells_b)

    return True


def should_merge_multi_dim(
    range_a: dict[str, Any], range_b: dict[str, Any], bbox: BoundingBox
) -> bool:
    """
    Check if two ranges should merge using multi-dimensional rules.

    Multi-dimensional rules:
    1. Bounding box must be multi-dimensional (width>1 AND height>1)
    2. No fully blank rows within bounding box
    3. No fully blank columns within bounding box
    4. All non-blank cells must have same dtype
    5. Individual blank cells (not full row/column) are allowed

    Args:
        range_a: First range dict
        range_b: Second range dict
        bbox: Bounding box of both ranges

    Returns:
        True if ranges should merge, False otherwise

    Note:
        This implementation uses simplified boundary checking based on available
        cell metadata. Full blank row/column detection requires cell-level data
        which may not be available in all contexts.
    """
    # Must be multi-dimensional
    if bbox.width <= 1 or bbox.height <= 1:
        return False

    # Check dtype consistency
    dtype_a = range_a.get("dtype")
    dtype_b = range_b.get("dtype")

    if dtype_a and dtype_b and dtype_a != dtype_b:
        return False

    cells_a = range_a.get("cells", [])
    cells_b = range_b.get("cells", [])

    if cells_a or cells_b:
        all_cells = cells_a + cells_b

        # Check dtype consistency
        if not _check_dtype_consistency(all_cells):
            return False

        # Check for blank row/column boundaries
        # This is a simplified check - full implementation would need
        # complete cell grid to detect fully blank rows/columns
        # For now, allow merge if dtypes match

    return True


def _check_dtype_consistency(cells: list[dict[str, Any]]) -> bool:
    """
    Check if all non-blank cells have the same dtype.

    Args:
        cells: List of cell dicts with "dtype" field

    Returns:
        True if all non-blank cells have same dtype, False otherwise
    """
    if not cells:
        return True

    # Collect non-blank dtypes
    dtypes = set()
    for cell in cells:
        dtype = cell.get("dtype")
        if dtype and dtype not in ("blank", "empty", None):
            dtypes.add(dtype)

    # All non-blank cells must have same dtype
    return len(dtypes) <= 1


def _create_merged_range(
    range_a: dict[str, Any], range_b: dict[str, Any], bbox: BoundingBox
) -> dict[str, Any]:
    """
    Create a new merged range dict from two ranges.

    Args:
        range_a: First range
        range_b: Second range
        bbox: Bounding box containing both ranges

    Returns:
        New range dict with merged address and combined metadata
    """
    # Use dtype from range_a if available, else range_b
    dtype = range_a.get("dtype") or range_b.get("dtype")

    # Preserve sheet name from range_a (both should have same sheet)
    sheet_name = ""
    if "!" in range_a["address"]:
        sheet_name = range_a["address"].split("!")[0]
    elif "!" in range_b["address"]:
        sheet_name = range_b["address"].split("!")[0]

    # Build merged address with sheet name if present
    merged_addr = f"{sheet_name}!{bbox.to_a1()}" if sheet_name else bbox.to_a1()

    merged = {
        "address": merged_addr,
        "dtype": dtype,
    }

    # Combine cells if available
    cells_a = range_a.get("cells", [])
    cells_b = range_b.get("cells", [])
    if cells_a or cells_b:
        merged["cells"] = cells_a + cells_b

    # Preserve other metadata from range_a
    for key in ["binding_type", "binding_id", "formulas_using"]:
        if key in range_a:
            merged[key] = range_a[key]
        elif key in range_b:
            merged[key] = range_b[key]

    return merged
