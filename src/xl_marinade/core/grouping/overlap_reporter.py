# ABOUTME: Overlap detection and reporting for constant bindings
# ABOUTME: Generates diagnostic reports for overlapping constant ranges

from typing import Any

from xl_marinade.core.grouping.geometry import parse_a1_address


def detect_overlaps(bindings: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Detect overlapping constant bindings.

    Args:
        bindings: List of binding dicts with keys:
            - "address": A1 notation
            - "binding_type": Type of binding (filter for "constant")
            - Optional: "binding_id", other metadata

    Returns:
        List of overlap dicts, each containing:
            - binding_a: First binding info (dict with address, binding_type, etc.)
            - binding_b: Second binding info
            - overlap_region: A1 address of overlap region
            - overlap_cell_count: Number of overlapping cells

    Examples:
        >>> bindings = [
        ...     {"address": "F6:F100", "binding_type": "constant", "binding_id": "hash1"},
        ...     {"address": "F50:F120", "binding_type": "constant", "binding_id": "hash2"}
        ... ]
        >>> overlaps = detect_overlaps(bindings)
        >>> overlaps[0]["overlap_region"]
        'F50:F100'
        >>> overlaps[0]["overlap_cell_count"]
        51
    """
    overlaps = []

    # Filter to constant bindings only
    constant_bindings = [b for b in bindings if b.get("binding_type") == "constant"]

    # Check each pair for overlap
    for i, binding_a in enumerate(constant_bindings):
        for binding_b in constant_bindings[i + 1 :]:
            overlap_region = compute_overlap(binding_a["address"], binding_b["address"])

            if overlap_region:
                overlap_count = count_cells(overlap_region)
                size_a = count_cells(binding_a["address"])
                size_b = count_cells(binding_b["address"])

                rationale = _infer_rationale(overlap_count, size_a, size_b)

                overlaps.append(
                    {
                        "binding_a": _extract_binding_info(binding_a),
                        "binding_b": _extract_binding_info(binding_b),
                        "overlap_region": overlap_region,
                        "overlap_cell_count": overlap_count,
                        "rationale": rationale,
                    }
                )

    return overlaps


def _infer_rationale(overlap_count: int, size_a: int, size_b: int) -> str:
    """
    Infer rationale for overlap based on sizes.
    """
    if overlap_count == size_a and overlap_count == size_b:
        return "Duplicate bindings covering identical range."
    elif overlap_count == size_a:
        return "Binding A is fully contained within Binding B."
    elif overlap_count == size_b:
        return "Binding B is fully contained within Binding A."
    else:
        return "Partial overlap between bindings."


def generate_overlap_report(overlaps: list[dict[str, Any]]) -> dict[str, Any]:
    """
    Generate overlap report dict matching design doc §7 format.

    Args:
        overlaps: List of overlap dicts from detect_overlaps()

    Returns:
        Report dict suitable for JSON serialization with structure:
        {
            "summary": {
                "total_constant_bindings": int,
                "bindings_with_overlaps": int,
                "overlap_pairs": int
            },
            "overlaps": list of overlap dicts,
            "recommendations": list of strings
        }

    Note:
        Story 6 will write this to output/binding_overlaps.json.
    """
    # Count unique bindings involved in overlaps
    bindings_with_overlaps = set()
    for overlap in overlaps:
        binding_a_id = overlap["binding_a"].get("binding_id")
        binding_b_id = overlap["binding_b"].get("binding_id")
        if binding_a_id:
            bindings_with_overlaps.add(binding_a_id)
        if binding_b_id:
            bindings_with_overlaps.add(binding_b_id)

    return {
        "summary": {
            "total_constant_bindings": None,  # Caller should fill this in
            "bindings_with_overlaps": len(bindings_with_overlaps),
            "overlap_pairs": len(overlaps),
        },
        "overlaps": overlaps,
        "recommendations": [
            "Review overlap regions - may indicate distinct logical tables",
            "Consider manual resolution if overlaps are unintended",
        ],
    }


def compute_overlap(address_a: str, address_b: str) -> str | None:
    """
    Compute overlap region between two A1 addresses.

    Args:
        address_a: First A1 address (e.g., "F6:F100" or "Sheet1!F6:F100")
        address_b: Second A1 address (e.g., "F50:F120" or "Sheet1!F50:F120")

    Returns:
        A1 address of overlap region, or None if no overlap

    Examples:
        >>> compute_overlap("F6:F100", "F50:F120")
        'F50:F100'
        >>> compute_overlap("A1:A10", "B1:B10")
        None
    """
    # Strip sheet names if present
    sheet_a = ""
    addr_a = address_a
    if "!" in address_a:
        sheet_a, addr_a = address_a.split("!", 1)
        sheet_a = sheet_a.replace("'", "")

    sheet_b = ""
    addr_b = address_b
    if "!" in address_b:
        sheet_b, addr_b = address_b.split("!", 1)
        sheet_b = sheet_b.replace("'", "")

    # Check sheet equality if both have sheet names
    if sheet_a and sheet_b and sheet_a.lower() != sheet_b.lower():
        return None

    bbox_a = parse_a1_address(addr_a)
    bbox_b = parse_a1_address(addr_b)

    intersection = bbox_a.intersection(bbox_b)

    if intersection:
        # Preserve sheet name from first address
        overlap_addr = intersection.to_a1()
        if sheet_a:
            return f"{sheet_a}!{overlap_addr}"
        return overlap_addr

    return None


def count_cells(address: str) -> int:
    """
    Count number of cells in an A1 address range.

    Args:
        address: A1 notation (e.g., "F6:F100", "A1", or "Sheet1!F6:F100")

    Returns:
        Number of cells in range

    Examples:
        >>> count_cells("F6:F100")
        95
        >>> count_cells("A1:B4")
        8
        >>> count_cells("A1")
        1
    """
    # Strip sheet name if present
    addr = address.split("!")[-1] if "!" in address else address
    bbox = parse_a1_address(addr)
    return bbox.cell_count


def _extract_binding_info(binding: dict[str, Any]) -> dict[str, Any]:
    """
    Extract relevant binding info for overlap report.

    Args:
        binding: Full binding dict

    Returns:
        Dict with keys: binding_id, address, binding_type
    """
    return {
        "binding_id": binding.get("binding_id"),
        "address": binding["address"],
        "binding_type": binding.get("binding_type", "constant"),
    }
