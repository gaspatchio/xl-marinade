# ABOUTME: Range matching utilities for semantic formula generation
# ABOUTME: Uses BoundingBox from xl_marinade.core for O(n×m) containment checks

import re
from dataclasses import dataclass

# REUSE existing infrastructure - don't duplicate!
from xl_marinade.core.grouping.geometry import BoundingBox, parse_a1_address


@dataclass(frozen=True)
class SheetRange:
    """
    A BoundingBox with optional sheet context.

    Wraps BoundingBox to add sheet-aware matching (case-insensitive).
    Uses existing BoundingBox.contains() for geometric containment.
    """

    sheet: str | None  # Normalized to lowercase, or None for current sheet
    bounds: BoundingBox

    def __post_init__(self) -> None:
        # Normalize sheet name to lowercase for case-insensitive matching
        if self.sheet:
            object.__setattr__(self, "sheet", self.sheet.lower().strip("'\""))

    def contains(self, other: "SheetRange") -> bool:
        """Check if this range fully contains another (sheet-aware)."""
        # Sheet comparison (already lowercase due to __post_init__)
        if self.sheet and other.sheet and self.sheet != other.sheet:
            return False
        # If self has no sheet, it matches any sheet (context-dependent) - wait, is that right?
        # The spec says: "If ref.sheet and binding.sheet ... if ref.sheet.lower() != binding.sheet.lower(): return False"
        # If binding (self) has sheet, ref (other) must have same sheet.
        # If binding (self) has NO sheet (local name?), it probably matches only local refs?
        # Let's follow the logic in the spec:

        # Spec logic:
        # if ref.sheet and binding.sheet:
        #     if ref.sheet.lower() != binding.sheet.lower():
        #         return False

        # What if one has sheet and other doesn't?
        # Usually binding has sheet. Ref might or might not.

        # Let's stick to the spec's `ranges_match` logic but adapted for this class.
        # "contains" usually means strict containment.

        if self.sheet and other.sheet and self.sheet != other.sheet:
            return False

        # Delegate geometric containment to existing BoundingBox
        return self.bounds.contains(other.bounds)

    def matches(self, other: "SheetRange") -> bool:
        """Check if ranges are equal or this contains other."""
        if self.sheet == other.sheet and self.bounds == other.bounds:
            return True
        return self.contains(other)


def parse_range_with_sheet(range_str: str) -> SheetRange | None:
    """
    Parse A1 range to SheetRange, or None if invalid.

    Handles:
    - Single cell: "A1" → SheetRange(None, BoundingBox(1,1,1,1))
    - Range: "A1:B10" → SheetRange(None, BoundingBox(1,10,1,2))
    - With sheet: "Sheet1!A1:B10" → SheetRange("sheet1", BoundingBox(1,10,1,2))
    - Quoted sheet: "'Sheet Name'!A1:B10" → strips quotes
    - Absolute refs: "$A$1:$B$10" → strips $ signs

    Uses existing parse_a1_address() from geometry.py for parsing.
    """
    if not range_str or not range_str.strip():
        return None

    range_str = range_str.strip().replace("$", "")

    # Extract sheet name if present
    sheet = None
    address = range_str
    if "!" in range_str:
        parts = range_str.split("!", 1)
        sheet = parts[0].strip("'\"")  # Remove quotes around sheet name
        address = parts[1]

    # Check for named range (no numbers) or 3D reference
    # A named range won't have digits usually, or might look like a name.
    # parse_a1_address expects A1 notation.
    # Simple check: address must contain digits.
    if not re.search(r"\d", address) or ":" in (sheet or ""):
        return None

    try:
        bounds = parse_a1_address(address)  # REUSE existing parser
        return SheetRange(sheet=sheet, bounds=bounds)
    except ValueError:
        return None


def find_binding_for_range(
    range_ref: str,
    current_sheet: str | None,
    binding_ranges: dict[tuple[str | None, BoundingBox], str],
) -> str | None:
    """
    Find binding label that contains or exactly matches range_ref.

    Returns the label of the FIRST binding whose range contains the reference,
    or None if no binding matches.
    """
    ref_range = parse_range_with_sheet(range_ref)
    if not ref_range:
        return None

    # Apply current sheet if ref has no sheet
    ref_sheet = ref_range.sheet or (current_sheet.lower() if current_sheet else None)

    for (binding_sheet, binding_bounds), label in binding_ranges.items():
        binding_sr = SheetRange(sheet=binding_sheet, bounds=binding_bounds)
        ref_sr = SheetRange(sheet=ref_sheet, bounds=ref_range.bounds)

        if binding_sr.matches(ref_sr):
            return label

    return None
