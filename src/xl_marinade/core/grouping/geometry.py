# ABOUTME: Bounding box geometry helpers for constant range grouping
# ABOUTME: Provides BoundingBox class and utilities for computing range dimensions and overlaps

import re
from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class BoundingBox:
    """
    Represents a rectangular bounding box in Excel coordinates.

    Attributes:
        min_row: Minimum row (1-indexed)
        max_row: Maximum row (1-indexed, inclusive)
        min_col: Minimum column (1-indexed)
        max_col: Maximum column (1-indexed, inclusive)
    """

    min_row: int
    max_row: int
    min_col: int
    max_col: int

    @property
    def width(self) -> int:
        """Number of columns in bounding box."""
        return self.max_col - self.min_col + 1

    @property
    def height(self) -> int:
        """Number of rows in bounding box."""
        return self.max_row - self.min_row + 1

    @property
    def cell_count(self) -> int:
        """Total number of cells in bounding box."""
        return self.width * self.height

    def to_a1(self) -> str:
        """Convert bounding box to A1 notation."""
        start = _col_num_to_letter(self.min_col) + str(self.min_row)
        end = _col_num_to_letter(self.max_col) + str(self.max_row)
        if start == end:
            return start
        return f"{start}:{end}"

    def contains(self, other: "BoundingBox") -> bool:
        """Check if this bounding box fully contains another."""
        return (
            self.min_row <= other.min_row
            and self.max_row >= other.max_row
            and self.min_col <= other.min_col
            and self.max_col >= other.max_col
        )

    def overlaps(self, other: "BoundingBox") -> bool:
        """Check if this bounding box overlaps with another."""
        return not (
            self.max_row < other.min_row
            or self.min_row > other.max_row
            or self.max_col < other.min_col
            or self.min_col > other.max_col
        )

    def intersection(self, other: "BoundingBox") -> Optional["BoundingBox"]:
        """Compute intersection of two bounding boxes, or None if no overlap."""
        if not self.overlaps(other):
            return None

        return BoundingBox(
            min_row=max(self.min_row, other.min_row),
            max_row=min(self.max_row, other.max_row),
            min_col=max(self.min_col, other.min_col),
            max_col=min(self.max_col, other.max_col),
        )


def parse_a1_address(address: str) -> BoundingBox:
    """
    Parse an A1-style address into a BoundingBox.

    Args:
        address: A1 notation like "F6", "F6:F10", "A1:B4"

    Returns:
        BoundingBox representing the range

    Raises:
        ValueError: If address format is invalid

    Examples:
        >>> parse_a1_address("F6")
        BoundingBox(min_row=6, max_row=6, min_col=6, max_col=6)
        >>> parse_a1_address("F6:F10")
        BoundingBox(min_row=6, max_row=10, min_col=6, max_col=6)
        >>> parse_a1_address("A1:B4")
        BoundingBox(min_row=1, max_row=4, min_col=1, max_col=2)
    """
    address = address.strip()

    if ":" in address:
        start, end = address.split(":", 1)
        start_row, start_col = _parse_cell_address(start)
        end_row, end_col = _parse_cell_address(end)

        return BoundingBox(
            min_row=min(start_row, end_row),
            max_row=max(start_row, end_row),
            min_col=min(start_col, end_col),
            max_col=max(start_col, end_col),
        )
    else:
        row, col = _parse_cell_address(address)
        return BoundingBox(min_row=row, max_row=row, min_col=col, max_col=col)


def compute_bounding_box(*addresses: str) -> BoundingBox:
    """
    Compute bounding box that contains all given addresses.

    Args:
        *addresses: One or more A1-style addresses

    Returns:
        BoundingBox containing all addresses

    Raises:
        ValueError: If no addresses provided or any address is invalid

    Examples:
        >>> compute_bounding_box("F6:F10", "F11:F20")
        BoundingBox(min_row=6, max_row=20, min_col=6, max_col=6)
        >>> compute_bounding_box("A1:A3", "B2:B4")
        BoundingBox(min_row=1, max_row=4, min_col=1, max_col=2)
    """
    if not addresses:
        raise ValueError("At least one address required")

    boxes = [parse_a1_address(addr) for addr in addresses]

    return BoundingBox(
        min_row=min(box.min_row for box in boxes),
        max_row=max(box.max_row for box in boxes),
        min_col=min(box.min_col for box in boxes),
        max_col=max(box.max_col for box in boxes),
    )


def _parse_cell_address(cell: str) -> tuple[int, int]:
    """
    Parse a single cell address like 'F6' into (row, col) 1-indexed.

    Returns:
        Tuple of (row_number, col_number) both 1-indexed

    Raises:
        ValueError: If cell address format is invalid
    """
    cell = cell.strip().upper().replace("$", "")
    match = re.match(r"^([A-Z]+)(\d+)$", cell)

    if not match:
        raise ValueError(f"Invalid cell address: {cell}")

    col_letter, row_str = match.groups()
    row = int(row_str)
    col = _col_letter_to_num(col_letter)

    return row, col


def _col_letter_to_num(col_letter: str) -> int:
    """Convert column letter to 1-indexed number (A=1, Z=26, AA=27)."""
    num = 0
    for char in col_letter:
        num = num * 26 + (ord(char) - ord("A") + 1)
    return num


def _col_num_to_letter(col_num: int) -> str:
    """Convert 1-indexed column number to letter (1=A, 26=Z, 27=AA)."""
    letter = ""
    while col_num > 0:
        col_num -= 1
        letter = chr(col_num % 26 + ord("A")) + letter
        col_num //= 26
    return letter
