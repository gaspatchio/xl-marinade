# ABOUTME: Packed 64-bit cell identity encoding for memory-efficient cell addressing.
# ABOUTME: Encodes (sheet_id, row, col) as a single integer with deterministic bit layout.

"""
Cell Identity Encoding

Layout: [20 bits sheet_id][24 bits row][20 bits col]
Encoding: cell_id = (sheet_id << 44) | (row << 20) | col

All inputs are 1-based:
- sheet_id: 1..1,048,575 (20 bits)
- row: 1..1,048,576 (24 bits, Excel max)
- col: 1..16,384 (20 bits, Excel max)

Overflow/underflow: raises ValueError with clear error message.
"""

# Excel limits (1-based)
MAX_SHEET_ID = (1 << 20) - 1  # 1,048,575
MAX_ROW = 1_048_576
MAX_COL = 16_384

# Bit layout
SHEET_SHIFT = 44
ROW_SHIFT = 20


def pack(sheet_id: int, row: int, col: int) -> int:
    """
    Pack (sheet_id, row, col) into a 64-bit cell_id.

    Args:
        sheet_id: 1-based sheet identifier (1..1,048,575)
        row: 1-based row (1..1,048,576)
        col: 1-based column (1..16,384)

    Returns:
        64-bit integer cell_id

    Raises:
        ValueError: if any component is out of bounds
    """
    # Validate bounds
    if not (1 <= sheet_id <= MAX_SHEET_ID):
        raise ValueError(f"sheet_id out of bounds: {sheet_id} (must be 1..{MAX_SHEET_ID})")
    if not (1 <= row <= MAX_ROW):
        raise ValueError(f"row out of bounds: {row} (must be 1..{MAX_ROW})")
    if not (1 <= col <= MAX_COL):
        raise ValueError(f"col out of bounds: {col} (must be 1..{MAX_COL})")

    # Pack
    return (sheet_id << SHEET_SHIFT) | (row << ROW_SHIFT) | col


def unpack(cell_id: int) -> tuple[int, int, int]:
    """
    Unpack cell_id into (sheet_id, row, col).

    Args:
        cell_id: 64-bit integer cell identifier

    Returns:
        Tuple of (sheet_id, row, col), all 1-based
    """
    sheet_id = (cell_id >> SHEET_SHIFT) & ((1 << 20) - 1)
    row = (cell_id >> ROW_SHIFT) & ((1 << 24) - 1)
    col = cell_id & ((1 << 20) - 1)

    return sheet_id, row, col


def col_to_a1(col: int) -> str:
    """
    Convert 1-based column number to Excel column letters (A, B, ..., Z, AA, ...).

    Args:
        col: 1-based column number (1..16,384)

    Returns:
        Column letters (uppercase)

    Raises:
        ValueError: if col is out of bounds
    """
    if not (1 <= col <= MAX_COL):
        raise ValueError(f"col out of bounds: {col} (must be 1..{MAX_COL})")

    result = []
    col -= 1  # Convert to 0-based for modulo arithmetic

    while True:
        result.append(chr(ord("A") + (col % 26)))
        col //= 26
        if col == 0:
            break
        col -= 1  # Adjust for Excel's base-26 system

    return "".join(reversed(result))


def a1_to_col(col_str: str) -> int:
    """
    Convert Excel column letters to 1-based column number.

    Args:
        col_str: Column letters (case-insensitive)

    Returns:
        1-based column number

    Raises:
        ValueError: if col_str is invalid or out of bounds
    """
    col_str = col_str.upper()

    if not col_str or not col_str.isalpha():
        raise ValueError(f"Invalid column string: {col_str}")

    col = 0
    for char in col_str:
        col = col * 26 + (ord(char) - ord("A") + 1)

    if not (1 <= col <= MAX_COL):
        raise ValueError(f"Column {col_str} out of bounds (col={col}, max={MAX_COL})")

    return col


def to_a1(row: int, col: int) -> str:
    """
    Convert (row, col) to A1 notation (e.g., "A1", "AA100").

    Args:
        row: 1-based row (1..1,048,576)
        col: 1-based column (1..16,384)

    Returns:
        A1 string (uppercase, no sheet prefix, no $)

    Raises:
        ValueError: if row or col is out of bounds
    """
    if not (1 <= row <= MAX_ROW):
        raise ValueError(f"row out of bounds: {row} (must be 1..{MAX_ROW})")

    return f"{col_to_a1(col)}{row}"


def from_a1(a1: str) -> tuple[int, int]:
    """
    Parse A1 notation to (row, col).

    Args:
        a1: A1 string (e.g., "A1", "AA100", case-insensitive)

    Returns:
        Tuple of (row, col), both 1-based

    Raises:
        ValueError: if a1 is invalid or out of bounds
    """
    a1 = a1.upper()

    # Split into column letters and row digits
    i = 0
    while i < len(a1) and a1[i].isalpha():
        i += 1

    if i == 0 or i == len(a1):
        raise ValueError(f"Invalid A1 notation: {a1}")

    col_str = a1[:i]
    row_str = a1[i:]

    try:
        row = int(row_str)
    except ValueError:
        raise ValueError(f"Invalid row in A1 notation: {a1}")

    col = a1_to_col(col_str)

    if not (1 <= row <= MAX_ROW):
        raise ValueError(f"row out of bounds: {row} (must be 1..{MAX_ROW})")

    return row, col
