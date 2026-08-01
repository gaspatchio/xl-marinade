# ABOUTME: Input validation for workbook paths and roots JSON structure per ADR-009.
# ABOUTME: Enforces single-root constraint, A1 notation, contiguity, and file format validation.
"""Input validation for IR extraction."""

import re
from pathlib import Path
from typing import Any


def validate_workbook_path(path: Path) -> None:
    """
    Validate workbook path meets requirements.

    Requirements:
    - Must exist
    - Must have .xlsx or .xlsm extension (Story 10: .xlsm for VBA support)
    - Must be a file (not directory)

    Args:
        path: Path to workbook file.

    Raises:
        FileNotFoundError: If file doesn't exist.
        ValueError: If path is invalid (wrong extension, not a file).
    """
    if not path.exists():
        raise FileNotFoundError(f"Workbook not found: {path}")

    if not path.is_file():
        raise ValueError(f"Workbook path is not a file: {path}")

    if path.suffix.lower() not in [".xlsx", ".xlsm"]:
        raise ValueError(f"Workbook must have .xlsx or .xlsm extension, got: {path.suffix}")


def validate_a1_range(range_str: str) -> None:
    """
    Validate A1 range notation.

    Valid formats:
    - Single cell: "A1", "Z99", "AA100"
    - Range: "A1:B10", "C5:C5" (single cell range)
    - Spill reference: "A1#" (dynamic array spill)

    Invalid formats:
    - Non-contiguous: "A1:A5,C1:C5"
    - Multi-sheet: "Sheet1:Sheet3!A1"
    - 3D references: "Sheet1:Sheet2!A1:B10"

    Args:
        range_str: A1 range string.

    Raises:
        ValueError: If range format is invalid.
    """
    # Check for multi-region (non-contiguous)
    if "," in range_str:
        raise ValueError(f"Non-contiguous ranges not supported: {range_str}")

    # Check for multi-sheet (3D reference)
    if ":" in range_str and "!" in range_str:
        parts = range_str.split("!")
        if len(parts) > 1 and ":" in parts[0]:
            raise ValueError(f"Multi-sheet (3D) references not supported: {range_str}")

    # Validate A1 notation pattern
    # Pattern: [A-Z]+[0-9]+ or [A-Z]+[0-9]+:[A-Z]+[0-9]+ or [A-Z]+[0-9]+#
    single_cell_pattern = r"^[A-Z]+[0-9]+$"
    range_pattern = r"^[A-Z]+[0-9]+:[A-Z]+[0-9]+$"
    spill_pattern = r"^[A-Z]+[0-9]+#$"

    if not (
        re.match(single_cell_pattern, range_str)
        or re.match(range_pattern, range_str)
        or re.match(spill_pattern, range_str)
    ):
        raise ValueError(f"Invalid A1 range notation: {range_str}")


def validate_roots_json(data: dict[str, Any]) -> None:
    """
    Validate roots JSON structure per ADR-009.

    Required structure:
    {
        "user_root": {
            "sheet": "Sheet1",
            "range": "A1" or "A1:B10",
            "label_hint": "optional string"  # optional field
        }
    }

    Single-root constraint: Must have exactly one root (user_root object, not array).

    Args:
        data: Parsed JSON dictionary.

    Raises:
        ValueError: If structure is invalid or constraints violated.
    """
    # Check for user_root key
    if "user_root" not in data:
        raise ValueError("roots JSON must contain 'user_root' key")

    user_root = data["user_root"]

    # user_root must be a dict (single root), not an array
    if not isinstance(user_root, dict):
        raise ValueError("user_root must be an object (single root), not an array")

    # Check required fields
    if "sheet" not in user_root:
        raise ValueError("user_root must contain 'sheet' field")
    if "range" not in user_root:
        raise ValueError("user_root must contain 'range' field")

    # Validate sheet name
    sheet = user_root["sheet"]
    if not isinstance(sheet, str) or not sheet.strip():
        raise ValueError(f"sheet must be non-empty string, got: {sheet}")

    # Validate range format
    range_str = user_root["range"]
    if not isinstance(range_str, str) or not range_str.strip():
        raise ValueError(f"range must be non-empty string, got: {range_str}")

    validate_a1_range(range_str)

    # Validate optional label_hint
    if "label_hint" in user_root:
        label_hint = user_root["label_hint"]
        if not isinstance(label_hint, str):
            raise ValueError(f"label_hint must be string if present, got: {type(label_hint)}")
