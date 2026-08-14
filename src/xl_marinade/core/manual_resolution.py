# ABOUTME: Manual resolution provider for loading and applying manual lookup resolution overrides
# ABOUTME: Supports JSON-based manual resolution format per semantic lookup resolution design doc §4

import json
from pathlib import Path
from typing import Any


class ManualResolutionProvider:
    """
    Loads and provides manual resolution overrides from JSON.

    Per design doc §4, manual resolutions allow users to specify exact semantic
    dependencies for complex lookup formulas that automatic resolution cannot handle.

    **SCOPE LIMITATION (as of 2025-11-30):**
    Manual resolution currently only applies to **VLOOKUP** functions. Other lookup
    functions (HLOOKUP, XLOOKUP, INDEX, MATCH, CHOOSE) do not check for manual
    resolutions. To extend support, add manual_provider checks in their respective
    resolve_*_semantic methods in resolution.py.

    Attributes:
        resolutions: Dictionary mapping cell_address → resolution_data
    """

    def __init__(self, json_path: Path | None = None) -> None:
        """
        Load manual resolutions from JSON file.

        Args:
            json_path: Path to manual_resolutions.json file (optional)

        Raises:
            ValueError: If JSON file is invalid or cannot be parsed
        """
        self.resolutions: dict[str, dict[str, Any]] = {}
        if json_path and json_path.exists():
            self.resolutions = load_manual_resolutions(json_path)

    def get_resolution(self, cell_address: str) -> dict[str, Any] | None:
        """
        Get manual resolution for cell, or None if not present.

        Args:
            cell_address: Cell address (A1 notation, optionally sheet-qualified)

        Returns:
            Resolution data dict or None if no manual resolution exists

        Example:
            >>> provider.get_resolution("Projection!A10")
            {"resolved_ref": "F6:F607", "reason": "Column index D1 evaluates to 2"}
        """
        return self.resolutions.get(cell_address)


def load_manual_resolutions(json_path: Path) -> dict[str, dict[str, Any]]:
    """
    Parse manual resolutions JSON file.

    Per design doc §4, JSON format (STRICT SCHEMA REQUIRED):
    {
      "version": "1.0",
      "workbook_guid": "optional_guid",
      "resolutions": [
        {
          "cell_address": "Sheet!A10",
          "function": "VLOOKUP",
          "current_status": "conservative_fallback",
          "syntactic_ref": "E6:G607",
          "resolved_ref": "F6:F607",
          "reason": "Column index D1 evaluates to 2 based on model logic"
        }
      ]
    }

    **REQUIRED FIELDS:**
    - Top-level: "resolutions" array (not a flat key-value map)
    - Per resolution: "cell_address" (sheet-qualified), "resolved_ref" (target range)
    - Optional: "function", "current_status", "syntactic_ref", "reason"

    **COMMON ERRORS:**
    - Simplified format like {"Sheet!A1": "Sheet!B1"} will FAIL validation
    - Missing "resolutions" wrapper array will FAIL
    - Unquoted sheet names in addresses may cause parsing issues

    Args:
        json_path: Path to JSON file

    Returns:
        Dictionary mapping cell_address → resolution_data

    Raises:
        ValueError: If JSON is invalid or missing required fields
    """
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON in manual resolutions file: {e}") from e
    except OSError as e:
        raise ValueError(f"Cannot read manual resolutions file: {e}") from e

    # Validate top-level structure
    if not isinstance(data, dict):
        raise ValueError("Manual resolutions JSON must be an object")

    if "resolutions" not in data:
        raise ValueError("Manual resolutions JSON missing 'resolutions' field")

    if not isinstance(data["resolutions"], list):
        raise ValueError("Manual resolutions 'resolutions' field must be a list")

    # Build mapping
    result: dict[str, dict[str, Any]] = {}

    for idx, resolution in enumerate(data["resolutions"]):
        # Validate each resolution entry
        if not isinstance(resolution, dict):
            raise ValueError(f"Resolution entry {idx} must be an object")

        if "cell_address" not in resolution:
            raise ValueError(f"Resolution entry {idx} missing 'cell_address'")

        if "resolved_ref" not in resolution:
            raise ValueError(f"Resolution entry {idx} missing 'resolved_ref'")

        cell_address = resolution["cell_address"]
        if not isinstance(cell_address, str):
            raise ValueError(f"Resolution entry {idx} 'cell_address' must be a string")

        # Store resolution data (excluding cell_address key)
        resolution_data = {
            "resolved_ref": resolution["resolved_ref"],
            "reason": resolution.get("reason", "Manual override"),
            "function": resolution.get("function"),
            "current_status": resolution.get("current_status"),
            "syntactic_ref": resolution.get("syntactic_ref"),
        }

        result[cell_address] = resolution_data

    return result
