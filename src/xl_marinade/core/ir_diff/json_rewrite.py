# ABOUTME: Address-bearing JSON blob canonicalization for the diff algorithm.
# ABOUTME: Rewrites sheet-qualified addresses in evidence, spatial, destinations, extras blobs.

from __future__ import annotations

import json
import re
from collections.abc import Mapping

from openpyxl.utils import column_index_from_string, get_column_letter

from xl_marinade.core.ir_diff.formula_rewrite import AxisIndex


def rewrite_json_blob(
    json_text: str | None,
    sheet_map: dict[str, str],
    row_map: dict[int, int | None],
    col_map: dict[int, int | None],
) -> str | None:
    """Rewrite sheet-qualified addresses in a JSON blob.

    Recursively walks the JSON structure and rewrites any string values that
    look like sheet-qualified A1 addresses (e.g., "Sheet1!A1:B10").

    Args:
        json_text: JSON string to rewrite, or None.
        sheet_map: A sheet name -> B sheet name.
        row_map: A row -> B row.
        col_map: A col -> B col.

    Returns:
        Rewritten JSON string with sorted keys, or None if input was None.
    """
    if json_text is None:
        return None

    try:
        data = json.loads(json_text)
    except (json.JSONDecodeError, TypeError):
        return json_text

    rewritten = _rewrite_value(data, sheet_map, row_map, col_map)
    return json.dumps(rewritten, sort_keys=True, separators=(",", ":"))


_A1_CELL = re.compile(r"(\$?)([A-Z]+)(\$?)(\d+)")


def _move_address_cells(s: str, axis_index: Mapping[str, tuple[AxisIndex, AxisIndex]]) -> str:
    """Move the cells of sheet-qualified A1 addresses through axis maps."""

    def _move(m: re.Match[str]) -> str:
        sheet = m.group(1) or m.group(2)
        indices = axis_index.get(sheet)
        if indices is None:
            return m.group(0)
        rows, cols = indices
        moved = []
        for cell in (m.group(3), m.group(4)):
            if cell is None:
                continue
            c = _A1_CELL.fullmatch(cell)
            if c is None:
                return m.group(0)
            new_row = rows.map(int(c.group(4)))
            new_col = cols.map(column_index_from_string(c.group(2)))
            if new_row is None or new_col is None:
                return m.group(0)  # An endpoint was deleted: the range changed
            moved.append(f"{c.group(1)}{get_column_letter(new_col)}{c.group(3)}{new_row}")
        return m.group(0)[: m.start(3) - m.start(0)] + ":".join(moved)

    return _A1_PATTERN.sub(_move, s)


def _rewrite_value(value, sheet_map, row_map, col_map):
    """Recursively rewrite addresses in a JSON value."""
    if isinstance(value, str):
        return _rewrite_address_string(value, sheet_map, row_map, col_map)
    elif isinstance(value, dict):
        return {k: _rewrite_value(v, sheet_map, row_map, col_map) for k, v in value.items()}
    elif isinstance(value, list):
        return [_rewrite_value(v, sheet_map, row_map, col_map) for v in value]
    return value


# Pattern for A1 addresses: Sheet!A1 or Sheet!A1:B10 or just A1:B10
_A1_PATTERN = re.compile(
    r"""
    (?:                          # Optional sheet prefix
        (?:'([^']+)')            #   Quoted sheet name (group 1)
        |                        #   or
        ([A-Za-z_]\w*)           #   Simple sheet name (group 2)
    )
    !                            # Separator
    (\$?[A-Z]+\$?\d+)           # First cell (group 3)
    (?::(\$?[A-Z]+\$?\d+))?     # Optional second cell (group 4)
    """,
    re.VERBOSE,
)


def _rewrite_address_string(
    s: str,
    sheet_map: dict[str, str],
    row_map: dict[int, int | None],
    col_map: dict[int, int | None],
) -> str:
    """Rewrite sheet-qualified addresses in a string value."""

    def _replace(m: re.Match) -> str:
        sheet = m.group(1) or m.group(2)
        cell1 = m.group(3)
        cell2 = m.group(4)

        new_sheet = sheet_map.get(sheet, sheet)

        # Quote sheet name if it contains spaces
        if " " in new_sheet:
            prefix = f"'{new_sheet}'!"
        else:
            prefix = f"{new_sheet}!"

        if cell2:
            return f"{prefix}{cell1}:{cell2}"
        return f"{prefix}{cell1}"

    return _A1_PATTERN.sub(_replace, s)


def canonicalize_destinations_json(
    destinations_json: str,
    sheet_map: dict[str, str],
    axis_index: Mapping[str, tuple[AxisIndex, AxisIndex]] | None = None,
) -> str:
    """Canonicalize a defined_names.destinations JSON list.

    Rewrites sheet names in destination addresses and, when ``axis_index``
    (A sheet name -> (row index, col index)) is given, moves each cell of a
    destination through that sheet's row/column maps -- so a name whose
    range Excel shifted left when columns were deleted reads as unchanged
    (#53). A destination with an endpoint in a deleted row/column is left
    as it was: its range really did change.

    Args:
        destinations_json: JSON list of A1 ref strings.
        sheet_map: A sheet name -> B sheet name.
        axis_index: Per-sheet row/column index maps, A sheet name keyed.

    Returns:
        Canonicalized JSON string.
    """
    try:
        dests = json.loads(destinations_json)
    except (json.JSONDecodeError, TypeError):
        return destinations_json

    if not isinstance(dests, list):
        return destinations_json

    rewritten = []
    for d in dests:
        if isinstance(d, str):
            if axis_index:
                d = _move_address_cells(d, axis_index)
            rewritten.append(_rewrite_address_string(d, sheet_map, {}, {}))
        else:
            rewritten.append(d)

    return json.dumps(rewritten, sort_keys=True, separators=(",", ":"))
