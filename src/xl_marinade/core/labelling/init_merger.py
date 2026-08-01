# ABOUTME: Detects and merges init+propagation patterns on projection sheets.
# ABOUTME: Combines split bindings like J7 (init) + J8:J607 (propagation) into one.

import json
import sqlite3
from dataclasses import dataclass
from typing import Any

from xl_marinade.core.bindings import Binding


@dataclass
class InitPattern:
    """A detected init + propagation pattern."""

    sheet_name: str
    column: int  # 1-based column number (for vertical) or row number (for horizontal)
    orientation: str  # "vertical" (column-based) or "horizontal" (row-based)
    init_binding_id: str  # Binding containing init cell(s)
    propagation_binding_id: str  # Binding containing propagation cells
    init_row_start: int  # First row of init (vertical) or first col (horizontal)
    init_row_end: int  # Last row of init (vertical) or last col (horizontal)
    propagation_row_start: int  # First row of propagation (vertical) or first col (horizontal)
    propagation_row_end: int  # Last row of propagation (vertical) or last col (horizontal)
    init_formula: str | None  # Formula in init cell (None if constant/value)
    propagation_formula: str  # Formula in propagation cells
    combined_label: str  # Label for merged variable


@dataclass
class CellInfo:
    """Information about a single cell within a binding.

    Used for splitting multi-cell bindings into individual cell references.

    Attributes:
        address: Full cell address (e.g., "Projection!M7")
        sheet: Sheet name (e.g., "Projection")
        column: Column letter (e.g., "M")
        row: Row number (e.g., 7)
        value: Cell value snapshot from IR
    """

    address: str
    sheet: str
    column: str
    row: int
    value: Any


@dataclass
class MultiInitCandidate:
    """Candidate for multi-cell init split-merge.

    Represents a pairing of one cell from a multi-cell constant binding
    with a calculation binding that should be merged.

    Attributes:
        init_binding_id: ID of the multi-cell constant binding
        init_cell: Specific cell from the init binding
        propagation_binding_id: ID of the calculation binding to merge with
        propagation_binding: The full propagation Binding object
        confidence: Confidence score (0.0-1.0) based on heuristics
        reason: Human-readable explanation of why this is a candidate
    """

    init_binding_id: str
    init_cell: CellInfo
    propagation_binding_id: str
    propagation_binding: "Binding"  # Forward reference
    confidence: float
    reason: str


def _parse_column_letter(address_a1: str) -> str | None:
    """Extract column letter from A1 address.

    Args:
        address_a1: Address like "J7", "J7:J607", or "Sheet!J7"

    Returns:
        Column letter(s) like "J" or None if parsing fails
    """
    # Strip sheet name if present
    if "!" in address_a1:
        address_a1 = address_a1.split("!")[-1]

    # Handle range notation (take first cell)
    if ":" in address_a1:
        address_a1 = address_a1.split(":")[0]

    # Extract column letter (everything before digits)
    col_letter = ""
    for char in address_a1:
        if char.isalpha():
            col_letter += char
        else:
            break

    return col_letter if col_letter else None


def _col_letter_to_number(col_letter: str) -> int:
    """Convert column letter to 1-based column number.

    Args:
        col_letter: Column letter like "A", "Z", "AA"

    Returns:
        1-based column number (A=1, B=2, Z=26, AA=27)
    """
    result = 0
    for char in col_letter.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def _parse_row_from_address(address: str) -> int:
    """Extract row number from cell address.

    Args:
        address: Address like "J7" or "Sheet1!J7"

    Returns:
        Row number (1-based)
    """
    # Strip sheet name if present
    if "!" in address:
        address = address.split("!")[-1]

    # Extract digits
    row_str = ""
    for char in address:
        if char.isdigit():
            row_str += char

    return int(row_str) if row_str else 0


def _get_binding_column(binding: Binding) -> int | None:
    """Get the column number for a binding (if it's a single column).

    Args:
        binding: Binding object

    Returns:
        1-based column number or None if not a single column
    """
    if binding.shape_cols != 1:
        return None

    col_letter = _parse_column_letter(binding.address_a1)
    if not col_letter:
        return None

    return _col_letter_to_number(col_letter)


def _get_binding_row_range(binding: Binding) -> tuple[int, int]:
    """Get start and end row numbers for a binding.

    Args:
        binding: Binding object

    Returns:
        Tuple of (start_row, end_row) both 1-based
    """
    start_row = _parse_row_from_address(binding.top_left_a1)
    end_row = start_row + binding.shape_rows - 1
    return (start_row, end_row)


def _get_binding_col_range(binding: Binding) -> tuple[int, int]:
    """Get start and end column numbers for a binding.

    Args:
        binding: Binding object

    Returns:
        Tuple of (start_col, end_col) both 1-based
    """
    col_letter = _parse_column_letter(binding.top_left_a1)
    if not col_letter:
        return (0, 0)
    start_col = _col_letter_to_number(col_letter)
    end_col = start_col + binding.shape_cols - 1
    return (start_col, end_col)


def _get_binding_row(binding: Binding) -> int | None:
    """Get the row number for a binding (if it's a single row).

    Args:
        binding: Binding object

    Returns:
        1-based row number or None if not a single row
    """
    if binding.shape_rows != 1:
        return None
    return _parse_row_from_address(binding.top_left_a1)


def _number_to_col_letter(col: int) -> str:
    """Convert 1-based column number to letter(s).

    Args:
        col: 1-based column number (A=1, B=2, Z=26, AA=27)

    Returns:
        Column letter(s) like "A", "Z", "AA"
    """
    result = ""
    while col > 0:
        col -= 1
        result = chr(col % 26 + ord("A")) + result
        col //= 26
    return result


def _get_formula_for_cell(
    ir_db_path: str,
    sheet: str,
    address: str,
    ir_db_conn: sqlite3.Connection | None = None,
    cell_formulas: dict[str, str] | None = None,
) -> str | None:
    """Get formula for a specific cell from IR database.

    Args:
        ir_db_path: Path to IR database
        sheet: Sheet name
        address: Cell address (e.g., "J7")
        ir_db_conn: Optional pre-existing connection (avoids opening a new one)
        cell_formulas: Optional pre-loaded formula dict (avoids DB queries entirely)

    Returns:
        Formula string or None if no formula
    """
    # Fast path: use pre-loaded formula dict
    if cell_formulas is not None:
        from xl_marinade.core.ref_converter import quote_sheet_name

        full_addr = f"{quote_sheet_name(sheet)}!{address}"
        formula = cell_formulas.get(full_addr, "")
        return formula if formula else None

    conn = ir_db_conn or sqlite3.connect(ir_db_path)
    try:
        cursor = conn.cursor()

        # Fast schema preferred: fetch per-cell R1C1 formula.
        # This avoids formula_a1_example drift where representative A1 formulas can
        # come from a different row than the requested cell.
        try:
            cursor.execute(
                """
                SELECT f.formula_r1c1
                FROM cells c
                JOIN sheets s ON c.sheet_id = s.sheet_id
                JOIN formulas f ON c.formula_id = f.formula_id
                WHERE s.sheet_name = ? AND c.a1 = ?
                LIMIT 1
                """,
                (sheet, address),
            )
            row = cursor.fetchone()
            if row and row[0]:
                return str(row[0])
        except sqlite3.Error:
            pass

        # Prefer fast schema view if present.
        full_address = f"{sheet}!{address}"
        try:
            cursor.execute(
                "SELECT formula FROM agent_cells WHERE cell_address = ?",
                (full_address,),
            )
            row = cursor.fetchone()
            if row and row[0]:
                return str(row[0])
        except sqlite3.Error:
            pass

        # Legacy/test schema: cells(cell_address_a1, formula_a1) (no sheet stored).
        try:
            cursor.execute(
                "SELECT formula_a1 FROM cells WHERE cell_address_a1 = ?",
                (full_address,),
            )
            row = cursor.fetchone()
            if row and row[0]:
                return str(row[0])
        except sqlite3.Error:
            pass

        # Some legacy schemas store addresses without sheet prefix.
        try:
            cursor.execute(
                "SELECT formula_a1 FROM cells WHERE cell_address_a1 = ?",
                (address,),
            )
            row = cursor.fetchone()
            if row and row[0]:
                return str(row[0])
        except sqlite3.Error:
            pass

        return None
    finally:
        if ir_db_conn is None:
            conn.close()


def _extract_column_header_from_ir(
    binding_id: str,
    ir_db_path: str,
    ir_db_conn: sqlite3.Connection | None = None,
) -> str | None:
    """Extract column header from binding's label_candidates.

    Looks for 'scan_above' type candidates which represent column headers.

    Args:
        binding_id: Binding ID to look up
        ir_db_path: Path to IR database
        ir_db_conn: Optional pre-existing connection (avoids opening a new one)

    Returns:
        Column header text if found, None otherwise
    """
    conn = ir_db_conn or sqlite3.connect(ir_db_path)
    try:
        cursor = conn.cursor()
        row = None
        try:
            cursor.execute(
                "SELECT spatial_candidates FROM agent_bindings WHERE binding_id = ?",
                (binding_id,),
            )
            row = cursor.fetchone()
        except sqlite3.Error:
            row = None

        if not row:
            try:
                cursor.execute(
                    "SELECT label_candidates_json FROM bindings WHERE binding_id = ?",
                    (binding_id,),
                )
                row = cursor.fetchone()
            except sqlite3.Error:
                row = None

        if not row or not row[0]:
            return None

        try:
            candidates_data = json.loads(row[0])
            # Handle case where JSON is a dict with 'label_candidates' key
            if isinstance(candidates_data, dict):
                candidates = candidates_data.get("label_candidates", [])
            elif isinstance(candidates_data, list):
                # Backwards compatibility: if it's already a list
                candidates = candidates_data
            else:
                return None
        except (json.JSONDecodeError, TypeError):
            return None

        # Find first scan_above candidate (column header)
        for candidate in candidates:
            if isinstance(candidate, dict) and candidate.get("type") == "scan_above":
                # Try 'text' field first (old format), then 'literals' (new format)
                header = candidate.get("text")
                if not header:
                    # New format: literals is a list of strings
                    literals = candidate.get("literals", [])
                    if literals and isinstance(literals, list):
                        # Pick first non-empty literal
                        header = next((lit for lit in literals if lit and lit.strip()), None)

                if header:
                    return str(header).strip()

        return None
    finally:
        if ir_db_conn is None:
            conn.close()


def _get_best_label_for_pattern(
    init_binding_id: str,
    propagation_binding_id: str,
    col_letter: str,
    sheet_name: str,
    ir_db_path: str,
    overlay_db_path: str | None,
    ir_db_conn: sqlite3.Connection | None = None,
) -> str:
    """Select best label for init pattern with intelligent fallback chain.

    Priority order:
    1. Propagation binding label from overlay DB (if overlay available)
    2. Init binding label from overlay DB (if overlay available)
    3. Column header from IR label_candidates (propagation first, then init)
    4. Synthetic name: time_series_{col_letter}

    Args:
        init_binding_id: Init binding ID
        propagation_binding_id: Propagation binding ID
        col_letter: Column letter (e.g., 'D')
        sheet_name: Sheet name (for logging/debugging)
        ir_db_path: Path to IR database
        overlay_db_path: Path to overlay database (None if not available)

    Returns:
        Best label for the merged time-series variable
    """
    # Priority 1 & 2: Try overlay DB labels
    if overlay_db_path:
        conn = None
        try:
            conn = sqlite3.connect(overlay_db_path)

            # Try propagation binding first
            cursor = conn.execute(
                "SELECT label FROM semantic_variables WHERE binding_id = ?",
                (propagation_binding_id,),
            )
            row = cursor.fetchone()
            if row and row[0]:
                return str(row[0])

            # Fallback to init binding
            cursor = conn.execute(
                "SELECT label FROM semantic_variables WHERE binding_id = ?", (init_binding_id,)
            )
            row = cursor.fetchone()
            if row and row[0]:
                return str(row[0])

        except sqlite3.Error:
            # Overlay DB error - continue to IR fallback
            pass
        finally:
            if conn:
                conn.close()

    # Priority 3: Try column header from IR label_candidates
    # Try propagation binding first
    header = _extract_column_header_from_ir(
        propagation_binding_id, ir_db_path, ir_db_conn=ir_db_conn
    )
    if header:
        return header

    # Fallback to init binding
    header = _extract_column_header_from_ir(init_binding_id, ir_db_path, ir_db_conn=ir_db_conn)
    if header:
        return header

    # Priority 4: Synthetic name (last resort)
    return f"time_series_{col_letter}"


def _propagation_references_init(
    init_binding: Binding,
    prop_binding: Binding,
    ir_db_path: str,
    orientation: str = "vertical",
    ir_db_conn: sqlite3.Connection | None = None,
    cell_formulas: dict[str, str] | None = None,
) -> bool:
    """Check if propagation formula references the init cell.

    For vertical (column-based): checks if prop formula references the init row.
    For horizontal (row-based): checks if prop formula references the init column.

    Checks A1-style refs, relative R1C1 refs, and absolute R1C1 refs.

    Args:
        init_binding: Init binding
        prop_binding: Propagation binding
        ir_db_path: Path to IR database
        orientation: "vertical" or "horizontal"

    Returns:
        True if propagation references init
    """
    # Get first cell of propagation
    prop_start_row = _parse_row_from_address(prop_binding.top_left_a1)
    col_letter = _parse_column_letter(prop_binding.address_a1)

    if not col_letter:
        return False

    prop_first_cell = f"{col_letter}{prop_start_row}"

    # Get formula for that cell
    formula = _get_formula_for_cell(
        ir_db_path,
        prop_binding.sheet,
        prop_first_cell,
        ir_db_conn=ir_db_conn,
        cell_formulas=cell_formulas,
    )

    if not formula:
        return False

    formula_upper = formula.upper()

    if orientation == "vertical":
        # Init end row is the last row of the init binding
        init_end_row = (
            _parse_row_from_address(init_binding.top_left_a1) + init_binding.shape_rows - 1
        )

        # Check A1-style reference (e.g., "J7" in formula)
        init_ref = f"{col_letter}{init_end_row}"
        if init_ref in formula_upper:
            return True

        # Check absolute R1C1 reference (e.g., "R7C" in R7C10)
        abs_pattern = f"R{init_end_row}C"
        if abs_pattern in formula_upper:
            return True

        # Check relative R1C1 reference (e.g., "R[-1]C")
        offset = prop_start_row - init_end_row
        if f"R[-{offset}]C" in formula_upper:
            return True

        # Adjacent check
        return bool(offset == 1 and "R[-1]C" in formula_upper)

    else:  # horizontal
        # Init end column is the last column of the init binding
        init_col_letter = _parse_column_letter(init_binding.top_left_a1)
        if not init_col_letter:
            return False
        init_end_col = _col_letter_to_number(init_col_letter) + init_binding.shape_cols - 1
        init_end_col_letter = _number_to_col_letter(init_end_col)

        prop_start_col = _col_letter_to_number(col_letter)

        # Check A1-style reference (e.g., "D7" where D is init col)
        init_ref = f"{init_end_col_letter}{prop_start_row}"
        if init_ref in formula_upper:
            return True

        # Check absolute R1C1 reference (e.g., "RC4" for column 4)
        abs_pattern = f"RC{init_end_col}"
        if abs_pattern in formula_upper:
            return True

        # Check relative R1C1 reference (e.g., "RC[-1]")
        offset = prop_start_col - init_end_col
        if f"RC[-{offset}]" in formula_upper:
            return True

        # Adjacent check
        return bool(offset == 1 and "RC[-1]" in formula_upper)


def _find_init_propagation_pair(
    bindings: list[Binding],
    ir_db_path: str,
    overlay_db_path: str | None = None,
    orientation: str = "vertical",
    ir_db_conn: sqlite3.Connection | None = None,
    cell_formulas: dict[str, str] | None = None,
) -> InitPattern | None:
    """Find init+propagation pair in a list of same-axis bindings.

    Works for both vertical (column-based) and horizontal (row-based) arrays.

    Args:
        bindings: List of bindings on the same axis (same column or same row)
        ir_db_path: Path to IR database
        overlay_db_path: Path to overlay database (optional)
        orientation: "vertical" (column-based) or "horizontal" (row-based)

    Returns:
        InitPattern if found, None otherwise
    """
    if len(bindings) < 2:
        return None

    if orientation == "vertical":
        # Sort by start row; size = shape_rows
        sorted_bindings = sorted(bindings, key=lambda b: _get_binding_row_range(b)[0])
    else:
        # Sort by start column; size = shape_cols
        sorted_bindings = sorted(bindings, key=lambda b: _get_binding_col_range(b)[0])

    # Check consecutive pairs
    for i in range(len(sorted_bindings) - 1):
        init_candidate = sorted_bindings[i]
        prop_candidate = sorted_bindings[i + 1]

        if orientation == "vertical":
            init_size = init_candidate.shape_rows
            prop_size = prop_candidate.shape_rows
            init_start, init_end = _get_binding_row_range(init_candidate)
            prop_start, prop_end = _get_binding_row_range(prop_candidate)
        else:
            init_size = init_candidate.shape_cols
            prop_size = prop_candidate.shape_cols
            init_start, init_end = _get_binding_col_range(init_candidate)
            prop_start, prop_end = _get_binding_col_range(prop_candidate)

        # Check size relationship (init should be smaller)
        if init_size >= prop_size:
            continue

        # Must be adjacent or with small gap (max 1 cell)
        if prop_start - init_end > 2:
            continue

        # Check if propagation references init cell
        if not _propagation_references_init(
            init_candidate,
            prop_candidate,
            ir_db_path,
            orientation,
            ir_db_conn=ir_db_conn,
            cell_formulas=cell_formulas,
        ):
            continue

        # Get formulas — use the END cell of init (not start, which may be a header)
        col_letter = _parse_column_letter(init_candidate.address_a1)
        if not col_letter:
            continue

        if orientation == "vertical":
            init_end_row = (
                _parse_row_from_address(init_candidate.top_left_a1) + init_candidate.shape_rows - 1
            )
            init_cell = f"{col_letter}{init_end_row}"
            prop_first_cell = f"{col_letter}{prop_start}"
        else:
            init_end_col_letter = _number_to_col_letter(init_end)
            init_row = _parse_row_from_address(init_candidate.top_left_a1)
            init_cell = f"{init_end_col_letter}{init_row}"
            prop_col_letter = _number_to_col_letter(prop_start)
            prop_first_cell = f"{prop_col_letter}{init_row}"

        init_formula = _get_formula_for_cell(
            ir_db_path,
            init_candidate.sheet,
            init_cell,
            ir_db_conn=ir_db_conn,
            cell_formulas=cell_formulas,
        )
        prop_formula = _get_formula_for_cell(
            ir_db_path,
            prop_candidate.sheet,
            prop_first_cell,
            ir_db_conn=ir_db_conn,
            cell_formulas=cell_formulas,
        )

        # Only propagation must have a formula (init can be a constant)
        if not prop_formula:
            continue

        # Create pattern - use intelligent label selection
        combined_label = _get_best_label_for_pattern(
            init_binding_id=init_candidate.binding_id,
            propagation_binding_id=prop_candidate.binding_id,
            col_letter=col_letter,
            sheet_name=init_candidate.sheet,
            ir_db_path=ir_db_path,
            overlay_db_path=overlay_db_path,
            ir_db_conn=ir_db_conn,
        )

        return InitPattern(
            sheet_name=init_candidate.sheet,
            column=_col_letter_to_number(col_letter)
            if orientation == "vertical"
            else _parse_row_from_address(init_candidate.top_left_a1),
            orientation=orientation,
            init_binding_id=init_candidate.binding_id,
            propagation_binding_id=prop_candidate.binding_id,
            init_row_start=init_start,
            init_row_end=init_end,
            propagation_row_start=prop_start,
            propagation_row_end=prop_end,
            init_formula=init_formula,
            propagation_formula=prop_formula,
            combined_label=combined_label,
        )

    return None


def detect_init_patterns(
    bindings: list[Binding],
    ir_db_path: str,
    overlay_db_path: str | None = None,
    ir_db_conn: sqlite3.Connection | None = None,
    cell_formulas: dict[str, str] | None = None,
) -> list[InitPattern]:
    """Detect init + propagation patterns in bindings.

    Looks for pairs of bindings on any sheet where:
    1. Both are on the same axis (same column for vertical, same row for horizontal)
    2. One is a small range (init) at the start
    3. One is a larger range (propagation) continuing after
    4. Propagation formula references the init cell

    Works for both column-based (vertical) and row-based (horizontal) arrays.

    Args:
        bindings: All bindings from the model
        ir_db_path: Path to IR database for formula lookup
        overlay_db_path: Path to overlay database for label lookup (optional)

    Returns:
        List of detected InitPattern objects
    """
    patterns = []

    # --- Vertical (column-based): group by (sheet, column) ---
    by_sheet_col: dict[tuple[str, int], list[Binding]] = {}

    for binding in bindings:
        # Only check 1D column bindings
        if binding.shape_cols != 1:
            continue

        col = _get_binding_column(binding)
        if col is None:
            continue

        key = (binding.sheet, col)
        if key not in by_sheet_col:
            by_sheet_col[key] = []
        by_sheet_col[key].append(binding)

    for (_sheet, _col), col_bindings in by_sheet_col.items():
        pattern = _find_init_propagation_pair(
            col_bindings,
            ir_db_path,
            overlay_db_path,
            orientation="vertical",
            ir_db_conn=ir_db_conn,
            cell_formulas=cell_formulas,
        )
        if pattern:
            patterns.append(pattern)

    # --- Horizontal (row-based): group by (sheet, row) ---
    by_sheet_row: dict[tuple[str, int], list[Binding]] = {}

    for binding in bindings:
        # Only check 1D row bindings
        if binding.shape_rows != 1:
            continue

        row = _get_binding_row(binding)
        if row is None:
            continue

        key = (binding.sheet, row)
        if key not in by_sheet_row:
            by_sheet_row[key] = []
        by_sheet_row[key].append(binding)

    for (_sheet, _row), row_bindings in by_sheet_row.items():
        pattern = _find_init_propagation_pair(
            row_bindings,
            ir_db_path,
            overlay_db_path,
            orientation="horizontal",
            ir_db_conn=ir_db_conn,
            cell_formulas=cell_formulas,
        )
        if pattern:
            patterns.append(pattern)

    return patterns


def create_init_merge_mutations(patterns: list[InitPattern]) -> list[dict[str, Any]]:
    """Create merge mutations for detected init patterns.

    Generates mutations that will merge init and propagation bindings
    into a single variable with initiation metadata.

    Args:
        patterns: Detected init patterns

    Returns:
        List of mutation dictionaries ready for MutationLogger
    """
    mutations = []

    for pattern in patterns:
        if pattern.orientation == "vertical":
            axis_label = "row"
            reasoning = (
                f"Init pattern merge: row {pattern.init_row_start} + rows "
                f"{pattern.propagation_row_start}-{pattern.propagation_row_end} "
                f"into time-series variable"
            )
        else:
            axis_label = "col"
            reasoning = (
                f"Init pattern merge: col {pattern.init_row_start} + cols "
                f"{pattern.propagation_row_start}-{pattern.propagation_row_end} "
                f"into time-series variable (horizontal)"
            )

        mutation = {
            "action": "merge_bindings",
            "parameters": {
                "source_binding_ids": [pattern.init_binding_id, pattern.propagation_binding_id],
                "new_binding_id": f"merged_{pattern.propagation_binding_id[:16]}",
                "label": pattern.combined_label,
                "metadata": {
                    "entity_type": "time_series",
                    "orientation": pattern.orientation,
                    "initiation_formula": pattern.init_formula,
                    "propagation_formula": pattern.propagation_formula,
                    f"init_{axis_label}": pattern.init_row_start,
                    "actuarial_class": "Calculation",
                    "actuarial_class_confidence": 0.90,
                },
            },
            "reasoning": reasoning,
        }
        mutations.append(mutation)

    return mutations


# Story 37: Multi-cell init split-merge helpers


def expand_binding_to_cells(binding: Binding, ir_conn: sqlite3.Connection) -> list[CellInfo]:
    """Expand a multi-cell binding into individual CellInfo objects.

    Queries the IR database to get all cells belonging to this binding,
    including their values and addresses. Falls back to using the binding's
    own cells list with direct value lookups when cell_to_binding is empty
    (common for constant bindings created during grouping in the fast pipeline).

    Args:
        binding: The binding to expand
        ir_conn: Open connection to IR database

    Returns:
        List of CellInfo objects, one per cell in the binding.
        Empty list if binding has no cells or query fails.
    """
    import re

    cells: list[CellInfo] = []

    # Extract sheet name from binding
    sheet = binding.sheet

    # Query cells from IR database (fast schema preferred, legacy/test fallback).
    rows: list[tuple] = []
    try:
        cursor = ir_conn.execute(
            """
            SELECT ac.cell_address, ac.value
            FROM cell_to_binding ctb
            JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
            WHERE ctb.binding_id = ?
            ORDER BY ac.cell_address
            """,
            (binding.binding_id,),
        )
        rows = cursor.fetchall()
    except sqlite3.Error:
        try:
            cursor = ir_conn.execute(
                """
                SELECT cell_address_a1, value_snapshot
                FROM cells
                WHERE binding_id = ?
                ORDER BY cell_address_a1
                """,
                (binding.binding_id,),
            )
            rows = cursor.fetchall()
        except sqlite3.Error:
            pass

    # Fallback: if database returned nothing, use binding.cells and look up
    # values directly from the cells table by (sheet_id, a1).
    if not rows and hasattr(binding, "cells") and binding.cells:
        rows = _expand_cells_from_binding_attrs(binding, ir_conn)

    for row in rows:
        cell_addr: str = row[0]
        value = row[1]

        # Extract column and row from address (handles Sheet!A1 and A1 formats)
        match = re.match(r"(?:.*!)?([A-Z]+)(\d+)$", cell_addr)
        if match:
            col_letter = match.group(1)
            row_num = int(match.group(2))

            cells.append(
                CellInfo(
                    address=cell_addr, sheet=sheet, column=col_letter, row=row_num, value=value
                )
            )

    return cells


def _expand_cells_from_binding_attrs(
    binding: Binding, ir_conn: sqlite3.Connection
) -> list[tuple[str, Any]]:
    """Fallback: expand cells using binding.cells and direct DB value lookup.

    Used when cell_to_binding doesn't have entries for this binding (e.g.,
    constant bindings created during grouping in the fast pipeline).
    """

    # Get sheet_id for value lookups
    sheet_id = None
    try:
        cursor = ir_conn.execute(
            "SELECT sheet_id FROM sheets WHERE sheet_name = ?", (binding.sheet,)
        )
        row = cursor.fetchone()
        if row:
            sheet_id = row[0]
    except sqlite3.Error:
        return []

    if sheet_id is None:
        return []

    if not binding.cells:
        return []

    # Pair each binding cell with its sheet-stripped A1 form for the IN-list query.
    pairs: list[tuple[str, str]] = []
    for cell_addr in binding.cells:
        a1 = cell_addr.split("!")[-1] if "!" in cell_addr else cell_addr
        pairs.append((cell_addr, a1))

    # Single batched query keyed on (sheet_id, a1 IN (...)) — replaces the
    # per-cell roundtrip pattern that was 99% of init-merger wall time on LH
    # (9,500 fetchones × 51 ms each = 1,050s).  SQLite has a default parameter
    # limit (SQLITE_MAX_VARIABLE_NUMBER) of 999, so chunk to stay well under.
    a1_to_value: dict[str, Any] = {}
    CHUNK = 500
    for chunk_start in range(0, len(pairs), CHUNK):
        chunk = pairs[chunk_start : chunk_start + CHUNK]
        a1s = [p[1] for p in chunk]
        placeholders = ",".join("?" * len(a1s))
        try:
            cursor = ir_conn.execute(
                f"""
                SELECT c.a1, jv.json
                FROM cells c
                LEFT JOIN json_blobs jv ON c.value_blob_id = jv.blob_id
                WHERE c.sheet_id = ? AND c.a1 IN ({placeholders})
                """,
                (sheet_id, *a1s),
            )
            for row_a1, row_value in cursor.fetchall():
                a1_to_value[row_a1] = row_value
        except sqlite3.Error:
            # Match legacy behaviour: silently fall through; missing cells
            # become value=None below.
            pass

    return [(cell_addr, a1_to_value.get(a1)) for cell_addr, a1 in pairs]


def formula_references_cell(formula: str, cell_column: str, cell_row: int) -> bool:
    """Check if a formula references a specific cell.

    Handles all Excel reference styles: A1, $A$1, A$1, $A1.
    Uses word boundaries to avoid false positives (e.g., AM7 != M7).

    Args:
        formula: The formula string (e.g., "=M7+1")
        cell_column: Column letter (e.g., "M")
        cell_row: Row number (e.g., 7)

    Returns:
        True if formula references the cell, False otherwise.
    """
    import re

    if not formula or not formula.startswith("="):
        return False

    # Build patterns to match all reference styles
    # Use word boundary \b to avoid matching AM7 when looking for M7
    col = cell_column.upper()
    row = str(cell_row)

    patterns = [
        rf"(?<![A-Z]){col}{row}(?!\d)",  # M7 (not AM7, not M70)
        rf"\${col}\${row}(?!\d)",  # $M$7
        rf"(?<![A-Z]){col}\${row}(?!\d)",  # M$7
        rf"\${col}{row}(?!\d)",  # $M7
    ]

    return any(re.search(pattern, formula, re.IGNORECASE) for pattern in patterns)


def values_are_coherent(init_value: Any, propagation_sample: list[Any]) -> float:
    """Check if init value is coherent with propagation values.

    Returns a score between 0.0 and 1.0 indicating how well the init value
    fits as an initialization for the propagation binding.

    Scoring:
    - Zero init with numeric propagation: 1.0 (very common for accumulators)
    - Non-zero numeric init with numeric propagation: 0.8
    - Text init with numeric propagation: 0.0 (likely a header, not init)
    - Both text: 0.3 (possible but low confidence)
    - Other: 0.5

    Args:
        init_value: The initialization value
        propagation_sample: Sample of propagation values (first 10 or so)

    Returns:
        Coherence score (0.0-1.0)
    """
    if init_value is None or not propagation_sample:
        return 0.0

    # Helper to check if value is numeric
    # NOTE: Values from the IR database are JSON-encoded, so numeric values
    # may be stored as '"0"' or '"243.5"' (with JSON string quotes).
    def _strip_json_quotes(v: Any) -> Any:
        if isinstance(v, str) and len(v) >= 2 and v.startswith('"') and v.endswith('"'):
            return v[1:-1]
        return v

    def is_numeric(v: Any) -> bool:
        if isinstance(v, int | float):
            return True
        if isinstance(v, str):
            try:
                float(_strip_json_quotes(v).replace(",", ""))
                return True
            except (ValueError, AttributeError):
                return False
        return False

    init_is_numeric = is_numeric(init_value)

    # Check if propagation values are numeric (sample first 10)
    prop_sample = [v for v in propagation_sample[:10] if v is not None]
    if not prop_sample:
        return 0.0

    prop_is_numeric = all(is_numeric(v) for v in prop_sample)

    # Scoring logic
    if init_is_numeric and prop_is_numeric:
        try:
            if float(str(_strip_json_quotes(init_value)).replace(",", "")) == 0:
                return 1.0  # Zero init is very common
        except (ValueError, TypeError):
            pass
        return 0.8

    # Handle case where propagation values are formulas (not evaluated)
    # This happens when values are not cached in the workbook
    def is_formula_string(v: Any) -> bool:
        s = str(v).strip("\"'")
        return s.startswith("=")

    prop_has_formulas = any(is_formula_string(v) for v in prop_sample)
    if init_is_numeric and prop_has_formulas:
        try:
            if float(str(_strip_json_quotes(init_value)).replace(",", "")) == 0:
                return 1.0  # Zero init feeding formula = very likely valid
        except (ValueError, TypeError):
            pass
        return 0.7  # Numeric init feeding formula

    if not init_is_numeric and prop_is_numeric:
        return 0.0  # Text init with numeric prop = likely header

    if not init_is_numeric and not prop_is_numeric:
        return 0.3  # Both text, low confidence

    return 0.5


def is_column_adjacent(col_a: str, col_b: str) -> bool:
    """Check if column B immediately follows column A.

    Examples: A→B (True), B→C (True), A→C (False), Z→AA (True)

    Args:
        col_a: First column letter (e.g., "A")
        col_b: Second column letter (e.g., "B")

    Returns:
        True if col_b is the immediate successor of col_a
    """

    def col_to_num(col: str) -> int:
        """Convert column letter(s) to number (A=1, Z=26, AA=27)."""
        result = 0
        for char in col.upper():
            result = result * 26 + (ord(char) - ord("A") + 1)
        return result

    return col_to_num(col_b) == col_to_num(col_a) + 1


def detect_multi_init_patterns(
    bindings: list[Binding],
    ir_db_path: str,
    overlay_db_path: str | None = None,
    ir_db_conn: sqlite3.Connection | None = None,
) -> list[MultiInitCandidate]:
    """Detect multi-cell init patterns that should be split.

    Looks for constant bindings with multiple cells that border
    multiple calculation bindings, where each init cell should
    merge with its corresponding calculation binding.

    Heuristics (confidence scoring):
    - Reference Pattern: +0.40 (calculation formula references init cell)
    - Spatial Adjacency: +0.20 (init cell directly borders propagation)
    - Value Coherence: +0.00-0.20 (init value pattern makes sense)
    - Constant Classification: +0.10 (init is constant, not formula)
    - Actuarial Context: +0.10 (sheet suggests projection/time-series)

    Threshold: >= 0.60 confidence required for candidate

    Args:
        bindings: All bindings from the model
        ir_db_path: Path to IR database
        overlay_db_path: Path to overlay database (optional, for future use)

    Returns:
        List of MultiInitCandidate objects for each cell-propagation pair
    """
    import re

    candidates: list[MultiInitCandidate] = []

    conn = ir_db_conn or sqlite3.connect(ir_db_path)
    _close_conn = ir_db_conn is None

    try:
        for binding in bindings:
            # Skip single-cell bindings (handled by existing Story 31 merger)
            if binding.shape_rows == 1 and binding.shape_cols == 1:
                continue

            # Skip formula bindings (only constants can be inits)
            if binding.binding_type != "constant":
                continue

            # Determine orientation
            if binding.shape_rows == 1 and binding.shape_cols > 1:
                orientation = "column"  # Row of constants above columns
            elif binding.shape_cols == 1 and binding.shape_rows > 1:
                orientation = "row"  # Column of constants left of rows
            else:
                continue  # Matrix - out of scope for this story

            # Expand binding to individual cells
            cells = expand_binding_to_cells(binding, conn)
            if not cells:
                continue

            # Find candidate propagation bindings
            for cell in cells:
                for prop_binding in bindings:
                    # Skip same binding
                    if prop_binding.binding_id == binding.binding_id:
                        continue

                    # Must be on same sheet
                    if prop_binding.sheet != cell.sheet:
                        continue

                    # Must have formula (is a calculation)
                    if prop_binding.binding_type == "constant":
                        continue

                    # Propagation must be a multi-cell binding (column or row)
                    # Single-cell formulas are handled by Story 31's 1-to-1 init merger
                    if prop_binding.shape_rows == 1 and prop_binding.shape_cols == 1:
                        continue

                    # Parse propagation binding's top-left cell
                    prop_match = re.match(r"(?:.*!)?([A-Z]+)(\d+)$", prop_binding.top_left_a1)
                    if not prop_match:
                        continue

                    prop_col = prop_match.group(1)
                    prop_row = int(prop_match.group(2))

                    # Check spatial adjacency based on orientation
                    is_adjacent = False
                    if orientation == "column":
                        # Init should be directly above propagation (same col, row-1)
                        is_adjacent = cell.column == prop_col and cell.row == prop_row - 1
                    elif orientation == "row":
                        # Init should be directly left of propagation (same row, col-1)
                        is_adjacent = cell.row == prop_row and is_column_adjacent(
                            cell.column, prop_col
                        )

                    if not is_adjacent:
                        continue  # Not adjacent, skip this pairing

                    # Start building confidence score
                    score = 0.0
                    reasons: list[str] = []

                    # Spatial adjacency: +0.20
                    score += 0.20
                    reasons.append("spatial_adjacency")

                    # Constant classification: +0.10
                    score += 0.10
                    reasons.append("is_constant")

                    # Reference pattern check: +0.40 (STRONGEST)
                    # Query formula directly from cells table (cell_to_binding
                    # may be empty during the fast pipeline snapshot).
                    prop_first_cell_a1 = f"{prop_col}{prop_row}"
                    formula_row = None
                    try:
                        cursor = conn.execute(
                            """
                            SELECT f.formula_r1c1
                            FROM cells c
                            JOIN sheets s ON c.sheet_id = s.sheet_id
                            JOIN formulas f ON c.formula_id = f.formula_id
                            WHERE s.sheet_name = ? AND c.a1 = ?
                            LIMIT 1
                            """,
                            (prop_binding.sheet, prop_first_cell_a1),
                        )
                        formula_row = cursor.fetchone()
                    except sqlite3.Error:
                        pass
                    if not formula_row:
                        try:
                            cursor = conn.execute(
                                """
                                SELECT ac.formula
                                FROM cell_to_binding ctb
                                JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
                                WHERE ctb.binding_id = ? AND ac.formula IS NOT NULL
                                LIMIT 1
                                """,
                                (prop_binding.binding_id,),
                            )
                            formula_row = cursor.fetchone()
                        except sqlite3.Error:
                            pass
                    if (
                        formula_row
                        and formula_row[0]
                        and formula_references_cell(formula_row[0], cell.column, cell.row)
                    ):
                        score += 0.40
                        reasons.append("formula_references_init")

                    # Value coherence check: +0.00-0.20
                    # Query values directly from cells table (cell_to_binding
                    # may be empty during the fast pipeline snapshot).
                    prop_values: list = []
                    try:
                        cursor = conn.execute(
                            """
                            SELECT jv.json
                            FROM cells c
                            JOIN sheets s ON c.sheet_id = s.sheet_id
                            JOIN json_blobs jv ON c.value_blob_id = jv.blob_id
                            WHERE s.sheet_name = ? AND c.col = ? AND c.row >= ? AND c.row <= ?
                            LIMIT 10
                            """,
                            (
                                prop_binding.sheet,
                                _col_letter_to_number(prop_col),
                                prop_row,
                                prop_row + 9,
                            ),
                        )
                        prop_values = [r[0] for r in cursor.fetchall()]
                    except sqlite3.Error:
                        pass
                    if not prop_values:
                        try:
                            cursor = conn.execute(
                                """
                                SELECT ac.value
                                FROM cell_to_binding ctb
                                JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
                                WHERE ctb.binding_id = ?
                                LIMIT 10
                                """,
                                (prop_binding.binding_id,),
                            )
                            prop_values = [r[0] for r in cursor.fetchall()]
                        except sqlite3.Error:
                            pass
                    coherence = values_are_coherent(cell.value, prop_values)
                    coherence_contribution = coherence * 0.20
                    score += coherence_contribution
                    if coherence > 0.5:
                        reasons.append("value_coherent")

                    # Actuarial context: +0.10 (sheet suggests projection)
                    if "projection" in cell.sheet.lower():
                        score += 0.10
                        reasons.append("projection_sheet")

                    # Threshold check: >= 0.60
                    if score >= 0.60:
                        candidates.append(
                            MultiInitCandidate(
                                init_binding_id=binding.binding_id,
                                init_cell=cell,
                                propagation_binding_id=prop_binding.binding_id,
                                propagation_binding=prop_binding,
                                confidence=round(score, 2),
                                reason=", ".join(reasons),
                            )
                        )

    finally:
        if _close_conn:
            conn.close()

    return candidates


def create_multi_init_merge_mutations(
    candidates: list[MultiInitCandidate],
    overlay_db_path: str | None = None,
    ir_db_path: str | None = None,
    ir_db_conn: sqlite3.Connection | None = None,
) -> list[dict]:
    """Create merge mutations for multi-cell init patterns.

    Creates:
    1. One merge_bindings mutation per cell-propagation pair (with cell_subset)
    2. One disable_binding mutation for each original multi-cell binding (at end)

    The cell_subset parameter allows partial merging: only the specified cell
    from the init binding is included, and the init binding stays active for
    subsequent partial merges. After all merges, disable_binding removes the
    original multi-cell binding.

    Args:
        candidates: List of MultiInitCandidate objects from detect_multi_init_patterns
        overlay_db_path: Path to overlay database for label lookup
        ir_db_path: Path to IR database for label fallback

    Returns:
        List of mutation dictionaries ready for MutationLogger
    """
    import hashlib

    mutations: list[dict] = []

    # Group candidates by original init binding (for disable mutation at end)
    by_init_binding: dict[str, list[MultiInitCandidate]] = {}
    for c in candidates:
        if c.init_binding_id not in by_init_binding:
            by_init_binding[c.init_binding_id] = []
        by_init_binding[c.init_binding_id].append(c)

    for init_binding_id, binding_candidates in by_init_binding.items():
        # 1. Create merge mutations for each cell-propagation pair with cell_subset
        for candidate in binding_candidates:
            # Get label using Story 35's label selection logic
            if ir_db_path:
                label = _get_best_label_for_pattern(
                    init_binding_id=init_binding_id,
                    propagation_binding_id=candidate.propagation_binding_id,
                    col_letter=candidate.init_cell.column,
                    sheet_name=candidate.init_cell.sheet,
                    ir_db_path=ir_db_path,
                    overlay_db_path=overlay_db_path,
                    ir_db_conn=ir_db_conn,
                )
            else:
                label = f"time_series_{candidate.init_cell.column}"

            # Generate unique binding ID for the merged entity
            hash_input = (
                f"multi_init_{init_binding_id}_"
                f"{candidate.init_cell.address}_"
                f"{candidate.propagation_binding_id}"
            )
            new_id = hashlib.sha256(hash_input.encode()).hexdigest()[:16]
            new_binding_id = f"time_series_{new_id}"

            mutations.append(
                {
                    "action": "merge_bindings",
                    "parameters": {
                        "source_binding_ids": [init_binding_id, candidate.propagation_binding_id],
                        "new_binding_id": new_binding_id,
                        "cell_subset": [candidate.init_cell.address],
                        "label": label,
                        "metadata": {
                            "entity_type": "time_series",
                            "actuarial_class": "Calculation",
                            "init_pattern": True,
                            "split_from": init_binding_id,
                            "confidence": candidate.confidence,
                        },
                    },
                    "reasoning": f"Multi-init partial merge: {candidate.reason}",
                }
            )

        # 2. Disable the original multi-cell binding (after all partial merges)
        cell_columns = sorted({c.init_cell.column for c in binding_candidates})
        mutations.append(
            {
                "action": "disable_binding",
                "parameters": {
                    "binding_id": init_binding_id,
                    "reason": f"Split into cell-level init patterns for columns {', '.join(cell_columns)}",
                },
                "reasoning": "Disabling original multi-cell binding after partial merges",
            }
        )

    return mutations


def detect_projection_init_singletons(
    bindings: list[Binding],
    ir_db_path: str,
    already_consumed: set[str] | None = None,
    min_projection_rows: int = 50,
    overlay_db_path: str | None = None,
    ir_db_conn: sqlite3.Connection | None = None,
    cell_formulas: dict[str, str] | None = None,
) -> list[InitPattern]:
    """Detect singleton init values directly above long projection columns.

    Unlike Story 31, this does NOT require the propagation formula to reference
    the init cell. It uses spatial adjacency + projection length + numeric value
    as the signal.

    Handles cases like Q7 (=1, discount factor at t=0) above Q8:Q607 where
    Q8's formula does not reference Q7.

    Criteria:
    - Init is a 1×1 binding (formula or constant) with a numeric value
    - Propagation is a single-column formula binding with many rows (>=50)
    - Init is directly above propagation (same column, adjacent row)
    - Both on the same sheet

    Args:
        bindings: All bindings from the model
        ir_db_path: Path to IR database
        already_consumed: Binding IDs already merged by earlier passes (skip these)
        min_projection_rows: Minimum rows for a binding to be considered a projection
        overlay_db_path: Path to overlay database for label lookup (optional)

    Returns:
        List of detected InitPattern objects
    """
    import re

    if already_consumed is None:
        already_consumed = set()

    patterns = []

    # Index projection columns: single-column formula bindings with many rows
    projection_map: dict[tuple[str, int, int], Binding] = {}
    for b in bindings:
        if b.binding_id in already_consumed:
            continue
        if b.binding_type != "formula" or b.shape_cols != 1:
            continue
        if b.shape_rows < min_projection_rows:
            continue
        match = re.match(r"(?:.*!)?([A-Z]+)(\d+)$", b.top_left_a1)
        if match:
            col_num = _col_letter_to_number(match.group(1))
            row_num = int(match.group(2))
            projection_map[(b.sheet, row_num, col_num)] = b

    if not projection_map:
        return patterns

    # Check singleton candidates
    conn = ir_db_conn or sqlite3.connect(ir_db_path)
    _close_conn = ir_db_conn is None
    try:
        for b in bindings:
            if b.binding_id in already_consumed:
                continue
            if b.shape_rows != 1 or b.shape_cols != 1:
                continue

            # Parse position
            match = re.match(r"(?:.*!)?([A-Z]+)(\d+)$", b.top_left_a1)
            if not match:
                continue
            col_letter = match.group(1)
            col_num = _col_letter_to_number(col_letter)
            row_num = int(match.group(2))

            # Check for projection directly below
            target = projection_map.get((b.sheet, row_num + 1, col_num))
            if not target or target.binding_id in already_consumed:
                continue

            # Check that init value is numeric (exclude text headers).
            # Query cells table directly (cell_to_binding may be empty
            # during the fast pipeline snapshot).
            val_row = None
            try:
                cursor = conn.execute(
                    """
                    SELECT jv.json
                    FROM cells c
                    JOIN sheets s ON c.sheet_id = s.sheet_id
                    LEFT JOIN json_blobs jv ON c.value_blob_id = jv.blob_id
                    WHERE s.sheet_name = ? AND c.a1 = ?
                    LIMIT 1
                    """,
                    (b.sheet, f"{col_letter}{row_num}"),
                )
                val_row = cursor.fetchone()
            except sqlite3.Error:
                pass

            if not val_row or val_row[0] is None:
                continue

            raw_val = val_row[0]
            # Strip JSON string quotes if present
            if (
                isinstance(raw_val, str)
                and len(raw_val) >= 2
                and raw_val.startswith('"')
                and raw_val.endswith('"')
            ):
                raw_val = raw_val[1:-1]
            try:
                float(str(raw_val).replace(",", ""))
            except (ValueError, TypeError):
                continue  # Not numeric, skip (likely a header label)

            # Get formulas for the pattern
            init_formula = _get_formula_for_cell(
                ir_db_path,
                b.sheet,
                f"{col_letter}{row_num}",
                ir_db_conn=conn,
                cell_formulas=cell_formulas,
            )
            prop_first_row = row_num + 1
            prop_formula = _get_formula_for_cell(
                ir_db_path,
                target.sheet,
                f"{col_letter}{prop_first_row}",
                ir_db_conn=conn,
                cell_formulas=cell_formulas,
            )
            if not prop_formula:
                continue

            # Get label
            combined_label = _get_best_label_for_pattern(
                init_binding_id=b.binding_id,
                propagation_binding_id=target.binding_id,
                col_letter=col_letter,
                sheet_name=b.sheet,
                ir_db_path=ir_db_path,
                overlay_db_path=overlay_db_path,
                ir_db_conn=conn,
            )

            _, prop_end = _get_binding_row_range(target)

            patterns.append(
                InitPattern(
                    sheet_name=b.sheet,
                    column=col_num,
                    orientation="vertical",
                    init_binding_id=b.binding_id,
                    propagation_binding_id=target.binding_id,
                    init_row_start=row_num,
                    init_row_end=row_num,
                    propagation_row_start=prop_first_row,
                    propagation_row_end=prop_end,
                    init_formula=init_formula,
                    propagation_formula=prop_formula,
                    combined_label=combined_label,
                )
            )

    finally:
        if _close_conn:
            conn.close()

    return patterns
