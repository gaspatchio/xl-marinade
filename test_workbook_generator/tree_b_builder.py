# ABOUTME: Builder for Tree B statistical analysis with VLOOKUP, INDEX, MATCH, and array formulas.
# ABOUTME: Creates lookup tables, statistical functions, and data analysis patterns for test workbook validation.


from openpyxl.worksheet.worksheet import Worksheet

from .workbook_builder import write_cell_formula, write_cell_value


def create_tree_b(ws: Worksheet, start_row: int = 7) -> list[str]:
    """
    Creates Tree B statistical analysis with VLOOKUP, INDEX, MATCH, and array formulas.

    Tree B contains:
    - Header row with section title
    - Column headers (Data, Lookup, Index, Match, Array, Result)
    - Three rows of statistical data (100, 200, 300)
    - Lookup table for VLOOKUP references
    - Formulas creating statistical analysis chain: data → lookups → statistical functions → results

    Args:
        ws: The worksheet to write to
        start_row: Starting row for Tree B (1-indexed, default: 7)

    Returns:
        List[str]: Cell addresses of all cells created in Tree B

    Example:
        >>> wb = create_workbook()
        >>> ws = wb.active
        >>> cells = create_tree_b(ws, start_row=7)
        >>> len(cells)  # Total cells (headers + data + lookup table + formulas)
        40
    """
    if start_row < 1:
        raise ValueError("start_row must be >= 1")

    created_cells = []

    # Row offsets from start_row
    header_row = start_row
    column_header_row = start_row + 1
    data_row_1 = start_row + 2  # 100
    data_row_3 = start_row + 4  # 300

    # Create lookup table starting after the main Tree B section
    lookup_table_start_row = start_row + 6

    # Create main header
    write_cell_value(ws, f"A{header_row}", "Tree B - Statistical Analysis")
    created_cells.append(f"A{header_row}")

    # Create column headers
    headers = {"A": "Data", "B": "Lookup", "C": "Index", "D": "Match", "E": "Array", "F": "Result"}

    for col, header_text in headers.items():
        write_cell_value(ws, f"{col}{column_header_row}", header_text)
        created_cells.append(f"{col}{column_header_row}")

    # Create lookup table for VLOOKUP references
    # Table structure: Value | Description | Multiplier
    lookup_data = [
        (100, "Alpha", 1.5),
        (200, "Beta", 2.0),
        (300, "Gamma", 2.5),
        (400, "Delta", 3.0),
    ]

    # Create lookup table headers (starting at H column to avoid conflicts)
    lookup_headers = ["Value", "Description", "Multiplier"]
    for i, header in enumerate(lookup_headers):
        col_letter = chr(ord("H") + i)  # H, I, J
        write_cell_value(ws, f"{col_letter}{lookup_table_start_row}", header)
        created_cells.append(f"{col_letter}{lookup_table_start_row}")

    # Populate lookup table data
    for row_idx, (value, desc, mult) in enumerate(lookup_data):
        row = lookup_table_start_row + 1 + row_idx
        write_cell_value(ws, f"H{row}", value)
        write_cell_value(ws, f"I{row}", desc)
        write_cell_value(ws, f"J{row}", mult)
        created_cells.extend([f"H{row}", f"I{row}", f"J{row}"])

    # Create data rows with statistical formulas
    data_values = [100, 200, 300]

    for i, data_value in enumerate(data_values):
        row = data_row_1 + i

        # Data value (column A)
        write_cell_value(ws, f"A{row}", data_value)
        created_cells.append(f"A{row}")

        # VLOOKUP formula (column B): lookup description from table
        # VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
        lookup_range = f"H${lookup_table_start_row + 1}:J${lookup_table_start_row + 4}"
        write_cell_formula(ws, f"B{row}", f"=VLOOKUP(A{row},{lookup_range},2,FALSE)")
        created_cells.append(f"B{row}")

        # INDEX formula (column C): get multiplier value using INDEX
        # INDEX(array, row_num, [col_num])
        multiplier_range = f"J${lookup_table_start_row + 1}:J${lookup_table_start_row + 4}"
        write_cell_formula(
            ws,
            f"C{row}",
            f"=INDEX({multiplier_range},MATCH(A{row},H${lookup_table_start_row + 1}:H${lookup_table_start_row + 4},0))",
        )
        created_cells.append(f"C{row}")

        # MATCH formula (column D): find position of data value in lookup table
        # MATCH(lookup_value, lookup_array, [match_type])
        value_range = f"H${lookup_table_start_row + 1}:H${lookup_table_start_row + 4}"
        write_cell_formula(ws, f"D{row}", f"=MATCH(A{row},{value_range},0)")
        created_cells.append(f"D{row}")

        # Array formula (column E): statistical function on current row's data range
        if i == 0:
            # First row: SUM of data values
            data_range = f"A{data_row_1}:A{data_row_3}"
            write_cell_formula(ws, f"E{row}", f"=SUM({data_range})")
        elif i == 1:
            # Second row: AVERAGE of data values
            data_range = f"A{data_row_1}:A{data_row_3}"
            write_cell_formula(ws, f"E{row}", f"=AVERAGE({data_range})")
        else:
            # Third row: MAX of data values
            data_range = f"A{data_row_1}:A{data_row_3}"
            write_cell_formula(ws, f"E{row}", f"=MAX({data_range})")
        created_cells.append(f"E{row}")

        # Result formula (column F): combine lookup result with array calculation
        write_cell_formula(ws, f"F{row}", f"=C{row}*E{row}")
        created_cells.append(f"F{row}")

    return created_cells


def create_lookup_table_data(ws: Worksheet, start_row: int, start_col: str = "H") -> list[str]:
    """
    Creates a separate lookup table for VLOOKUP references.

    This is used internally by create_tree_b but can also be called separately
    for testing or when building custom workbook layouts.

    Args:
        ws: The worksheet to write to
        start_row: Starting row for the lookup table (1-indexed)
        start_col: Starting column letter (default: "H")

    Returns:
        List[str]: Cell addresses of all cells created in the lookup table
    """
    created_cells = []

    # Create headers
    headers = ["Value", "Description", "Multiplier"]
    for i, header in enumerate(headers):
        col_letter = chr(ord(start_col) + i)
        write_cell_value(ws, f"{col_letter}{start_row}", header)
        created_cells.append(f"{col_letter}{start_row}")

    # Create data rows
    lookup_data = [
        (100, "Alpha", 1.5),
        (200, "Beta", 2.0),
        (300, "Gamma", 2.5),
        (400, "Delta", 3.0),
    ]

    for row_idx, (value, desc, mult) in enumerate(lookup_data):
        row = start_row + 1 + row_idx
        for col_idx, data in enumerate([value, desc, mult]):
            col_letter = chr(ord(start_col) + col_idx)
            write_cell_value(ws, f"{col_letter}{row}", data)
            created_cells.append(f"{col_letter}{row}")

    return created_cells
