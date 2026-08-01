# ABOUTME: Builder for Tree A financial calculations with dependency chains and time series patterns.
# ABOUTME: Creates input values, calculation formulas, and final outputs following the test workbook design.


from openpyxl.worksheet.worksheet import Worksheet

from .workbook_builder import write_cell_formula, write_cell_value


def create_tree_a(ws: Worksheet, start_row: int = 1) -> list[str]:
    """
    Creates Tree A financial calculations with proper dependency chains.

    Tree A contains:
    - Header row with section title
    - Column headers (Input, Value, Calc1, Calc2, Time, Growth, Final)
    - Three rows of financial data (Rate, Years, Base)
    - Formulas creating dependency chain: input → calc1 → calc2 → time/growth → final

    Args:
        ws: The worksheet to write to
        start_row: Starting row for Tree A (1-indexed, default: 1)

    Returns:
        List[str]: Cell addresses of all cells created in Tree A

    Example:
        >>> wb = create_workbook()
        >>> ws = wb.active
        >>> cells = create_tree_a(ws, start_row=1)
        >>> len(cells)  # Should be 15-20 cells
        19
    """
    if start_row < 1:
        raise ValueError("start_row must be >= 1")

    created_cells = []

    # Row offsets from start_row
    header_row = start_row
    column_header_row = start_row + 1
    rate_row = start_row + 2
    years_row = start_row + 3
    base_row = start_row + 4

    # Create main header
    write_cell_value(ws, f"A{header_row}", "Tree A - Financial Calculations")
    created_cells.append(f"A{header_row}")

    # Create column headers
    headers = {
        "A": "Input",
        "B": "Value",
        "C": "Calc1",
        "D": "Calc2",
        "E": "Time",
        "F": "Growth",
        "G": "Final",
    }

    for col, header_text in headers.items():
        write_cell_value(ws, f"{col}{column_header_row}", header_text)
        created_cells.append(f"{col}{column_header_row}")

    # Create input rows with labels and values
    inputs = [("Rate", 0.05, rate_row), ("Years", 10, years_row), ("Base", 1000, base_row)]

    for label, value, row in inputs:
        # Input label (column A)
        write_cell_value(ws, f"A{row}", label)
        created_cells.append(f"A{row}")

        # Input value (column B)
        write_cell_value(ws, f"B{row}", value)
        created_cells.append(f"B{row}")

        # Calc1 formula (column C): multiply value by row-specific factor
        multiplier = row - start_row  # 2, 3, 4 for rows 3, 4, 5 when start_row=1
        write_cell_formula(ws, f"C{row}", f"=B{row}*{multiplier}")
        created_cells.append(f"C{row}")

        # Calc2 formula (column D): add row-specific increment to Calc1
        increment = row - start_row - 1  # 1, 2, 3 for rows 3, 4, 5 when start_row=1
        write_cell_formula(ws, f"D{row}", f"=C{row}+{increment}")
        created_cells.append(f"D{row}")

        # Time formula (column E): multiply Calc2 by original input value
        write_cell_formula(ws, f"E{row}", f"=D{row}*B{row}")
        created_cells.append(f"E{row}")

        # Growth formula (column F): square the Time value
        write_cell_formula(ws, f"F{row}", f"=E{row}^2")
        created_cells.append(f"F{row}")

        # Final formula (column G): sum Time and Growth
        write_cell_formula(ws, f"G{row}", f"=E{row}+F{row}")
        created_cells.append(f"G{row}")

    return created_cells
