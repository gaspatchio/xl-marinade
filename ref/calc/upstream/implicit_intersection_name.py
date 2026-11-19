"""Implicit intersection through a defined name that covers a column range.

Builds a synthetic workbook in a temp directory, then prints the result Excel
documents next to the one IronCalc 0.9.1 gives. Run: uv run --extra calc python implicit_intersection_name.py
"""

import tempfile
from pathlib import Path

import ironcalc
import openpyxl
from openpyxl.workbook.defined_name import DefinedName

with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for i in range(1, 6):
        ws.cell(i, 1, i * 100)
    wb.defined_names["Col"] = DefinedName("Col", attr_text="S!$A$1:$A$5")
    ws["C3"] = "=Col"
    wb.save(path)
    model = ironcalc.load_from_xlsx(str(path), "en", "UTC", "en")
    model.evaluate()
    print("C3 = Col, where Col names S!A1:A5 (single-cell use of a range name)")
    print("  Excel (documented): 300 (implicit intersection takes the cell on row 3)")
    print(f"  IronCalc 0.9.1:     {model.get_cell_value(0, 3, 3)!r}")
