"""A range whose second corner repeats the sheet name: SUM(S!$A$1:S!$A3).

Builds a synthetic workbook in a temp directory, then prints the result Excel
documents next to the one IronCalc 0.9.1 gives. Run: uv run --extra calc python sheet_qualified_range.py
"""

import tempfile
from pathlib import Path

import ironcalc
import openpyxl

with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for i, v in enumerate([1, 2, 3, 4], start=1):
        ws.cell(i, 1, v)
    ws["C1"] = "=SUM(S!$A$1:S!$A3)"
    wb.save(path)
    model = ironcalc.load_from_xlsx(str(path), "en", "UTC", "en")
    model.evaluate()
    print("SUM(S!$A$1:S!$A3) over 1, 2, 3, 4")
    print("  Excel (documented): 6")
    print(f"  IronCalc 0.9.1:     {model.get_cell_value(0, 1, 3)!r}")
