"""Precision as displayed (fullPrecision=0) is ignored.

Builds a synthetic workbook in a temp directory, then prints the result Excel
documents next to the one IronCalc 0.9.1 gives. Run: uv run --extra calc python precision_as_displayed.py
"""

import tempfile
from pathlib import Path

import ironcalc
import openpyxl

with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "=1/3"
    ws["A1"].number_format = "0.00"
    ws["A2"] = "=A1*3"
    wb.calculation.fullPrecision = False
    wb.save(path)
    model = ironcalc.load_from_xlsx(str(path), "en", "UTC", "en")
    model.evaluate()
    print("A1 = 1/3 shown as 0.00, A2 = A1*3, precision as displayed on")
    print("  Excel (documented): 0.99 (A1 is stored as 0.33)")
    print(f"  IronCalc 0.9.1:     {model.get_cell_value(0, 2, 1)!r} (setting ignored, no error)")
