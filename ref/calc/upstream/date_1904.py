"""The 1904 date system is ignored.

Builds a synthetic workbook in a temp directory, then prints the result Excel
documents next to the one IronCalc 0.9.1 gives. Run: uv run --extra calc python date_1904.py
"""

import tempfile
from pathlib import Path

import ironcalc
import openpyxl

with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / "book.xlsx"
    wb = openpyxl.Workbook()
    wb.epoch = openpyxl.utils.datetime.CALENDAR_MAC_1904
    wb.active["A1"] = "=DATE(2020,1,1)"
    wb.save(path)
    model = ironcalc.load_from_xlsx(str(path), "en", "UTC", "en")
    model.evaluate()
    print("DATE(2020,1,1) in a workbook with date1904=1")
    print("  Excel (documented): 42369")
    print(f"  IronCalc 0.9.1:     {model.get_cell_value(0, 1, 1)!r} (setting ignored, no error)")
