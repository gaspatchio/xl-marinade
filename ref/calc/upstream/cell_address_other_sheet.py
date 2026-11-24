"""CELL("address", ref) on a reference to another sheet.

Builds a synthetic workbook in a temp directory, then prints the result Excel
documents next to the one IronCalc 0.9.1 gives. Run: uv run --extra calc python cell_address_other_sheet.py
"""

import tempfile
from pathlib import Path

import ironcalc
import openpyxl

with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "A"
    wb.create_sheet("B")["C3"] = 1
    ws["A1"] = '=CELL("address",B!C3)'
    wb.save(path)
    model = ironcalc.load_from_xlsx(str(path), "en", "UTC", "en")
    model.evaluate()
    print('CELL("address", B!C3) entered on sheet A')
    print("  Excel (documented): an address text such as '[book]B!$C$3', not an error")
    print(f"  IronCalc 0.9.1:     {model.get_cell_value(0, 1, 1)!r}")
