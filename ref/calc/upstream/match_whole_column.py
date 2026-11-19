"""Approximate MATCH over a whole column.

Builds a synthetic workbook in a temp directory, then prints the result Excel
documents next to the one IronCalc 0.9.1 gives. Run: uv run --extra calc python match_whole_column.py
"""

import tempfile
from pathlib import Path

import ironcalc
import openpyxl

with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, v in enumerate([10, 20, 30], start=1):
        ws.cell(i, 1, v)
    ws["C1"] = "=MATCH(25,A:A,1)"
    wb.save(path)
    model = ironcalc.load_from_xlsx(str(path), "en", "UTC", "en")
    model.evaluate()
    print("MATCH(25, A:A, 1) over 10, 20, 30 then empty cells")
    print("  Excel (documented): 2 (empty cells are skipped)")
    print(f"  IronCalc 0.9.1:     {model.get_cell_value(0, 1, 3)!r}")
