"""Approximate VLOOKUP of a text key against a numeric column.

Builds a synthetic workbook in a temp directory, then prints the result Excel
documents next to the one IronCalc 0.9.1 gives. Run: uv run --extra calc python vlookup_text_vs_numbers.py
"""

import tempfile
from pathlib import Path

import ironcalc
import openpyxl


def calc(path: Path, cell: str) -> object:
    model = ironcalc.load_from_xlsx(str(path), "en", "UTC", "en")
    model.evaluate()
    return model.get_cell_value(0, int(cell[1:]), ord(cell[0]) - 64)


with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, v in enumerate([1, 2, 3, 4], start=1):
        ws.cell(i, 1, v)
        ws.cell(i, 2, v * 10)
    ws["D1"] = "abc"
    ws["E1"] = "=VLOOKUP(D1,A1:B4,2,TRUE)"
    wb.save(path)
    print('VLOOKUP("abc", numbers, 2, TRUE)')
    print("  Excel (documented): #N/A (text never matches a number)")
    print(f"  IronCalc 0.9.1:     {calc(path, 'E1')!r}")
