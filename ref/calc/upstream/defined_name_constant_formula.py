"""Defined names that hold a constant or a formula are rejected.

Builds a synthetic workbook in a temp directory, then prints the result Excel
documents next to the one IronCalc 0.9.1 gives. Run: uv run --extra calc python defined_name_constant_formula.py
"""

import tempfile
from pathlib import Path

import ironcalc
import openpyxl
from openpyxl.workbook.defined_name import DefinedName

for label, text in [("constant", "0.05"), ("formula", "Sheet!$A$1*2")]:
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "book.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet"
        ws["A1"] = 10
        wb.defined_names["Rate"] = DefinedName("Rate", attr_text=text)
        ws["B1"] = "=Rate"
        wb.save(path)
        print(f"name Rate holds a {label}: {text}; B1 = Rate")
        print("  Excel (documented): " + ("0.05" if label == "constant" else "20"))
        try:
            model = ironcalc.load_from_xlsx(str(path), "en", "UTC", "en")
            model.evaluate()
            print(f"  IronCalc 0.9.1:     {model.get_cell_value(0, 1, 2)!r}")
        except Exception as exc:
            print(f"  IronCalc 0.9.1:     load failed: {exc}")
