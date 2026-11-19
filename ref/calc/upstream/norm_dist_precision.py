"""NORM.DIST (cumulative) is not accurate to full double precision.

Builds a synthetic workbook in a temp directory, then prints the result Excel
documents next to the one IronCalc 0.9.1 gives. The reference value is the exact normal CDF (statistics.NormalDist),
which Excel matches to about 1e-15. Run: uv run --extra calc python norm_dist_precision.py
"""

import tempfile
from pathlib import Path
from statistics import NormalDist

import ironcalc
import openpyxl

xs = [-0.0707687815502868] + [i / 20 for i in range(-200, 201)]
with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / "book.xlsx"
    wb = openpyxl.Workbook()
    for i, x in enumerate(xs, start=1):
        wb.active.cell(i, 1, x)
        wb.active.cell(i, 2, f"=NORM.DIST(A{i},0,1,TRUE)")
    wb.save(path)
    model = ironcalc.load_from_xlsx(str(path), "en", "UTC", "en")
    model.evaluate()
    errors = [
        (abs(model.get_cell_value(0, i, 2) - NormalDist().cdf(x)), x)
        for i, x in enumerate(xs, start=1)
    ]
    first = model.get_cell_value(0, 1, 2)
    print(f"NORM.DIST({xs[0]!r},0,1,TRUE)")
    print(f"  Excel (documented, ~1e-15): {NormalDist().cdf(xs[0])!r}")
    print(f"  IronCalc 0.9.1:             {first!r}")
    worst, at = max(errors)
    print(f"largest error over {len(xs) - 1} points on [-10, 10]: {worst:.2e} at x={at}")
    print("  Excel (documented): below 1e-14")
