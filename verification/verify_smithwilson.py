"""Value proof: SmithWilson workbook vs lifelib reference values.

Compares every mapped cell — the full 65x25 Wilson grid, the calibration
vectors (m, mu, zeta incl. the MINVERSE/MMULT array result) and the
extrapolated curve (mu, P, R) — against verification/reference/smithwilson.json.
"""

import json
import sys
from datetime import date
from pathlib import Path

import xlwings as xw

from excel_lock import hold

REPO = Path(__file__).resolve().parents[1]
_XLSM = REPO / "workbooks" / "SmithWilson" / "SmithWilson.xlsm"
WB_PATH = _XLSM if _XLSM.exists() else REPO / "workbooks" / "SmithWilson" / "SmithWilson.xlsx"
REF_PATH = REPO / "verification" / "reference" / "smithwilson.json"
REPORT = REPO / "verification" / "reports" / "smithwilson_value_proof.json"

FIRST_ROW = 4
RTOL = 1e-9
ATOL = 1e-9


def close(a, b):
    return abs(a - b) <= max(ATOL, RTOL * max(abs(a), abs(b)))


def main():
    ref = json.loads(REF_PATH.read_text())
    with hold("SmithWilson value proof"):
        return _run(ref)


def _run(ref):
    n = ref["params"]["N"]
    t_max = ref["params"]["T_MAX"]
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    failures = []
    ncmp = 0
    try:
        wb = app.books.open(str(WB_PATH))
        app.calculate()

        def compare_vector(name, sheet, col, want_vec, first=FIRST_ROW):
            nonlocal ncmp
            m = len(want_vec)
            got = wb.sheets[sheet].range(f"{col}{first}:{col}{first + m - 1}").value
            if m == 1:
                got = [got]
            for i, (g, w) in enumerate(zip(got, want_vec)):
                ncmp += 1
                gv = float(g) if isinstance(g, (int, float)) else None
                if gv is None or not close(gv, float(w)):
                    failures.append((name, f"i={i + 1}", w, g))

        compare_vector("u", "Extrapolation", "A", ref["u"])
        compare_vector("mu", "Extrapolation", "B", ref["mu"])
        compare_vector("P", "Extrapolation", "C", ref["P"])
        compare_vector("R", "Extrapolation", "D", ref["R"])
        compare_vector("m", "Calibration", "D", ref["m"])
        compare_vector("mu@cal", "Calibration", "E", ref["mu"][:n])
        compare_vector("zeta", "Calibration", "G", ref["zeta"])
        compare_vector("spot_rates", "Spot_Rates", "C", ref["spot_rates"])

        grid = wb.sheets["Wilson"].range(
            f"B{FIRST_ROW}:Z{FIRST_ROW + t_max - 1}").value
        for i, row in enumerate(grid):
            for j, g in enumerate(row):
                ncmp += 1
                w = ref["W"][i][j]
                gv = float(g) if isinstance(g, (int, float)) else None
                if gv is None or not close(gv, float(w)):
                    failures.append(("W", f"i={i + 1},j={j + 1}", w, g))

        all_checks = wb.names["AllChecksPass"].refers_to_range.value
        ncmp += 1
        if all_checks is not True:
            failures.append(("AllChecksPass", "check", True, all_checks))

        wb.save()
        wb.close()
    finally:
        app.quit()

    report = {
        "workbook": str(WB_PATH.relative_to(REPO)),
        "reference": str(REF_PATH.relative_to(REPO)),
        "date": date.today().isoformat(),
        "total_comparisons": ncmp,
        "total_failures": len(failures),
        "tolerance": {"rtol": RTOL, "atol": ATOL},
        "failures": [
            {"cell": c, "where": wh, "expected": e, "got": g}
            for c, wh, e, g in failures[:200]
        ],
        "verdict": "PASS" if not failures else "FAIL",
    }
    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text(json.dumps(report, indent=1))
    print(f"{report['verdict']}: {ncmp} comparisons, "
          f"{len(failures)} failures -> {REPORT}")
    for f in failures[:15]:
        print("  FAIL:", f)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
