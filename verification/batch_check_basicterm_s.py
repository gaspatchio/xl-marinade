"""All-points reconciliation: workbook engine vs lifelib for every model point.

Python mirror of the workbook's RunAllModelPoints VBA macro: loop every
model point, recalc in Excel, compare office premium and PV net cash flow
against verification/reference/basicterm_all_points.csv. Proves the engine
reconciles across the whole portfolio, not just the 10 deep-proof points.

Usage: batch_check_basicterm_s.py [n_points]   (default: all 10,000)
"""

import json
import sys
import time
from datetime import date
from pathlib import Path

import pandas as pd
import xlwings as xw

from excel_lock import hold

REPO = Path(__file__).resolve().parents[1]
XLSM = REPO / "workbooks" / "BasicTerm_S" / "BasicTerm_S.xlsm"
XLSX = REPO / "workbooks" / "BasicTerm_S" / "BasicTerm_S.xlsx"
REF = REPO / "verification" / "reference" / "basicterm_all_points.csv"
# _external suffix: the canonical report (basicterm_s_batch_reconciliation.json)
# records the in-workbook VBA run; this script's runs must not overwrite it
REPORT = REPO / "verification" / "reports" / "basicterm_s_batch_reconciliation_external.json"

TOL = 1e-6  # relative, same as the workbook's BatchTol default


def main():
    n_limit = int(sys.argv[1]) if len(sys.argv) > 1 else None
    ref = pd.read_csv(REF, index_col=0)
    if n_limit:
        ref = ref.iloc[:n_limit]

    wb_path = XLSM if XLSM.exists() else XLSX
    with hold("BasicTerm_S batch reconciliation"):
        return _run(ref, wb_path)


def _run(ref, wb_path):
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    t0 = time.time()
    try:
        wb = app.books.open(str(wb_path))
        app.calculation = "manual"
        point_cell = wb.names["PointID"].refers_to_range
        prem_cell = wb.names["PremiumPP"].refers_to_range
        pv_cell = wb.names["PV_NetCF"].refers_to_range

        n_ok, worst_prem, worst_pv, failures = 0, 0.0, 0.0, []
        for i, (pid, row) in enumerate(ref.iterrows()):
            point_cell.value = int(pid)
            app.calculate()
            prem = float(prem_cell.value)
            pv = float(pv_cell.value)
            prem_ref = float(row["premium_pp"])
            pv_ref = float(row["pv_net_cf"])
            d_prem = abs(prem - prem_ref)
            r_pv = abs(pv - pv_ref) / max(1.0, abs(pv_ref))
            worst_prem = max(worst_prem, d_prem)
            worst_pv = max(worst_pv, r_pv)
            ok = bool(d_prem <= TOL * max(1.0, abs(prem_ref)) and r_pv <= TOL)
            n_ok += ok
            if not ok and len(failures) < 50:
                failures.append({"point": int(pid), "prem_wb": prem,
                                 "prem_ref": prem_ref, "pv_wb": pv,
                                 "pv_ref": pv_ref})
            if (i + 1) % 1000 == 0:
                rate = (i + 1) / (time.time() - t0)
                print(f"{i + 1}/{len(ref)} points ({rate:.0f}/s), "
                      f"worst prem diff {worst_prem:.2e}, worst PV rel {worst_pv:.2e}")

        point_cell.value = 1
        app.calculation = "automatic"
        app.calculate()
        wb.save()
        wb.close()
    finally:
        app.quit()

    verdict = "PASS" if n_ok == len(ref) else "FAIL"
    report = {
        "workbook": str(wb_path.relative_to(REPO)),
        "date": date.today().isoformat(),
        "points_run": len(ref),
        "reconciled": n_ok,
        "mismatches": len(ref) - n_ok,
        "max_abs_diff_premium": worst_prem,
        "max_rel_diff_pv_net_cf": worst_pv,
        "tolerance_relative": TOL,
        "runtime_seconds": round(time.time() - t0, 1),
        "failures": failures,
        "verdict": verdict,
    }
    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text(json.dumps(report, indent=1))
    print(f"\n{verdict}: {n_ok}/{len(ref)} points reconciled "
          f"(max |prem diff| {worst_prem:.2e}, max PV rel {worst_pv:.2e}) -> {REPORT}")
    return 0 if verdict == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
