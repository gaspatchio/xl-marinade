"""Value proof: BasicTerm_S.xlsx vs lifelib reference values.

For each test model point: set PointID in the workbook, recalculate in
Excel (via xlwings), read every mapped cell, and compare against
verification/reference/basicterm_s.json. Writes a report to
verification/reports/ and exits non-zero on any mismatch.

Cell mapping (lifelib cell -> workbook location):
  time vectors -> Projection sheet columns, data rows start at row 4 (t=0)
  scalars      -> named cells (Summary/Assumptions) per SCALAR_MAP
"""

import json
import sys
from datetime import date
from pathlib import Path

import xlwings as xw

from excel_lock import hold

REPO = Path(__file__).resolve().parents[1]
_XLSM = REPO / "workbooks" / "BasicTerm_S" / "BasicTerm_S.xlsm"
WB_PATH = _XLSM if _XLSM.exists() else REPO / "workbooks" / "BasicTerm_S" / "BasicTerm_S.xlsx"
REF_PATH = REPO / "verification" / "reference" / "basicterm_s.json"
REPORT_DIR = REPO / "verification" / "reports"

FIRST_ROW = 4  # Projection row for t = 0

TIME_MAP = {  # lifelib cell -> Projection column
    "duration": "B",
    "age": "C",
    "mort_rate": "D",
    "mort_rate_mth": "E",
    "lapse_rate": "F",
    "pols_if": "G",
    "pols_death": "H",
    "pols_lapse": "I",
    "pols_maturity": "J",
    "premiums": "K",
    "claims": "L",
    "commissions": "M",
    "inflation_factor": "N",
    "expenses": "O",
    "net_cf": "P",
}
DISC_RATE_MTH_COL = "Q"  # lifelib disc_rate_mth() — 0-ary vector cell
DISC_FACTOR_COL = "R"
PV_NET_CF_T_COL = "S"  # Excel presentation column = net_cf(t) * disc_factor(t)

SCALAR_MAP = {  # lifelib cell -> workbook defined name
    "age_at_entry": "AgeAtEntry",
    "policy_term": "PolicyTerm",
    "sum_assured": "SumAssured",
    "proj_len": "ProjLen",
    "pv_claims": "PV_Claims",
    "pv_premiums": "PV_Premiums",
    "pv_expenses": "PV_Expenses",
    "pv_commissions": "PV_Commissions",
    "pv_net_cf": "PV_NetCF",
    "pv_pols_if": "PV_PolsIf",
    "net_premium_pp": "NetPremiumPP",
    "premium_pp": "PremiumPP",
}

RTOL = 1e-9
ATOL = 1e-9


def close(a, b):
    return abs(a - b) <= max(ATOL, RTOL * max(abs(a), abs(b)))


def main():
    ref = json.loads(REF_PATH.read_text())["lifelib_points"]
    with hold("BasicTerm_S value proof"):
        return _run(ref)


def _run(ref):
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    failures, results = [], {}
    try:
        wb = app.books.open(str(WB_PATH))
        for pid, exp in ref.items():
            wb.names["PointID"].refers_to_range.value = int(pid)
            app.calculate()
            n = int(exp["scalars"]["proj_len"])
            ncmp = 0

            for cell, name in SCALAR_MAP.items():
                got = wb.names[name].refers_to_range.value
                want = exp["scalars"][cell]
                ncmp += 1
                if got is None or not close(float(got), want):
                    failures.append((pid, cell, "scalar", want, got))

            sht = wb.sheets["Projection"]
            for cell, col in TIME_MAP.items():
                got = sht.range(f"{col}{FIRST_ROW}:{col}{FIRST_ROW + n - 1}").value
                want = exp["time_vectors"][cell]
                for t, (g, w) in enumerate(zip(got, want)):
                    ncmp += 1
                    if g is None or not close(float(g), w):
                        failures.append((pid, cell, f"t={t}", w, g))

            got = sht.range(
                f"{DISC_FACTOR_COL}{FIRST_ROW}:{DISC_FACTOR_COL}{FIRST_ROW + n - 1}").value
            for t, (g, w) in enumerate(zip(got, exp["disc_factors"])):
                ncmp += 1
                if g is None or not close(float(g), w):
                    failures.append((pid, "disc_factors", f"t={t}", w, g))

            got = sht.range(
                f"{DISC_RATE_MTH_COL}{FIRST_ROW}:{DISC_RATE_MTH_COL}{FIRST_ROW + n - 1}").value
            for t, (g, w) in enumerate(zip(got, exp["disc_rate_mth"])):
                ncmp += 1
                if g is None or not close(float(g), w):
                    failures.append((pid, "disc_rate_mth", f"t={t}", w, g))

            # pv_net_cf_t has no lifelib counterpart cell; it must equal
            # net_cf(t) * disc_factor(t) from the lifelib vectors
            got = sht.range(
                f"{PV_NET_CF_T_COL}{FIRST_ROW}:{PV_NET_CF_T_COL}{FIRST_ROW + n - 1}").value
            want_pv = [nc * df for nc, df in
                       zip(exp["time_vectors"]["net_cf"], exp["disc_factors"])]
            for t, (g, w) in enumerate(zip(got, want_pv)):
                ncmp += 1
                if g is None or not close(float(g), w):
                    failures.append((pid, "pv_net_cf_t", f"t={t}", w, g))

            all_checks = wb.names["AllChecksPass"].refers_to_range.value
            ncmp += 1
            if all_checks is not True:
                failures.append((pid, "AllChecksPass", "check", True, all_checks))

            n_fail = len([f for f in failures if f[0] == pid])
            results[pid] = {"comparisons": ncmp, "failures": n_fail}
            print(f"point {pid}: {ncmp} comparisons, {n_fail} failures")

        # leave the workbook on point 1 with fresh cached values (Marinade
        # reads cached values, so save after a final recalc)
        wb.names["PointID"].refers_to_range.value = 1
        app.calculate()
        wb.save()
        wb.close()
    finally:
        app.quit()

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    total = sum(r["comparisons"] for r in results.values())
    report = {
        "workbook": str(WB_PATH.relative_to(REPO)),
        "reference": str(REF_PATH.relative_to(REPO)),
        "date": date.today().isoformat(),
        "model_points": results,
        "total_comparisons": total,
        "total_failures": len(failures),
        "tolerance": {"rtol": RTOL, "atol": ATOL},
        "failures": [
            {"point": p, "cell": c, "where": w, "expected": e, "got": g}
            for p, c, w, e, g in failures[:200]
        ],
        "verdict": "PASS" if not failures else "FAIL",
    }
    out = REPORT_DIR / "basicterm_s_value_proof.json"
    out.write_text(json.dumps(report, indent=1))
    print(f"\n{report['verdict']}: {total} comparisons, "
          f"{len(failures)} failures -> {out}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
