"""Value proof: BasicTerm_SE.xlsm vs lifelib reference values.

Same pattern as verify_basicterm_s.py. SE additions: the three pols_if_at
timing vectors map to their own columns; reference entries that are null
(mortality unevaluable at negative durations) or NaN (0/0 net premium for
zero-count points) are skipped.
"""

import json
import math
import sys
from datetime import date
from pathlib import Path

import xlwings as xw

from excel_lock import hold

REPO = Path(__file__).resolve().parents[1]
_XLSM = REPO / "workbooks" / "BasicTerm_SE" / "BasicTerm_SE.xlsm"
WB_PATH = _XLSM if _XLSM.exists() else REPO / "workbooks" / "BasicTerm_SE" / "BasicTerm_SE.xlsx"
REF_PATH = REPO / "verification" / "reference" / "basicterm_se.json"
REPORT_DIR = REPO / "verification" / "reports"

FIRST_ROW = 4  # Projection row for t = 0

TIME_MAP = {
    "duration_mth": "B",
    "duration": "C",
    "age": "D",
    "is_active": "E",
    "mort_rate": "F",
    "mort_rate_mth": "G",
    "lapse_rate": "H",
    "pols_if": "I",
    "pols_maturity": "J",
    "pols_new_biz": "L",
    "pols_death": "N",
    "pols_lapse": "O",
    "premiums": "P",
    "claims": "Q",
    "commissions": "R",
    "inflation_factor": "S",
    "expenses": "T",
    "net_cf": "U",
}
TIMING_MAP = {"BEF_MAT": "I", "BEF_NB": "K", "BEF_DECR": "M"}
DISC_FACTOR_COL = "W"

# mort columns are guarded by is_active in the workbook (0 when inactive);
# lifelib computes real values at inactive-but-evaluable months. Compare
# these only where the policy is active.
ACTIVE_ONLY = {"mort_rate", "mort_rate_mth"}

SCALAR_MAP = {
    "age_at_entry": "AgeAtEntry",
    "policy_term": "PolicyTerm",
    "sum_assured": "SumAssured",
    "policy_count": "PolicyCount",
    "duration_mth_0": "Duration0",
    "pols_if_init": "PolsIfInit",
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


def skip(want):
    return want is None or (isinstance(want, float) and math.isnan(want))


def main():
    ref = json.loads(REF_PATH.read_text())["lifelib_points"]
    with hold("BasicTerm_SE value proof"):
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
            active = [bool(v) for v in exp["time_vectors"]["is_active"]]
            ncmp = 0

            for cell, name in SCALAR_MAP.items():
                want = exp["scalars"][cell]
                if skip(want):
                    continue
                got = wb.names[name].refers_to_range.value
                ncmp += 1
                if got is None or not isinstance(got, (int, float)) \
                        or not close(float(got), want):
                    failures.append((pid, cell, "scalar", want, got))

            sht = wb.sheets["Projection"]

            def compare_vector(cell_name, col, want_vec):
                nonlocal ncmp
                got = sht.range(f"{col}{FIRST_ROW}:{col}{FIRST_ROW + n - 1}").value
                if n == 1:
                    got = [got]
                for t, (g, w) in enumerate(zip(got, want_vec)):
                    if skip(w):
                        continue
                    if cell_name in ACTIVE_ONLY and not active[t]:
                        continue
                    ncmp += 1
                    gv = float(g) if isinstance(g, (int, float, bool)) else None
                    if gv is None or not close(gv, float(w)):
                        failures.append((pid, cell_name, f"t={t}", w, g))

            for cell, col in TIME_MAP.items():
                compare_vector(cell, col, exp["time_vectors"][cell])
            for timing, col in TIMING_MAP.items():
                compare_vector(f"pols_if_at[{timing}]", col, exp["pols_if_at"][timing])
            compare_vector("disc_factors", DISC_FACTOR_COL, exp["disc_factors"])

            all_checks = wb.names["AllChecksPass"].refers_to_range.value
            ncmp += 1
            if all_checks is not True:
                failures.append((pid, "AllChecksPass", "check", True, all_checks))

            n_fail = len([f for f in failures if f[0] == pid])
            results[pid] = {"comparisons": ncmp, "failures": n_fail}
            print(f"point {pid}: {ncmp} comparisons, {n_fail} failures")

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
    out = REPORT_DIR / "basicterm_se_value_proof.json"
    out.write_text(json.dumps(report, indent=1))
    print(f"\n{report['verdict']}: {total} comparisons, "
          f"{len(failures)} failures -> {out}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
