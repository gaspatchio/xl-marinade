"""Value proof: CashValue_SE.xlsm vs lifelib reference values.

Six (point, scenario) deep profiles; every mapped engine column including
the four AV timings, three policy-count timings and per-kind claims.
lifelib's own check_* flags are not compared (its math.isclose fails on
exact zeros for the single-premium products); the workbook's Checks sheet
asserts the same identities with per-row relative residuals (largest-side
divisor) and a count-of-violations test instead.
"""

import json
import sys
from datetime import date
from pathlib import Path

import xlwings as xw

from excel_lock import hold

REPO = Path(__file__).resolve().parents[1]
_XLSM = REPO / "workbooks" / "CashValue_SE" / "CashValue_SE.xlsm"
WB_PATH = _XLSM if _XLSM.exists() else REPO / "workbooks" / "CashValue_SE" / "CashValue_SE.xlsx"
REF_PATH = REPO / "verification" / "reference" / "cashvalue_se.json"
REPORT_DIR = REPO / "verification" / "reports"

FIRST_ROW = 4

TIME_MAP = {
    "duration_mth": "B", "duration": "C", "age": "D", "premium_pp": "E",
    "mort_rate": "F", "mort_rate_mth": "G", "lapse_rate": "H",
    "surr_charge_rate": "I",
    "pols_if": "J", "pols_maturity": "K", "pols_new_biz": "M",
    "pols_death": "O", "pols_lapse": "P",
    "prem_to_av_pp": "R", "maint_fee_pp": "T", "net_amt_at_risk": "U",
    "coi_rate": "V", "coi_pp": "W",
    "inv_return_mth": "Y", "inv_income_pp": "Z",
    "premiums": "AE", "prem_to_av": "AF",
    "claims_over_av": "AO", "surr_charge": "AK",
    "coi": "AP", "maint_fee": "AQ", "inv_income": "AR",
    "inflation_factor": "AS", "expenses": "AT", "commissions": "AU",
    "av_change": "AV", "net_cf": "AW",
    "margin_expense": "AX", "margin_mortality": "AY",
}
AV_PP_MAP = {"BEF_PREM": "Q", "BEF_FEE": "S", "BEF_INV": "X", "MID_MTH": "AA"}
AV_MAP = {"BEF_MAT": "AB", "BEF_NB": "AC", "BEF_FEE": "AD"}
POLS_MAP = {"BEF_MAT": "J", "BEF_NB": "L", "BEF_DECR": "N"}
CLAIMS_MAP = {"DEATH": "AH", "LAPSE": "AL", "MATURITY": "AM"}
DISC_FACTOR_COL = "BC"

SCALAR_MAP = {
    "age_at_entry": "AgeAtEntry",
    "policy_term": "PolicyTerm",
    "sum_assured": "SumAssured",
    "proj_len": "ProjLen",
    "mp_policy_count": "PolicyCount",
    "mp_duration_mth": "Duration0",
    "mp_premium_pp": "PremiumPPInput",
    "mp_av_pp_init": "AvPPInit",
    "mp_load_prem_rate": "LoadPremRate",
    "pols_if_init": "PolsIfInit",
    "pv_premiums": "PV_Premiums",
    "pv_claims": "PV_Claims",
    "pv_expenses": "PV_Expenses",
    "pv_commissions": "PV_Commissions",
    "pv_inv_income": "PV_InvIncome",
    "pv_av_change": "PV_AvChange",
    "pv_net_cf": "PV_NetCF",
    "pv_pols_if": "PV_PolsIf",
}

RTOL = 1e-9
ATOL = 1e-9


def close(a, b):
    return abs(a - b) <= max(ATOL, RTOL * max(abs(a), abs(b)))


def main():
    ref = json.loads(REF_PATH.read_text())["profiles"]
    # optional filter: verify_cashvalue_se.py 3|1  (single profile; reports
    # are written per profile and merged by report_merge below)
    if len(sys.argv) > 1:
        ref = {k: v for k, v in ref.items() if k in sys.argv[1:]}
    with hold("CashValue_SE value proof"):
        return _run(ref)


def _run(ref):
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    failures, results = [], {}
    try:
        wb = app.books.open(str(WB_PATH))
        for key, exp in ref.items():
            pid, scen = (int(x) for x in key.split("|"))
            wb.names["PointID"].refers_to_range.value = pid
            wb.names["ScenID"].refers_to_range.value = scen
            app.calculate()
            n = int(exp["scalars"]["proj_len"])
            ncmp = 0

            for cell, name in SCALAR_MAP.items():
                want = exp["scalars"][cell]
                got = wb.names[name].refers_to_range.value
                ncmp += 1
                if got is None or not isinstance(got, (int, float)) \
                        or not close(float(got), want):
                    failures.append((key, cell, "scalar", want, got))

            sht = wb.sheets["Projection"]

            def compare_vector(cell_name, col, want_vec):
                nonlocal ncmp
                got = sht.range(f"{col}{FIRST_ROW}:{col}{FIRST_ROW + n - 1}").value
                if n == 1:
                    got = [got]
                for t, (g, w) in enumerate(zip(got, want_vec)):
                    ncmp += 1
                    gv = float(g) if isinstance(g, (int, float, bool)) else None
                    if gv is None or not close(gv, float(w)):
                        failures.append((key, cell_name, f"t={t}", w, g))

            for cell, col in TIME_MAP.items():
                compare_vector(cell, col, exp["time_vectors"][cell])
            for tm, col in AV_PP_MAP.items():
                compare_vector(f"av_pp_at[{tm}]", col, exp["av_pp_at"][tm])
            for tm, col in AV_MAP.items():
                compare_vector(f"av_at[{tm}]", col, exp["av_at"][tm])
            for tm, col in POLS_MAP.items():
                compare_vector(f"pols_if_at[{tm}]", col, exp["pols_if_at"][tm])
            for kind, col in CLAIMS_MAP.items():
                compare_vector(f"claims[{kind}]", col, exp["claims"][kind])
            compare_vector("disc_factors", DISC_FACTOR_COL, exp["disc_factors"])

            all_checks = wb.names["AllChecksPass"].refers_to_range.value
            ncmp += 1
            if all_checks is not True:
                failures.append((key, "AllChecksPass", "check", True, all_checks))

            n_fail = len([f for f in failures if f[0] == key])
            results[key] = {"comparisons": ncmp, "failures": n_fail}
            print(f"profile {key}: {ncmp} comparisons, {n_fail} failures")

        wb.names["PointID"].refers_to_range.value = 1
        wb.names["ScenID"].refers_to_range.value = 1
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
        "profiles": results,
        "total_comparisons": total,
        "total_failures": len(failures),
        "tolerance": {"rtol": RTOL, "atol": ATOL},
        "failures": [
            {"profile": p, "cell": c, "where": w, "expected": e, "got": g}
            for p, c, w, e, g in failures[:200]
        ],
        "verdict": "PASS" if not failures else "FAIL",
    }
    suffix = ""
    if len(sys.argv) > 1:
        suffix = "_" + "_".join(k.replace("|", "-") for k in sys.argv[1:])
    out = REPORT_DIR / f"cashvalue_se_value_proof{suffix}.json"
    out.write_text(json.dumps(report, indent=1))
    print(f"\n{report['verdict']}: {total} comparisons, "
          f"{len(failures)} failures -> {out}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
