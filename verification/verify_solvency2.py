"""Value proof: Solvency2 workbook vs lifelib reference values.

Deep profiles (policy | scen | t0): policy scalars and resolved lookups,
the full PREM-basis commutation table, the shared policy schedule, every
mapped column on all seven stress sheets, and the SCR aggregation block.

Vectors are compared for t = 0..last_t (where lifelib defines values);
the guarded zero region beyond last_t is Excel-only structure covered by
the Checks sheet. Run with profile args to chunk:
    verify_solvency2.py "1|1|0" "101|1|0"
"""

import json
import sys
from datetime import date
from pathlib import Path

import xlwings as xw

from excel_lock import hold

REPO = Path(__file__).resolve().parents[1]
_XLSM = REPO / "workbooks" / "Solvency2" / "Solvency2.xlsm"
WB_PATH = _XLSM if _XLSM.exists() else REPO / "workbooks" / "Solvency2" / "Solvency2.xlsx"
REF_PATH = REPO / "verification" / "reference" / "solvency2.json"
REPORT_DIR = REPO / "verification" / "reports"

FIRST = 7      # first data row on time sheets
CFIRST = 4     # first data row on Commutation
RTOL = 1e-9
ATOL = 1e-9

SCALAR_MAP = {
    "product": "Product", "policy_type": "PolicyType", "gen": "Gen", "sex": "Sex",
    "issue_age": "IssueAge", "policy_term": "PolicyTerm", "policy_count": "PolicyCount",
    "sum_assured": "SumAssured", "prem_freq": "PremFreq",
    "int_rate_prem": "IntRatePrem", "table_id_prem": "TableIDPrem",
    "load_acq_sa": "LoadAcqSA", "load_maint_prem": "LoadMaintPrem",
    "load_maint_sa": "LoadMaintSA", "load_maint_sa2": "LoadMaintSA2",
    "load_maint_prem_waiver": "LoadWaiver", "init_surr_charge": "InitSurrCharge",
    "gross_prem_rate": "GrossPremRate", "ann_prem_rate": "AnnPremRate",
    "net_prem_rate_prem": "NetPremRatePrem", "cnsmp_tax": "CnsmpTax",
    "infl_rate": "InflRate", "last_age": "LastAge",
    "comm_init_prem": "CommInitPrem", "comm_ren_prem": "CommRenPrem",
    "comm_ren_term": "CommRenTerm", "exps_acq_ann_prem": "ExpsAcqAnnPrem",
    "exps_acq_pol": "ExpsAcqPol", "exps_acq_sa": "ExpsAcqSA",
    "exps_maint_ann_prem": "ExpsMaintAnnPrem", "exps_maint_pol": "ExpsMaintPol",
    "exps_maint_sa": "ExpsMaintSA", "last_t": "LastT",
    "size_premium_0": "SizePremium", "size_ann_prem_0": "SizeAnnPrem",
    "size_exps_acq_0": "SizeExpsAcq0", "size_exps_comm_init_0": "SizeExpsCommInit0",
}
COMM_MAP = {"qx": "B", "lx": "C", "dx": "D", "Dx": "E", "Cx": "F", "Nx": "G", "Mx": "H"}
SCHED_MAP = {
    "att_age": "B", "base_mort_rate": "C", "mort_factor": "D", "surr_rate_base": "E",
    "reserve_nlp_prem": "F", "surr_charge": "G", "cash_value_rate": "H",
    "size_benefit_surr": "I", "size_exps_comm_ren": "J", "disc_rate": "K",
}
STRESS_COL = {
    "infl_factor": "B", "mort_rate_factor": "C", "surr_rate": "D",
    "pols_if_end": "E", "pols_maturity": "F", "pols_if_beg": "G",
    "pols_surr_mass": "I", "pols_if_beg1": "J", "pols_death": "K",
    "pols_surr": "L", "size_exps_maint": "M", "prem_income": "N",
    "benefit_death": "O", "benefit_surr": "P", "benefit_total": "Q",
    "exps_acq": "R", "exps_comm_init": "S", "exps_comm_ren": "T",
    "exps_maint": "U", "exps_total": "V", "pv_prem_income": "W",
    "pv_benefit_total": "X", "pv_exps_total": "Y", "pv_net_cashflow": "Z",
}
STRESS_SHEET = {
    "base": "Proj_Base", "mort": "Proj_Mort", "longev": "Proj_Longev",
    "exps": "Proj_Exps", "lapse_up": "Proj_LapseUp",
    "lapse_down": "Proj_LapseDown", "lapse_mass": "Proj_LapseMass",
}
NAV_NAME = {"base": "NAV_Base", "mort": "NAV_Mort", "longev": "NAV_Longev",
            "exps": "NAV_Exps", "lapse_up": "NAV_LapseUp",
            "lapse_down": "NAV_LapseDown", "lapse_mass": "NAV_LapseMass"}


def close(a, b):
    return abs(a - b) <= max(ATOL, RTOL * max(abs(a), abs(b)))


def main():
    ref = json.loads(REF_PATH.read_text())["profiles"]
    if len(sys.argv) > 1:
        ref = {k: v for k, v in ref.items() if k in sys.argv[1:]}
    with hold("Solvency2 value proof"):
        return _run(ref)


def _run(ref):
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    failures, results = [], {}
    try:
        wb = app.books.open(str(WB_PATH))
        for key, exp in ref.items():
            pid, scen, t0 = (int(x) for x in key.split("|"))
            wb.names["PolicyID"].refers_to_range.value = pid
            wb.names["ScenID"].refers_to_range.value = scen
            wb.names["T0"].refers_to_range.value = t0
            app.calculate()
            n = int(exp["scalars"]["last_t"])
            ncmp = 0

            def cmp1(name, want, got):
                nonlocal ncmp
                ncmp += 1
                if isinstance(want, str):
                    if got != want:
                        failures.append((key, name, "scalar", want, got))
                    return
                gv = float(got) if isinstance(got, (int, float)) else None
                if gv is None or not close(gv, float(want)):
                    failures.append((key, name, "scalar", want, got))

            for sk, nm in SCALAR_MAP.items():
                cmp1(nm, exp["scalars"][sk], wb.names[nm].refers_to_range.value)

            def cmp_vec(label_, sheet, col, want, first_row, count):
                nonlocal ncmp
                got = wb.sheets[sheet].range(
                    f"{col}{first_row}:{col}{first_row + count - 1}").value
                if count == 1:
                    got = [got]
                for i, (g, w) in enumerate(zip(got, want)):
                    ncmp += 1
                    gv = float(g) if isinstance(g, (int, float)) else None
                    if gv is None or not close(gv, float(w)):
                        failures.append((key, label_, f"i={i}", w, g))

            for ck, col in COMM_MAP.items():
                cmp_vec(f"comm.{ck}", "Commutation", col, exp["commutation"][ck],
                        CFIRST, 131)
            for sk, col in SCHED_MAP.items():
                cmp_vec(f"sched.{sk}", "Policy_Sched", col, exp["shared"][sk][:n + 1],
                        FIRST, n + 1)
            for stress, vecs in exp["stresses"].items():
                sheet = STRESS_SHEET[stress]
                for vk, col in STRESS_COL.items():
                    if vk in vecs:
                        want = vecs[vk][:n + 1]
                    elif vk == "mort_rate_factor":
                        want = [1.0] * (n + 1)     # no override on this sheet
                    elif vk == "surr_rate":
                        want = exp["shared"]["surr_rate_base"][:n + 1]
                    elif vk == "pols_surr_mass":
                        want = [0.0] * (n + 1)
                    else:
                        continue
                    cmp_vec(f"{stress}.{vk}", sheet, col, want, FIRST, n + 1)

            for stress, nm in NAV_NAME.items():
                cmp1(nm, exp["scr"]["net_ast_value"][stress],
                     wb.names[nm].refers_to_range.value)
            for sh, nm in [("up", "LapseUp_Risk"), ("down", "LapseDown_Risk"),
                           ("mass", "LapseMass_Risk")]:
                cmp1(nm, exp["scr"]["lapse_risk"][sh], wb.names[nm].refers_to_range.value)
            for risk, want in exp["scr"]["life"].items():
                cmp1(f"Life_{risk}", want,
                     wb.names[f"Life_{risk.capitalize()}"].refers_to_range.value)
            cmp1("SCR_Life", exp["scr"]["scr_life"], wb.names["SCR_Life"].refers_to_range.value)

            allc = wb.names["AllChecksPass"].refers_to_range.value
            ncmp += 1
            if allc is not True:
                failures.append((key, "AllChecksPass", "check", True, allc))

            nf = len([f for f in failures if f[0] == key])
            results[key] = {"comparisons": ncmp, "failures": nf}
            print(f"profile {key}: {ncmp} comparisons, {nf} failures")

        wb.names["PolicyID"].refers_to_range.value = 1
        wb.names["ScenID"].refers_to_range.value = 1
        wb.names["T0"].refers_to_range.value = 0
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
    out = REPORT_DIR / f"solvency2_value_proof{suffix}.json"
    out.write_text(json.dumps(report, indent=1))
    print(f"\n{report['verdict']}: {total} comparisons, {len(failures)} failures -> {out}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
