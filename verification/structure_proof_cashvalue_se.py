"""Structure proof: CashValue_SE workbook dependency graph vs lifelib call graph.

Method as the BasicTerm proofs. Notes on mapping:
  - av_pp_at/av_at/pols_if_at timings each map to their own column
  - claims(t, kind) maps to per-kind columns; the aggregate is claims_total
  - claim_pp("DEATH") is the claim_pp_death column; LAPSE/MATURITY claim_pp
    are inlined into the claims_from_av columns (pure products)
  - inv_return_table is the Scenarios sheet grid + mu/sigma assumptions
  - model_point_table_ext's join = Control lookups into Model_Points and
    Product_Specs (spec attributes depend on spec_id, which is the join)
  - mort_table_last_age is the Mortality sheet helper (col I/J)
"""

import json
import re
import sqlite3
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
DB = REPO / "marinade" / "CashValue_SE.ir.db"
REPORT = REPO / "verification" / "reports" / "cashvalue_se_structure_proof.json"

KNOWN_EXTRACTOR_GAPS: dict = {}

LOC = {
    # engine columns (row 10 = t = 6)
    "t": "Projection!A10",
    "duration_mth": "Projection!B10",
    "duration": "Projection!C10",
    "age": "Projection!D10",
    "premium_pp": "Projection!E10",
    "mort_rate": "Projection!F10",
    "mort_rate_mth": "Projection!G10",
    "lapse_rate": "Projection!H10",
    "surr_charge_rate": "Projection!I10",
    "pols_if": "Projection!J10",
    "pols_if@0": "Projection!J4",       # t = 0 seed row (separate binding)
    "pols_maturity": "Projection!K10",
    "pols_if_bef_nb": "Projection!L10",
    "pols_new_biz": "Projection!M10",
    "pols_if_bef_decr": "Projection!N10",
    "pols_death": "Projection!O10",
    "pols_lapse": "Projection!P10",
    "av_pp_bef_prem": "Projection!Q10",
    "av_pp_bef_prem@0": "Projection!Q4",  # t = 0 seed row (separate binding)
    "prem_to_av_pp": "Projection!R10",
    "av_pp_bef_fee": "Projection!S10",
    "maint_fee_pp": "Projection!T10",
    "net_amt_at_risk": "Projection!U10",
    "coi_rate": "Projection!V10",
    "coi_pp": "Projection!W10",
    "av_pp_bef_inv": "Projection!X10",
    "inv_return_mth": "Projection!Y10",
    "inv_income_pp": "Projection!Z10",
    "av_pp_mid_mth": "Projection!AA10",
    "av_bef_mat": "Projection!AB10",
    "av_bef_nb": "Projection!AC10",
    "av_bef_fee": "Projection!AD10",
    "premiums": "Projection!AE10",
    "prem_to_av": "Projection!AF10",
    "claim_pp_death": "Projection!AG10",
    "claims_death": "Projection!AH10",
    "cfav_death": "Projection!AI10",
    "cfav_lapse": "Projection!AJ10",
    "surr_charge": "Projection!AK10",
    "claims_lapse": "Projection!AL10",
    "cfav_maturity": "Projection!AM10",
    "claims_total": "Projection!AN10",
    "claims_over_av": "Projection!AO10",
    "coi": "Projection!AP10",
    "maint_fee": "Projection!AQ10",
    "inv_income": "Projection!AR10",
    "inflation_factor": "Projection!AS10",
    "expenses": "Projection!AT10",
    "commissions": "Projection!AU10",
    "av_change": "Projection!AV10",
    "net_cf": "Projection!AW10",
    "margin_expense": "Projection!AX10",
    "margin_mortality": "Projection!AY10",
    "disc_rate_mth": "Projection!BB10",
    "disc_factors": "Projection!BC10",
    # run control
    "point_id": "Control!C4",
    "scen_id": "Control!C5",
    "spec_id": "Control!C9",
    "age_at_entry": "Control!C10",
    "term_input": "Control!C12",
    "policy_count": "Control!C13",
    "sum_assured": "Control!C14",
    "duration_mth_0": "Control!C15",
    "premium_pp_input": "Control!C16",
    "av_pp_init": "Control!C17",
    "premium_type": "Control!C21",
    "has_surr_charge": "Control!C22",
    "surr_charge_id": "Control!C23",
    "load_prem_rate": "Control!C24",
    "is_wl": "Control!C25",
    "mort_table_last_age": "Control!C29",
    "policy_term": "Control!C30",
    "pols_if_init": "Control!C31",
    # assumptions
    "expense_acq": "Assumptions!C4",
    "expense_maint": "Assumptions!C5",
    "inflation_rate": "Assumptions!C6",
    "commission_rate": "Assumptions!C7",
    "coi_multiple": "Assumptions!C8",
    "maint_fee_rate": "Assumptions!C9",
    "mu": "Assumptions!C19",
    "sigma": "Assumptions!C20",
    # data tables
    "mort_table": "Mortality!C10",
    "disc_rate_ann": "Discount!B10",
    "surr_charge_table": "Surr_Charges!C5",
    "std_norm_rand": "Scenarios!E5",
    "model_point_table[spec]": "Model_Points!B3",
    "model_point_table[age]": "Model_Points!C3",
    "model_point_table[term]": "Model_Points!E3",
    "model_point_table[count]": "Model_Points!F3",
    "model_point_table[sum]": "Model_Points!G3",
    "model_point_table[dur]": "Model_Points!H3",
    "model_point_table[prem]": "Model_Points!I3",
    "model_point_table[avinit]": "Model_Points!J3",
    "product_spec_table[premtype]": "Product_Specs!B3",
    "product_spec_table[hassurr]": "Product_Specs!C3",
    "product_spec_table[surrid]": "Product_Specs!D3",
    "product_spec_table[loadprem]": "Product_Specs!E3",
    "product_spec_table[iswl]": "Product_Specs!F3",
    # summary
    "pv_premiums": "Summary!C10",
    "pv_claims[DEATH]": "Summary!C11",
    "pv_claims[LAPSE]": "Summary!C12",
    "pv_claims[MATURITY]": "Summary!C13",
    "pv_claims": "Summary!C14",
    "pv_expenses": "Summary!C15",
    "pv_commissions": "Summary!C16",
    "pv_inv_income": "Summary!C17",
    "pv_av_change": "Summary!C18",
    "pv_pols_if": "Summary!C19",
    "pv_net_cf": "Summary!C20",
}

EXPECTED_EDGES = [
    ("duration_mth", "duration_mth_0"), ("duration_mth", "t"),
    ("duration", "duration_mth"),
    ("age", "age_at_entry"), ("age", "duration"),
    ("premium_pp", "premium_type"), ("premium_pp", "duration_mth"),
    ("premium_pp", "premium_pp_input"), ("premium_pp", "policy_term"),
    ("mort_rate", "mort_table"), ("mort_rate", "age"), ("mort_rate", "duration"),
    ("mort_rate_mth", "mort_rate"),
    ("lapse_rate", "duration"),
    ("surr_charge_rate", "has_surr_charge"), ("surr_charge_rate", "surr_charge_table"),
    ("surr_charge_rate", "surr_charge_id"), ("surr_charge_rate", "duration"),
    ("pols_if@0", "pols_if_init"), ("pols_if", "pols_if_bef_decr"),
    ("pols_if", "pols_lapse"), ("pols_if", "pols_death"),
    ("pols_maturity", "duration_mth"), ("pols_maturity", "policy_term"),
    ("pols_maturity", "pols_if"),
    ("pols_if_bef_nb", "pols_if"), ("pols_if_bef_nb", "pols_maturity"),
    ("pols_new_biz", "duration_mth"), ("pols_new_biz", "policy_count"),
    ("pols_if_bef_decr", "pols_if_bef_nb"), ("pols_if_bef_decr", "pols_new_biz"),
    ("pols_death", "pols_if_bef_decr"), ("pols_death", "mort_rate_mth"),
    ("pols_lapse", "pols_if_bef_decr"), ("pols_lapse", "pols_death"),
    ("pols_lapse", "lapse_rate"),
    ("av_pp_bef_prem@0", "av_pp_init"), ("av_pp_bef_prem", "av_pp_bef_inv"),
    ("av_pp_bef_prem", "inv_income_pp"),
    ("prem_to_av_pp", "load_prem_rate"), ("prem_to_av_pp", "premium_pp"),
    ("av_pp_bef_fee", "av_pp_bef_prem"), ("av_pp_bef_fee", "prem_to_av_pp"),
    ("maint_fee_pp", "maint_fee_rate"), ("maint_fee_pp", "av_pp_bef_fee"),
    ("net_amt_at_risk", "sum_assured"), ("net_amt_at_risk", "av_pp_bef_fee"),
    ("coi_rate", "coi_multiple"), ("coi_rate", "mort_rate_mth"),
    ("coi_pp", "coi_rate"), ("coi_pp", "net_amt_at_risk"),
    ("av_pp_bef_inv", "av_pp_bef_fee"), ("av_pp_bef_inv", "maint_fee_pp"),
    ("av_pp_bef_inv", "coi_pp"),
    ("inv_return_mth", "mu"), ("inv_return_mth", "sigma"),
    ("inv_return_mth", "std_norm_rand"), ("inv_return_mth", "scen_id"),
    ("inv_return_mth", "t"),
    ("inv_income_pp", "inv_return_mth"), ("inv_income_pp", "av_pp_bef_inv"),
    ("av_pp_mid_mth", "av_pp_bef_inv"), ("av_pp_mid_mth", "inv_income_pp"),
    ("av_bef_mat", "av_pp_bef_prem"), ("av_bef_mat", "pols_if"),
    ("av_bef_nb", "av_pp_bef_prem"), ("av_bef_nb", "pols_if_bef_nb"),
    ("av_bef_fee", "av_pp_bef_fee"), ("av_bef_fee", "pols_if_bef_decr"),
    ("premiums", "premium_pp"), ("premiums", "pols_if_bef_decr"),
    ("prem_to_av", "prem_to_av_pp"), ("prem_to_av", "pols_if_bef_decr"),
    ("claim_pp_death", "sum_assured"), ("claim_pp_death", "av_pp_mid_mth"),
    ("claims_death", "claim_pp_death"), ("claims_death", "pols_death"),
    ("cfav_death", "av_pp_mid_mth"), ("cfav_death", "pols_death"),
    ("cfav_lapse", "av_pp_mid_mth"), ("cfav_lapse", "pols_lapse"),
    ("surr_charge", "surr_charge_rate"), ("surr_charge", "av_pp_mid_mth"),
    ("surr_charge", "pols_lapse"),
    ("claims_lapse", "cfav_lapse"), ("claims_lapse", "surr_charge"),
    ("cfav_maturity", "av_pp_bef_prem"), ("cfav_maturity", "pols_maturity"),
    ("claims_total", "claims_death"), ("claims_total", "claims_lapse"),
    ("claims_total", "cfav_maturity"),
    ("claims_over_av", "claim_pp_death"), ("claims_over_av", "av_pp_mid_mth"),
    ("claims_over_av", "pols_death"),
    ("coi", "coi_pp"), ("coi", "pols_if_bef_decr"),
    ("maint_fee", "maint_fee_pp"), ("maint_fee", "pols_if_bef_decr"),
    ("inv_income", "inv_income_pp"), ("inv_income", "pols_if"),
    ("inv_income", "pols_death"), ("inv_income", "pols_lapse"),
    ("inflation_factor", "inflation_rate"), ("inflation_factor", "t"),
    ("expenses", "expense_acq"), ("expenses", "pols_new_biz"),
    ("expenses", "pols_if_bef_decr"), ("expenses", "expense_maint"),
    ("expenses", "inflation_factor"),
    ("commissions", "commission_rate"), ("commissions", "premiums"),
    ("av_change", "av_bef_mat"),
    ("net_cf", "premiums"), ("net_cf", "inv_income"), ("net_cf", "claims_total"),
    ("net_cf", "expenses"), ("net_cf", "commissions"), ("net_cf", "av_change"),
    ("margin_expense", "load_prem_rate"), ("margin_expense", "premium_pp"),
    ("margin_expense", "pols_if_bef_decr"), ("margin_expense", "surr_charge"),
    ("margin_expense", "maint_fee"), ("margin_expense", "commissions"),
    ("margin_expense", "expenses"),
    ("margin_mortality", "coi"), ("margin_mortality", "claims_over_av"),
    ("disc_rate_mth", "disc_rate_ann"), ("disc_rate_mth", "t"),
    ("disc_factors", "disc_rate_mth"), ("disc_factors", "t"),
    ("pv_premiums", "premiums"), ("pv_premiums", "disc_factors"),
    ("pv_claims[DEATH]", "claims_death"), ("pv_claims[DEATH]", "disc_factors"),
    ("pv_claims[LAPSE]", "claims_lapse"), ("pv_claims[LAPSE]", "disc_factors"),
    ("pv_claims[MATURITY]", "cfav_maturity"), ("pv_claims[MATURITY]", "disc_factors"),
    ("pv_claims", "pv_claims[DEATH]"), ("pv_claims", "pv_claims[LAPSE]"),
    ("pv_claims", "pv_claims[MATURITY]"),
    ("pv_expenses", "expenses"), ("pv_expenses", "disc_factors"),
    ("pv_commissions", "commissions"), ("pv_commissions", "disc_factors"),
    ("pv_inv_income", "inv_income"), ("pv_inv_income", "disc_factors"),
    ("pv_av_change", "av_change"), ("pv_av_change", "disc_factors"),
    ("pv_pols_if", "pols_if"), ("pv_pols_if", "disc_factors"),
    ("pv_net_cf", "pv_premiums"), ("pv_net_cf", "pv_inv_income"),
    ("pv_net_cf", "pv_claims"), ("pv_net_cf", "pv_expenses"),
    ("pv_net_cf", "pv_commissions"), ("pv_net_cf", "pv_av_change"),
    # model point / spec joins
    ("spec_id", "model_point_table[spec]"), ("spec_id", "point_id"),
    ("age_at_entry", "model_point_table[age]"), ("age_at_entry", "point_id"),
    ("term_input", "model_point_table[term]"), ("term_input", "point_id"),
    ("policy_count", "model_point_table[count]"), ("policy_count", "point_id"),
    ("sum_assured", "model_point_table[sum]"), ("sum_assured", "point_id"),
    ("duration_mth_0", "model_point_table[dur]"), ("duration_mth_0", "point_id"),
    ("premium_pp_input", "model_point_table[prem]"), ("premium_pp_input", "point_id"),
    ("av_pp_init", "model_point_table[avinit]"), ("av_pp_init", "point_id"),
    ("premium_type", "product_spec_table[premtype]"), ("premium_type", "spec_id"),
    ("has_surr_charge", "product_spec_table[hassurr]"), ("has_surr_charge", "spec_id"),
    ("surr_charge_id", "product_spec_table[surrid]"), ("surr_charge_id", "spec_id"),
    ("load_prem_rate", "product_spec_table[loadprem]"), ("load_prem_rate", "spec_id"),
    ("is_wl", "product_spec_table[iswl]"), ("is_wl", "spec_id"),
    ("policy_term", "is_wl"), ("policy_term", "mort_table_last_age"),
    ("policy_term", "age_at_entry"), ("policy_term", "term_input"),
    ("pols_if_init", "duration_mth_0"), ("pols_if_init", "policy_count"),
]

CELL_RE = re.compile(r"(?:'?([^'!]+)'?!)?\$?([A-Z]+)\$?(\d+)")


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def parse_addr(addr):
    m = CELL_RE.fullmatch(addr)
    sheet, col, row = m.groups()
    return sheet, col_to_num(col), int(row)


def range_contains(range_a1, sheet, col, row):
    if "!" not in range_a1:
        return False
    rsheet, rest = range_a1.split("!", 1)
    rsheet = rsheet.strip("'")
    if rsheet != sheet:
        return False
    parts = rest.split(":")
    _, c1, r1 = parse_addr(f"{rsheet}!{parts[0]}")
    if len(parts) == 2:
        _, c2, r2 = parse_addr(f"{rsheet}!{parts[1]}")
    else:
        c2, r2 = c1, r1
    return c1 <= col <= c2 and r1 <= row <= r2


def main():
    con = sqlite3.connect(DB)
    edges = con.execute(
        "SELECT from_address, to_address FROM agent_binding_dependencies").fetchall()

    missing, present, known_gaps = [], [], []
    for dep, prec in EXPECTED_EDGES:
        ds, dc, dr = parse_addr(LOC[dep])
        ps, pc, pr = parse_addr(LOC[prec])
        ok = any(
            range_contains(fa, ds, dc, dr) and range_contains(ta, ps, pc, pr)
            for fa, ta in edges
        )
        if ok:
            present.append(f"{dep} -> {prec}")
        elif (dep, prec) in KNOWN_EXTRACTOR_GAPS:
            known_gaps.append(f"{dep} -> {prec} [{KNOWN_EXTRACTOR_GAPS[(dep, prec)]}]")
        else:
            missing.append(f"{dep} -> {prec}")

    verdict = ("PASS" if not missing and not known_gaps
               else "PASS_WITH_KNOWN_EXTRACTOR_GAPS" if not missing
               else "FAIL")
    report = {
        "workbook": "workbooks/CashValue_SE/CashValue_SE.xlsm",
        "ir_db": "marinade/CashValue_SE.ir.db",
        "date": date.today().isoformat(),
        "expected_edges": len(EXPECTED_EDGES),
        "edges_found": len(present),
        "edges_missing_known_extractor_gaps": known_gaps,
        "edges_missing": missing,
        "verdict": verdict,
    }
    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text(json.dumps(report, indent=1))
    print(f"{verdict}: {len(present)}/{len(EXPECTED_EDGES)} lifelib "
          f"dependencies present in workbook IR")
    for m in missing:
        print("  MISSING:", m)
    return 1 if missing else 0


if __name__ == "__main__":
    sys.exit(main())
