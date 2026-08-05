"""Structure proof: BasicTerm_SE workbook dependency graph vs lifelib call graph.

Same method as structure_proof_basicterm_s.py; expected edges transcribed
from models/basiclife/BasicTerm_SE/Projection/__init__.py. Inline notes:
  - pols_if(t) == pols_if_at(t, "BEF_MAT") — one column (I)
  - claim_pp inlined into claims (sum_assured)
  - duration_mth's recursion is closed-form in Excel (Duration0 + t)
  - premium_pp inlines the premium_table lookup so its edges stay direct
  - is_active guards add Excel-only edges (extras are fine; the proof
    asserts lifelib edges exist, not that no others do)
"""

import json
import re
import sqlite3
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
DB = REPO / "marinade" / "BasicTerm_SE.ir.db"
REPORT = REPO / "verification" / "reports" / "basicterm_se_structure_proof.json"

KNOWN_EXTRACTOR_GAPS: dict = {}

LOC = {
    # Projection engine columns (row 10 = t = 6)
    "t": "Projection!A10",
    "duration_mth": "Projection!B10",
    "duration": "Projection!C10",
    "age": "Projection!D10",
    "is_active": "Projection!E10",
    "mort_rate": "Projection!F10",
    "mort_rate_mth": "Projection!G10",
    "lapse_rate": "Projection!H10",
    "pols_if": "Projection!I10",            # pols_if_at(t, "BEF_MAT")
    "pols_maturity": "Projection!J10",
    "pols_if_bef_nb": "Projection!K10",     # pols_if_at(t, "BEF_NB")
    "pols_new_biz": "Projection!L10",
    "pols_if_bef_decr": "Projection!M10",   # pols_if_at(t, "BEF_DECR")
    "pols_death": "Projection!N10",
    "pols_lapse": "Projection!O10",
    "premiums": "Projection!P10",
    "claims": "Projection!Q10",
    "commissions": "Projection!R10",
    "inflation_factor": "Projection!S10",
    "expenses": "Projection!T10",
    "net_cf": "Projection!U10",
    "disc_rate_mth": "Projection!V10",
    "disc_factors": "Projection!W10",
    # run control
    "point_id": "Control!C4",
    "age_at_entry": "Control!C8",
    "policy_term": "Control!C10",
    "sum_assured": "Control!C11",
    "policy_count": "Control!C12",
    "duration_mth_0": "Control!C13",
    "pols_if_init": "Control!C14",
    # basis assumptions
    "expense_acq": "Assumptions!C4",
    "expense_maint": "Assumptions!C5",
    "inflation_rate": "Assumptions!C6",
    # data tables
    "mort_table": "Mortality!C10",
    "disc_rate_ann": "Discount!B10",
    "premium_table": "Premium_Rates!C10",
    "model_point_table[age]": "Model_Points!B10",
    "model_point_table[term]": "Model_Points!D10",
    "model_point_table[count]": "Model_Points!E10",
    "model_point_table[sum]": "Model_Points!F10",
    "model_point_table[dur]": "Model_Points!G10",
    # summary
    "premium_pp": "Summary!C11",
    "net_premium_pp": "Summary!C12",
    "pv_pols_if": "Summary!C16",
    "pv_premiums": "Summary!C17",
    "pv_claims": "Summary!C18",
    "pv_expenses": "Summary!C19",
    "pv_commissions": "Summary!C20",
    "pv_net_cf": "Summary!C21",
}

EXPECTED_EDGES = [
    ("duration_mth", "duration_mth_0"), ("duration_mth", "t"),
    ("duration", "duration_mth"),
    ("age", "age_at_entry"), ("age", "duration"),
    ("is_active", "duration_mth"), ("is_active", "policy_term"),
    ("mort_rate", "mort_table"), ("mort_rate", "age"), ("mort_rate", "duration"),
    ("mort_rate_mth", "mort_rate"),
    ("lapse_rate", "duration"),
    ("pols_if", "pols_if_init"), ("pols_if", "is_active"),
    ("pols_if", "pols_if_bef_decr"), ("pols_if", "pols_lapse"), ("pols_if", "pols_death"),
    ("pols_maturity", "duration_mth"), ("pols_maturity", "policy_term"),
    ("pols_maturity", "pols_if"),
    ("pols_if_bef_nb", "pols_if"), ("pols_if_bef_nb", "pols_maturity"),
    ("pols_if_bef_nb", "is_active"),
    ("pols_new_biz", "duration_mth"), ("pols_new_biz", "policy_count"),
    ("pols_if_bef_decr", "pols_if_bef_nb"), ("pols_if_bef_decr", "pols_new_biz"),
    ("pols_if_bef_decr", "is_active"),
    ("pols_death", "pols_if_bef_decr"), ("pols_death", "mort_rate_mth"),
    ("pols_death", "is_active"),
    ("pols_lapse", "pols_if_bef_decr"), ("pols_lapse", "pols_death"),
    ("pols_lapse", "lapse_rate"), ("pols_lapse", "is_active"),
    ("premiums", "premium_pp"), ("premiums", "pols_if_bef_decr"),
    ("claims", "sum_assured"), ("claims", "pols_death"),
    ("commissions", "premiums"), ("commissions", "duration"),
    ("inflation_factor", "inflation_rate"), ("inflation_factor", "t"),
    ("expenses", "expense_acq"), ("expenses", "pols_new_biz"),
    ("expenses", "pols_if_bef_decr"), ("expenses", "expense_maint"),
    ("expenses", "inflation_factor"),
    ("net_cf", "premiums"), ("net_cf", "claims"),
    ("net_cf", "expenses"), ("net_cf", "commissions"),
    ("disc_rate_mth", "disc_rate_ann"), ("disc_rate_mth", "t"),
    ("disc_factors", "disc_rate_mth"), ("disc_factors", "t"),
    ("pv_premiums", "premiums"), ("pv_premiums", "disc_factors"),
    ("pv_claims", "claims"), ("pv_claims", "disc_factors"),
    ("pv_expenses", "expenses"), ("pv_expenses", "disc_factors"),
    ("pv_commissions", "commissions"), ("pv_commissions", "disc_factors"),
    ("pv_pols_if", "pols_if"), ("pv_pols_if", "disc_factors"),
    ("pv_net_cf", "pv_premiums"), ("pv_net_cf", "pv_claims"),
    ("pv_net_cf", "pv_expenses"), ("pv_net_cf", "pv_commissions"),
    ("net_premium_pp", "pv_claims"), ("net_premium_pp", "pv_pols_if"),
    ("premium_pp", "sum_assured"), ("premium_pp", "premium_table"),
    ("premium_pp", "age_at_entry"), ("premium_pp", "policy_term"),
    ("age_at_entry", "model_point_table[age]"), ("age_at_entry", "point_id"),
    ("policy_term", "model_point_table[term]"), ("policy_term", "point_id"),
    ("sum_assured", "model_point_table[sum]"), ("sum_assured", "point_id"),
    ("policy_count", "model_point_table[count]"), ("policy_count", "point_id"),
    ("duration_mth_0", "model_point_table[dur]"), ("duration_mth_0", "point_id"),
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
        "workbook": "workbooks/BasicTerm_SE/BasicTerm_SE.xlsm",
        "ir_db": "marinade/BasicTerm_SE.ir.db",
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
    for g in known_gaps:
        print("  KNOWN GAP:", g)
    return 1 if missing else 0


if __name__ == "__main__":
    sys.exit(main())
