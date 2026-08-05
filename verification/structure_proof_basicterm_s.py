"""Structure proof: the workbook's dependency graph matches lifelib's call graph.

The value proof (verify_basicterm_s.py) shows the numbers agree; this
script shows the *calculation structure* agrees. For every dependency in
the lifelib BasicTerm_S source (cell A calls cell B), we assert the Marinade
IR contains a formula edge between the workbook ranges that implement A
and B.

Expected edges below are transcribed from
models/basiclife/BasicTerm_S/Projection/__init__.py. Where the workbook
deliberately inlines or re-homes a lifelib cell, the mapping notes it:
  - claim_pp(t) == sum_assured() -> claims links to SumAssured directly
  - model_point() is not a cell   -> lookups link to Model_Points + PointID
  - disc_rate_mth uses duration   -> lifelib's inline t//12 is the duration column
  - lapse_rate constants          -> exposed as Assumptions inputs (extra precedents allowed)

Exit non-zero if any expected edge is missing from the IR.
"""

import json
import re
import sqlite3
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
DB = REPO / "marinade" / "BasicTerm_S.ir.db"
REPORT = REPO / "verification" / "reports" / "basicterm_s_structure_proof.json"

# lifelib cell -> representative workbook cell (any cell inside the
# binding that implements it)
LOC = {
    # Projection engine columns (row 10 = t=6, an ordinary in-force row)
    "t": "Projection!A10",
    "duration": "Projection!B10",
    "age": "Projection!C10",
    "mort_rate": "Projection!D10",
    "mort_rate_mth": "Projection!E10",
    "lapse_rate": "Projection!F10",
    "pols_if": "Projection!G10",
    "pols_death": "Projection!H10",
    "pols_lapse": "Projection!I10",
    "pols_maturity": "Projection!J10",
    "premiums": "Projection!K10",
    "claims": "Projection!L10",
    "commissions": "Projection!M10",
    "inflation_factor": "Projection!N10",
    "expenses": "Projection!O10",
    "net_cf": "Projection!P10",
    "disc_rate_mth": "Projection!Q10",
    "disc_factors": "Projection!R10",
    # run control (model point selection lives on Control, not Assumptions)
    "point_id": "Control!C4",
    "age_at_entry": "Control!C8",
    "policy_term": "Control!C10",
    "sum_assured": "Control!C11",
    "pols_if_init": "Control!C12",
    # basis assumptions
    "expense_acq": "Assumptions!C4",
    "expense_maint": "Assumptions!C5",
    "inflation_rate": "Assumptions!C6",
    "loading_prem": "Assumptions!C7",
    # data tables
    "mort_table": "Mortality!C10",
    "disc_rate_ann": "Discount!B10",
    # model_point_table is one lifelib table but four workbook columns; map
    # each lookup to the column it actually reads
    "model_point_table[age]": "Model_Points!B10",
    "model_point_table[term]": "Model_Points!D10",
    "model_point_table[sum]": "Model_Points!F10",
    # summary results
    "pv_pols_if": "Summary!C10",
    "pv_claims": "Summary!C11",
    "net_premium_pp": "Summary!C12",
    "premium_pp": "Summary!C13",
    "pv_premiums": "Summary!C17",
    "pv_expenses": "Summary!C19",
    "pv_commissions": "Summary!C20",
    "pv_net_cf": "Summary!C21",
}

# (dependent, precedent) — transcribed from the lifelib source, with the
# documented inline substitutions applied.
EXPECTED_EDGES = [
    ("duration", "t"),
    ("age", "age_at_entry"), ("age", "duration"),
    ("mort_rate", "mort_table"), ("mort_rate", "age"), ("mort_rate", "duration"),
    ("mort_rate_mth", "mort_rate"),
    ("lapse_rate", "duration"),
    ("pols_if", "pols_if_init"), ("pols_if", "policy_term"),
    ("pols_if", "pols_lapse"), ("pols_if", "pols_death"), ("pols_if", "pols_maturity"),
    ("pols_death", "pols_if"), ("pols_death", "mort_rate_mth"),
    ("pols_lapse", "pols_if"), ("pols_lapse", "pols_death"), ("pols_lapse", "lapse_rate"),
    ("pols_maturity", "policy_term"), ("pols_maturity", "pols_if"),
    ("pols_maturity", "pols_lapse"), ("pols_maturity", "pols_death"),
    ("claims", "sum_assured"), ("claims", "pols_death"),  # claim_pp inlined
    ("premiums", "premium_pp"), ("premiums", "pols_if"),
    ("commissions", "premiums"), ("commissions", "duration"),
    ("inflation_factor", "inflation_rate"), ("inflation_factor", "t"),
    ("expenses", "pols_if"), ("expenses", "expense_maint"),
    ("expenses", "inflation_factor"), ("expenses", "expense_acq"),
    ("net_cf", "premiums"), ("net_cf", "claims"),
    ("net_cf", "expenses"), ("net_cf", "commissions"),
    ("disc_rate_mth", "disc_rate_ann"), ("disc_rate_mth", "duration"),  # t//12
    ("disc_factors", "disc_rate_mth"), ("disc_factors", "t"),
    ("pv_premiums", "premiums"), ("pv_premiums", "disc_factors"),
    ("pv_claims", "claims"), ("pv_claims", "disc_factors"),
    ("pv_expenses", "expenses"), ("pv_expenses", "disc_factors"),
    ("pv_commissions", "commissions"), ("pv_commissions", "disc_factors"),
    ("pv_pols_if", "pols_if"), ("pv_pols_if", "disc_factors"),
    ("pv_net_cf", "pv_premiums"), ("pv_net_cf", "pv_claims"),
    ("pv_net_cf", "pv_expenses"), ("pv_net_cf", "pv_commissions"),
    ("net_premium_pp", "pv_claims"), ("net_premium_pp", "pv_pols_if"),
    ("premium_pp", "loading_prem"), ("premium_pp", "net_premium_pp"),
    ("age_at_entry", "model_point_table[age]"), ("age_at_entry", "point_id"),
    ("policy_term", "model_point_table[term]"), ("policy_term", "point_id"),
    ("sum_assured", "model_point_table[sum]"), ("sum_assured", "point_id"),
]

# Expected edges that Marinade's current extractor is known not to emit
# (verified manually against the stored formula text; see marinade-notes.md).
# Findings A1 and A2 were fixed by making lookup functions transparent for
# argument-ref extraction (xl-marinade ≥ 0.1.0), so this is empty — it
# stays as the mechanism for classifying any future extractor gap.
KNOWN_EXTRACTOR_GAPS: dict[tuple[str, str], str] = {}

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
    """range_a1 like 'Projection!A4:A244' or 'Assumptions!C4'."""
    if "!" in range_a1:
        rsheet, rest = range_a1.split("!", 1)
        rsheet = rsheet.strip("'")
    else:
        return False
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
        "SELECT from_address, to_address, from_label, to_label "
        "FROM agent_binding_dependencies").fetchall()

    def binding_of(lifelib_cell):
        sheet, col, row = parse_addr(LOC[lifelib_cell])
        found = set()
        for fa, ta, fl, tl in edges:
            if range_contains(fa, sheet, col, row):
                found.add(("from", fa))
            if range_contains(ta, sheet, col, row):
                found.add(("to", ta))
        return found

    missing, present, known_gaps = [], [], []
    for dep, prec in EXPECTED_EDGES:
        ds, dc, dr = parse_addr(LOC[dep])
        ps, pc, pr = parse_addr(LOC[prec])
        ok = any(
            range_contains(fa, ds, dc, dr) and range_contains(ta, ps, pc, pr)
            for fa, ta, _fl, _tl in edges
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
        "workbook": "workbooks/BasicTerm_S/BasicTerm_S.xlsm",
        "ir_db": "marinade/BasicTerm_S.ir.db",
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
          f"dependencies present in workbook IR "
          f"({len(known_gaps)} missing edges attributed to known Marinade gaps)")
    for m in missing:
        print("  MISSING:", m)
    for g in known_gaps:
        print("  KNOWN GAP:", g)
    return 1 if missing else 0


if __name__ == "__main__":
    sys.exit(main())
