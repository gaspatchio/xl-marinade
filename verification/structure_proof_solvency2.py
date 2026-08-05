"""Structure proof: Solvency2 workbook dependency graph vs lifelib call graph.

The largest transcription so far — the SCR_life evaluated subgraph across
Policy (premium/reserve rates via LifeTable commutation), Assumptions
(lookup cascade), BaseProj+PV (the engine, asserted on Proj_Base), the four
Override spaces (asserted on their stress sheets), and the SCR aggregation.

Mapping notes:
  - Named scalar cells are located from the IR's defined_names table, so the
    proof does not hardcode Policy_Basis row numbers.
  - Engine columns are asserted on Proj_Base at row 30 (t = 23); stress
    override cells on their own sheets. Shared pol/asmp/scen columns live on
    Policy_Sched (lifelib holds them once per policy too).
  - lifelib x/n/m arguments (IssueAge/PolicyTerm) reach the commutation
    functions through the anchor cells DxX/MxX/NxX/Dx_xn/Mx_xn/Nx_xn.
  - Shock factors reach the stress columns through each sheet's parameter
    cells (row 3), which read the FactorData table — both hops asserted.
  - disab/rev/cat Life charges are constants in the workbook (no lifelib
    override => stressed projection == base => charge 0); their lifelib
    edges are intentionally not transcribed (documented economy, not an
    Marinade gap).
"""

import json
import re
import sqlite3
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
DB = REPO / "marinade" / "Solvency2.ir.db"
REPORT = REPO / "verification" / "reports" / "solvency2_structure_proof.json"

KNOWN_EXTRACTOR_GAPS: dict = {}

ROW = 30           # representative row on time sheets (t = 23)
CROW = 34          # representative row on Commutation (x = 30)

# columns: machine name -> (sheet, col letter); engine on Proj_Base
S = "Policy_Sched"
COLS = {
    "AttAge": (S, "B"), "BaseMortRate": (S, "C"), "MortFactor": (S, "D"),
    "SurrRateBase": (S, "E"), "ReserveNLP": (S, "F"), "SurrChargeT": (S, "G"),
    "CashValueRate": (S, "H"), "SizeBenefitSurr": (S, "I"),
    "SizeExpsCommRen": (S, "J"), "DiscRate": (S, "K"),
    "qx": ("Commutation", "B"), "lx": ("Commutation", "C"),
    "dx": ("Commutation", "D"), "Dx": ("Commutation", "E"),
    "Cx": ("Commutation", "F"), "Nx": ("Commutation", "G"),
    "Mx": ("Commutation", "H"),
    "InflFactor": ("Proj_Base", "B"), "MortRateFactor": ("Proj_Base", "C"),
    "SurrRate": ("Proj_Base", "D"), "PolsIF_End": ("Proj_Base", "E"),
    "PolsMaturity": ("Proj_Base", "F"), "PolsIF_Beg": ("Proj_Base", "G"),
    "PolsNewBiz": ("Proj_Base", "H"), "PolsSurrMass": ("Proj_Base", "I"),
    "PolsIF_Beg1": ("Proj_Base", "J"), "PolsDeath": ("Proj_Base", "K"),
    "PolsSurr": ("Proj_Base", "L"), "SizeExpsMaint": ("Proj_Base", "M"),
    "PremIncome": ("Proj_Base", "N"), "BenefitDeath": ("Proj_Base", "O"),
    "BenefitSurr": ("Proj_Base", "P"), "BenefitTotal": ("Proj_Base", "Q"),
    "ExpsAcq": ("Proj_Base", "R"), "ExpsCommInit": ("Proj_Base", "S"),
    "ExpsCommRen": ("Proj_Base", "T"), "ExpsMaint": ("Proj_Base", "U"),
    "ExpsTotal": ("Proj_Base", "V"), "PV_PremIncome": ("Proj_Base", "W"),
    "PV_BenefitTotal": ("Proj_Base", "X"), "PV_ExpsTotal": ("Proj_Base", "Y"),
    "PV_NetCashflow": ("Proj_Base", "Z"), "PV_NCFCheck": ("Proj_Base", "AA"),
    # stress override cells on their own sheets
    "Mort.MortRateFactor": ("Proj_Mort", "C"),
    "Mort.PolsDeath": ("Proj_Mort", "K"),
    "LapseUp.SurrRate": ("Proj_LapseUp", "D"),
    "LapseUp.PolsSurr": ("Proj_LapseUp", "L"),
    "Mass.PolsSurrMass": ("Proj_LapseMass", "I"),
    "Mass.PolsIF_Beg1": ("Proj_LapseMass", "J"),
    "Mass.BenefitSurr": ("Proj_LapseMass", "P"),
    "Mass.PolsIF_Beg": ("Proj_LapseMass", "G"),
    "Exps.InflFactor": ("Proj_Exps", "B"),
    "Exps.SizeExpsMaint": ("Proj_Exps", "M"),
}
# sheet parameter cells (row 3)
PARAMS = {
    "Mort.FactorMain": ("Proj_Mort", "E", 3),
    "LapseUp.FactorMain": ("Proj_LapseUp", "E", 3),
    "LapseUp.FactorLimit": ("Proj_LapseUp", "F", 3),
    "Mass.FactorMain": ("Proj_LapseMass", "E", 3),
    "Exps.FactorMain": ("Proj_Exps", "E", 3),
    "Exps.InflShock": ("Proj_Exps", "G", 3),
}

NAMES = [
    "PolicyID", "ScenID", "T0", "Product", "PolicyType", "Gen", "Sex",
    "IssueAge", "PremFreq", "PolicyTerm", "PolicyCount", "SumAssured",
    "BaseMortTable", "MortFactorTable", "SurrTable", "CnsmpTax",
    "CommInitPrem", "CommRenPrem", "CommRenTerm", "ExpsAcqAnnPrem",
    "ExpsAcqPol", "ExpsAcqSA", "ExpsMaintAnnPrem", "ExpsMaintPol",
    "ExpsMaintSA", "InflRate", "IntRatePrem", "TableIDPrem",
    "LoadAcqP1", "LoadAcqP2", "LoadMaintPremP1", "LoadMaintPremP2",
    "LoadMaintSA", "LoadMaintSA2", "SurrP1", "SurrP2",
    "BaseMortCol", "PremMortCol", "MortFactorCol", "SurrCol",
    "LastAge", "LastT", "DxX", "MxX", "NxX", "Dx_xn", "Mx_xn", "Nx_xn",
    "Axn", "Exn", "AnnDue_xm_k", "AnnDue_xn_1", "AnnDue_gamma2",
    "LoadAcqSA", "LoadMaintPrem", "LoadWaiver", "InitSurrCharge",
    "GrossPremRate", "NetPremRatePrem", "AnnPremRate", "SizePremium",
    "SizeAnnPrem", "SizeExpsAcq0", "SizeExpsCommInit0",
    "NAV_Base", "NAV_Mort", "NAV_Longev", "NAV_Exps", "NAV_LapseUp",
    "NAV_LapseDown", "NAV_LapseMass", "LapseUp_Risk", "LapseDown_Risk",
    "LapseMass_Risk", "Life_Mort", "Life_Longev", "Life_Lapse", "Life_Exps",
    "SCR_Life",
]
# named ranges used as precedents (any contained cell counts)
RANGES = {
    "MP_Product": "MP_Product", "MP_ID": "MP_ID",
    "AsmpKeys": "AsmpKeys", "AsmpVals": "AsmpVals",
    "SpecKeys": "SpecKeys", "SpecVals": "SpecVals",
    "MortGrid": "MortGrid", "AsmpTblGrid": "AsmpTblGrid",
    "AsmpTblNames": "AsmpTblNames", "ScenGrid": "ScenGrid",
    "FactorVals": "FactorVals", "Comm_Dx": "Comm_Dx", "Comm_Nx": "Comm_Nx",
    "Comm_Mx": "Comm_Mx", "CorrGrid": "CorrGrid", "ProdGrid": "ProdGrid",
    "LifeVec": "LifeVec",
    "PVNCF_Base": "PVNCF_Base", "PVNCF_Mort": "PVNCF_Mort",
    "PVNCF_Longev": "PVNCF_Longev", "PVNCF_Exps": "PVNCF_Exps",
    "PVNCF_LapseUp": "PVNCF_LapseUp", "PVNCF_LapseDown": "PVNCF_LapseDown",
    "PVNCF_LapseMass": "PVNCF_LapseMass",
}

LOOKUPS_A = ["BaseMortTable", "MortFactorTable", "SurrTable",
             "CommInitPrem", "CommRenPrem", "CommRenTerm", "ExpsAcqAnnPrem",
             "ExpsAcqPol", "ExpsAcqSA", "ExpsMaintAnnPrem", "ExpsMaintPol",
             "ExpsMaintSA"]
# CnsmpTax and InflRate: lifelib calls AsmpLookup(item) with NO product
# argument (direct global lookup) - the workbook mirrors that, so their
# only precedents are the lookup table itself
LOOKUPS_A_GLOBAL = ["CnsmpTax", "InflRate"]
LOOKUPS_S = ["IntRatePrem", "TableIDPrem", "LoadAcqP1", "LoadAcqP2",
             "LoadMaintPremP1", "LoadMaintPremP2", "LoadMaintSA",
             "LoadMaintSA2", "SurrP1", "SurrP2"]


def expected_edges():
    E = []
    # --- policy attributes from the model point file
    for nm in ["Product", "PolicyType", "Gen", "Sex", "IssueAge", "PremFreq",
               "PolicyTerm", "PolicyCount", "SumAssured"]:
        E += [(nm, "MP_ID"), (nm, "PolicyID")]
    E.append(("Product", "MP_Product"))
    # --- lookup cascade: every resolved item reads its table and the match args
    for nm in LOOKUPS_A:
        E += [(nm, "AsmpKeys"), (nm, "AsmpVals"), (nm, "Product")]
    for nm in LOOKUPS_A_GLOBAL:
        E += [(nm, "AsmpKeys"), (nm, "AsmpVals")]
    for nm in LOOKUPS_S:
        E += [(nm, "SpecKeys"), (nm, "SpecVals"), (nm, "Product")]
    # --- derived policy quantities
    E += [("BaseMortCol", "BaseMortTable"), ("BaseMortCol", "Sex"),
          ("PremMortCol", "TableIDPrem"), ("PremMortCol", "Sex"),
          ("MortFactorCol", "MortFactorTable"), ("MortFactorCol", "AsmpTblNames"),
          ("SurrCol", "SurrTable"), ("SurrCol", "AsmpTblNames"),
          ("LastAge", "MortGrid"), ("LastAge", "BaseMortCol"),
          ("LastT", "LastAge"), ("LastT", "IssueAge"), ("LastT", "PolicyTerm"),
          ("DxX", "Comm_Dx"), ("DxX", "IssueAge"),
          ("MxX", "Comm_Mx"), ("MxX", "IssueAge"),
          ("NxX", "Comm_Nx"), ("NxX", "IssueAge"),
          ("Dx_xn", "Comm_Dx"), ("Dx_xn", "PolicyTerm"),
          ("Mx_xn", "Comm_Mx"), ("Mx_xn", "PolicyTerm"),
          ("Nx_xn", "Comm_Nx"), ("Nx_xn", "PolicyTerm"),
          ("Axn", "MxX"), ("Axn", "Mx_xn"), ("Axn", "DxX"),
          ("Exn", "Dx_xn"), ("Exn", "DxX"),
          ("AnnDue_xm_k", "NxX"), ("AnnDue_xm_k", "Nx_xn"),
          ("AnnDue_xm_k", "DxX"), ("AnnDue_xm_k", "Dx_xn"),
          ("AnnDue_xm_k", "PremFreq"),
          ("AnnDue_xn_1", "NxX"), ("AnnDue_xn_1", "Nx_xn"), ("AnnDue_xn_1", "DxX"),
          ("AnnDue_gamma2", "Nx_xn"), ("AnnDue_gamma2", "DxX"),
          ("LoadAcqSA", "LoadAcqP1"), ("LoadAcqSA", "LoadAcqP2"),
          ("LoadAcqSA", "PolicyTerm"),
          ("LoadMaintPrem", "LoadMaintPremP1"), ("LoadMaintPrem", "LoadMaintPremP2"),
          ("LoadMaintPrem", "PolicyTerm"),
          ("LoadWaiver", "PolicyTerm"),
          ("InitSurrCharge", "SurrP1"), ("InitSurrCharge", "SurrP2"),
          ("InitSurrCharge", "PolicyTerm")]
    E += [("GrossPremRate", p) for p in
          ["Product", "Exn", "Axn", "LoadAcqSA", "LoadMaintSA", "AnnDue_xm_k",
           "LoadMaintSA2", "AnnDue_gamma2", "LoadMaintPrem", "LoadWaiver",
           "PremFreq"]]
    E += [("NetPremRatePrem", p) for p in
          ["Axn", "LoadMaintSA2", "AnnDue_gamma2", "AnnDue_xn_1"]]
    E += [("AnnPremRate", "GrossPremRate"), ("AnnPremRate", "PremFreq"),
          ("SizePremium", "SumAssured"), ("SizePremium", "GrossPremRate"),
          ("SizePremium", "PremFreq"),
          ("SizeAnnPrem", "SumAssured"), ("SizeAnnPrem", "AnnPremRate"),
          ("SizeExpsAcq0", "SizeAnnPrem"), ("SizeExpsAcq0", "ExpsAcqAnnPrem"),
          ("SizeExpsAcq0", "SumAssured"), ("SizeExpsAcq0", "ExpsAcqSA"),
          ("SizeExpsAcq0", "ExpsAcqPol"),
          ("SizeExpsCommInit0", "SizePremium"), ("SizeExpsCommInit0", "CommInitPrem"),
          ("SizeExpsCommInit0", "CnsmpTax")]
    # --- commutation (LifeTable)
    E += [("qx", "MortGrid"), ("qx", "PremMortCol"),
          ("dx", "lx"), ("dx", "qx"), ("lx", "dx"),
          ("Dx", "lx"), ("Dx", "IntRatePrem"),
          ("Cx", "dx"), ("Cx", "IntRatePrem"),
          ("Nx", "Dx"), ("Mx", "Cx"), ("Mx", "Dx")]
    # --- shared policy schedule
    E += [("AttAge", "IssueAge"),
          ("BaseMortRate", "MortGrid"), ("BaseMortRate", "BaseMortCol"),
          ("BaseMortRate", "AttAge"), ("BaseMortRate", "LastT"),
          ("MortFactor", "AsmpTblGrid"), ("MortFactor", "MortFactorCol"),
          ("SurrRateBase", "AsmpTblGrid"), ("SurrRateBase", "SurrCol"),
          ("ReserveNLP", "Comm_Mx"), ("ReserveNLP", "Comm_Nx"),
          ("ReserveNLP", "Comm_Dx"), ("ReserveNLP", "Mx_xn"),
          ("ReserveNLP", "Nx_xn"), ("ReserveNLP", "LoadMaintSA2"),
          ("ReserveNLP", "NetPremRatePrem"), ("ReserveNLP", "AttAge"),
          ("ReserveNLP", "PolicyTerm"),
          ("SurrChargeT", "InitSurrCharge"), ("SurrChargeT", "PolicyTerm"),
          ("CashValueRate", "ReserveNLP"), ("CashValueRate", "SurrChargeT"),
          ("SizeBenefitSurr", "SumAssured"), ("SizeBenefitSurr", "CashValueRate"),
          ("SizeExpsCommRen", "CommRenTerm"), ("SizeExpsCommRen", "SizePremium"),
          ("SizeExpsCommRen", "CommRenPrem"), ("SizeExpsCommRen", "CnsmpTax"),
          ("DiscRate", "ScenGrid"), ("DiscRate", "ScenID")]
    # --- the engine, asserted on Proj_Base
    E += [("InflFactor", "InflRate"),
          ("PolsIF_End", "PolsIF_Beg1"), ("PolsIF_End", "PolsDeath"),
          ("PolsIF_End", "PolsSurr"),
          ("PolsMaturity", "PolicyTerm"), ("PolsMaturity", "PolsIF_End"),
          ("PolsIF_Beg", "PolsIF_End"), ("PolsIF_Beg", "PolsMaturity"),
          ("PolsNewBiz", "PolicyCount"),
          ("PolsIF_Beg1", "PolsIF_Beg"), ("PolsIF_Beg1", "PolsNewBiz"),
          ("PolsIF_Beg1", "PolsSurrMass"),
          ("PolsDeath", "PolsIF_Beg1"), ("PolsDeath", "BaseMortRate"),
          ("PolsDeath", "MortFactor"), ("PolsDeath", "MortRateFactor"),
          ("PolsSurr", "PolsIF_Beg1"), ("PolsSurr", "SurrRate"),
          ("SurrRate", "SurrRateBase"),
          ("SizeExpsMaint", "SizeAnnPrem"), ("SizeExpsMaint", "ExpsMaintAnnPrem"),
          ("SizeExpsMaint", "SumAssured"), ("SizeExpsMaint", "ExpsMaintSA"),
          ("SizeExpsMaint", "ExpsMaintPol"), ("SizeExpsMaint", "InflFactor"),
          ("PremIncome", "SizePremium"), ("PremIncome", "PolsIF_Beg1"),
          ("BenefitDeath", "SumAssured"), ("BenefitDeath", "PolsDeath"),
          ("BenefitSurr", "SizeBenefitSurr"), ("BenefitSurr", "PolsSurr"),
          ("BenefitSurr", "PolsSurrMass"),
          ("BenefitTotal", "BenefitDeath"), ("BenefitTotal", "BenefitSurr"),
          ("ExpsAcq", "SizeExpsAcq0"), ("ExpsAcq", "PolsNewBiz"),
          ("ExpsCommInit", "SizeExpsCommInit0"), ("ExpsCommInit", "PolsIF_Beg1"),
          ("ExpsCommRen", "SizeExpsCommRen"), ("ExpsCommRen", "PolsIF_Beg1"),
          ("ExpsMaint", "SizeExpsMaint"), ("ExpsMaint", "PolsIF_Beg1"),
          ("ExpsTotal", "ExpsCommInit"), ("ExpsTotal", "ExpsCommRen"),
          ("ExpsTotal", "ExpsAcq"), ("ExpsTotal", "ExpsMaint"),
          ("PV_PremIncome", "PremIncome"), ("PV_PremIncome", "DiscRate"),
          ("PV_BenefitTotal", "BenefitTotal"), ("PV_BenefitTotal", "DiscRate"),
          ("PV_ExpsTotal", "ExpsTotal"), ("PV_ExpsTotal", "DiscRate"),
          ("PV_NetCashflow", "PV_PremIncome"), ("PV_NetCashflow", "PV_ExpsTotal"),
          ("PV_NetCashflow", "PV_BenefitTotal"),
          ("PV_NCFCheck", "PremIncome"), ("PV_NCFCheck", "ExpsTotal"),
          ("PV_NCFCheck", "BenefitTotal"), ("PV_NCFCheck", "DiscRate"),
          ("PV_NCFCheck", "PV_NetCashflow")]
    # --- override cells on their stress sheets
    E += [("Mort.MortRateFactor", "T0"), ("Mort.MortRateFactor", "AttAge"),
          ("Mort.MortRateFactor", "LastAge"),
          ("Mort.MortRateFactor", "Mort.FactorMain"),
          ("Mort.FactorMain", "FactorVals"),
          ("Mort.PolsDeath", "Mort.MortRateFactor"),
          ("LapseUp.SurrRate", "SurrRateBase"), ("LapseUp.SurrRate", "T0"),
          ("LapseUp.SurrRate", "LapseUp.FactorMain"),
          ("LapseUp.SurrRate", "LapseUp.FactorLimit"),
          ("LapseUp.FactorMain", "FactorVals"),
          ("LapseUp.PolsSurr", "LapseUp.SurrRate"),
          ("Mass.PolsSurrMass", "T0"), ("Mass.PolsSurrMass", "Mass.PolsIF_Beg"),
          ("Mass.PolsSurrMass", "Mass.FactorMain"),
          ("Mass.FactorMain", "FactorVals"),
          ("Mass.PolsIF_Beg1", "Mass.PolsSurrMass"),
          ("Mass.BenefitSurr", "Mass.PolsSurrMass"),
          ("Exps.InflFactor", "InflRate"), ("Exps.InflFactor", "T0"),
          ("Exps.InflFactor", "Exps.InflShock"),
          ("Exps.InflShock", "FactorVals"),
          ("Exps.SizeExpsMaint", "Exps.InflFactor"),
          ("Exps.SizeExpsMaint", "Exps.FactorMain"),
          ("Exps.FactorMain", "FactorVals")]
    # --- SCR aggregation
    for sfx in ["Base", "Mort", "Longev", "Exps", "LapseUp", "LapseDown", "LapseMass"]:
        E += [(f"NAV_{sfx}", f"PVNCF_{sfx}"), (f"NAV_{sfx}", "T0")]
    E += [("LapseUp_Risk", "NAV_Base"), ("LapseUp_Risk", "NAV_LapseUp"),
          ("LapseDown_Risk", "NAV_Base"), ("LapseDown_Risk", "NAV_LapseDown"),
          ("LapseMass_Risk", "NAV_Base"), ("LapseMass_Risk", "NAV_LapseMass"),
          ("Life_Mort", "NAV_Base"), ("Life_Mort", "NAV_Mort"),
          ("Life_Longev", "NAV_Base"), ("Life_Longev", "NAV_Longev"),
          ("Life_Lapse", "LapseUp_Risk"), ("Life_Lapse", "LapseDown_Risk"),
          ("Life_Lapse", "LapseMass_Risk"),
          ("Life_Exps", "NAV_Base"), ("Life_Exps", "NAV_Exps"),
          ("ProdGrid", "LifeVec"), ("ProdGrid", "CorrGrid"),
          ("SCR_Life", "ProdGrid")]
    return E


CELL_RE = re.compile(r"(?:'?([^'!]+)'?!)?\$?([A-Z]+)\$?(\d+)")


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def parse_range(a1):
    """'Sheet!$A$4:$B$10' -> (sheet, c1, r1, c2, r2)."""
    sheet, rest = a1.split("!", 1)
    sheet = sheet.strip("'")
    parts = rest.split(":")
    m1 = CELL_RE.fullmatch(parts[0]) or CELL_RE.fullmatch(f"{sheet}!{parts[0]}")
    _s, c1, r1 = None, col_to_num(m1.group(2)), int(m1.group(3))
    if len(parts) == 2:
        m2 = CELL_RE.fullmatch(parts[1])
        c2, r2 = col_to_num(m2.group(2)), int(m2.group(3))
    else:
        c2, r2 = c1, r1
    return sheet, c1, r1, c2, r2


def main():
    con = sqlite3.connect(DB)
    # resolve every named cell/range to (sheet, c1, r1, c2, r2)
    loc = {}
    for name, dests in con.execute("SELECT name, destinations FROM defined_names"):
        refs = json.loads(dests)
        if refs:
            loc[name] = parse_range(refs[0])
    missing_names = [n for n in NAMES + list(RANGES) if n not in loc]
    if missing_names:
        sys.exit(f"names missing from IR defined_names: {missing_names}")

    def locate(token):
        """Return (sheet, col, row) representative cell for a LOC token."""
        if token in COLS:
            sheet, col = COLS[token]
            row = CROW if sheet == "Commutation" else ROW
            return sheet, col_to_num(col), row
        if token in PARAMS:
            sheet, col, row = PARAMS[token]
            return sheet, col_to_num(col), row
        sheet, c1, r1, _c2, _r2 = loc[token]
        return sheet, c1, r1

    def contains(token, sheet, col, row):
        """True if token's full extent contains the cell (for range names)."""
        if token in RANGES:
            s, c1, r1, c2, r2 = loc[token]
            return s == sheet and c1 <= col <= c2 and r1 <= row <= r2
        ts, tc, tr = locate(token)
        return ts == sheet and tc == col and tr == row

    edges = con.execute(
        "SELECT from_address, to_address FROM agent_binding_dependencies").fetchall()

    def addr_contains(range_a1, sheet, col, row):
        if "!" not in range_a1:
            return False
        s, c1, r1, c2, r2 = parse_range(range_a1)
        return s == sheet and c1 <= col <= c2 and r1 <= row <= r2

    expected = expected_edges()
    present, missing, gaps = [], [], []
    for dep, prec in expected:
        ds, dc, dr = locate(dep)
        ok = False
        for fa, ta in edges:
            if not addr_contains(fa, ds, dc, dr):
                continue
            if prec in RANGES:
                s, c1, r1, c2, r2 = loc[prec]
                ps, pc1, pr1, pc2, pr2 = parse_range(ta)
                if ps == s and not (pc2 < c1 or pc1 > c2 or pr2 < r1 or pr1 > r2):
                    ok = True
                    break
            else:
                ps, pc, pr = locate(prec)
                if addr_contains(ta, ps, pc, pr):
                    ok = True
                    break
        if ok:
            present.append(f"{dep} -> {prec}")
        elif (dep, prec) in KNOWN_EXTRACTOR_GAPS:
            gaps.append(f"{dep} -> {prec} [{KNOWN_EXTRACTOR_GAPS[(dep, prec)]}]")
        else:
            missing.append(f"{dep} -> {prec}")

    verdict = ("PASS" if not missing and not gaps
               else "PASS_WITH_KNOWN_EXTRACTOR_GAPS" if not missing else "FAIL")
    report = {
        "workbook": "workbooks/Solvency2/Solvency2.xlsm",
        "ir_db": "marinade/Solvency2.ir.db",
        "date": date.today().isoformat(),
        "expected_edges": len(expected),
        "edges_found": len(present),
        "edges_missing_known_extractor_gaps": gaps,
        "edges_missing": missing,
        "verdict": verdict,
    }
    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text(json.dumps(report, indent=1))
    print(f"{verdict}: {len(present)}/{len(expected)} lifelib dependencies "
          f"present in workbook IR")
    for m in missing:
        print("  MISSING:", m)
    for g in gaps:
        print("  KNOWN GAP:", g)
    return 1 if missing else 0


if __name__ == "__main__":
    sys.exit(main())
