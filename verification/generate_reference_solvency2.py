"""Generate lifelib reference values for the Solvency2 workbook.

Two outputs:
- solvency2_all_policies.csv — SCR_life, the four live Life(risk) charges,
  the three lapse shocks and the base net asset value for every policy
  (t0=0, scen 1): the in-workbook VBA reconciliation target.
- solvency2.json — deep profiles (policy, scen, t0): every cell the
  workbook maps, incl. PREM-basis commutation columns, the shared policy
  schedule, and the full vector set for each of the 7 stress projections.

Scope note: only the cells lifelib actually evaluates for SCR_life are
dumped (the SCR subgraph). Statutory reserve/profit cells (ChangeRsrv,
ProfitBefTax, InvstIncome, ...) and the VAL rate basis are outside it.
"""

import csv
import json
import time
from pathlib import Path

import modelx as mx

REPO = Path(__file__).resolve().parents[1]
OUT_DIR = REPO / "verification" / "reference"
CSV_OUT = OUT_DIR / "solvency2_all_policies.csv"
JSON_OUT = OUT_DIR / "solvency2.json"

MAX_T = 103   # grid rows t = 0..103 (max last_t across policies is 102)
MAX_X = 130   # commutation ages 0..130

STRESSES = {
    "base": ("base", None, None),
    "mort": ("mort", None, None),
    "longev": ("longev", None, None),
    "exps": ("exps", None, None),
    "lapse_up": ("lapse", "up", None),
    "lapse_down": ("lapse", "down", None),
    "lapse_mass": ("lapse", "mass", None),
}

RISKS = ["mort", "longev", "disab", "lapse", "exps", "rev", "cat"]

# (pid, scen, t0) deep-verification profiles
PROFILES = [(1, 1, 0), (2, 1, 0), (101, 1, 0), (102, 1, 0), (201, 1, 0),
            (1, 1, 5), (201, 3, 0)]


def padvec(fn, n, cast=float):
    """fn(t) for t=0..n, zero-padded to MAX_T."""
    out = [cast(fn(t)) for t in range(n + 1)]
    return out + [0.0] * (MAX_T - n)


def dump_profile(model, pid, scen, t0):
    s = model.SCR_life[t0, pid, scen]
    base = s.Projection["base", None, None]
    pol, asmp = base.pol, base.asmp
    n = int(base.last_t())

    scalars = {
        "product": pol.Product(), "policy_type": int(pol.PolicyType()),
        "gen": int(pol.Gen()), "sex": pol.Sex(),
        "issue_age": int(pol.IssueAge()), "policy_term": int(pol.PolicyTerm()),
        "policy_count": float(pol.PolicyCount()), "sum_assured": float(pol.SumAssured()),
        "prem_freq": int(pol.PremFreq()),
        "int_rate_prem": float(pol.IntRate("PREM")),
        "table_id_prem": int(pol.TableID("PREM")),
        "load_acq_sa": float(pol.LoadAcqSA()),
        "load_maint_prem": float(pol.LoadMaintPrem()),
        "load_maint_sa": float(pol.LoadMaintSA()),
        "load_maint_sa2": float(pol.LoadMaintSA2()),
        "load_maint_prem_waiver": float(pol.LoadMaintPremWaiverPrem()),
        "init_surr_charge": float(pol.InitSurrCharge()),
        "gross_prem_rate": float(pol.GrossPremRate()),
        "ann_prem_rate": float(pol.AnnPremRate()),
        "net_prem_rate_prem": float(pol.NetPremRate("PREM")),
        "cnsmp_tax": float(asmp.CnsmpTax()),
        "infl_rate": float(asmp.InflRate()),
        "last_age": int(asmp.LastAge()),
        "comm_init_prem": float(asmp.CommInitPrem()),
        "comm_ren_prem": float(asmp.CommRenPrem()),
        "comm_ren_term": int(asmp.CommRenTerm()),
        "exps_acq_ann_prem": float(asmp.ExpsAcqAnnPrem()),
        "exps_acq_pol": float(asmp.ExpsAcqPol()),
        "exps_acq_sa": float(asmp.ExpsAcqSA()),
        "exps_maint_ann_prem": float(asmp.ExpsMaintAnnPrem()),
        "exps_maint_pol": float(asmp.ExpsMaintPol()),
        "exps_maint_sa": float(asmp.ExpsMaintSA()),
        "last_t": n,
        "size_premium_0": float(base.SizePremium(0)),
        "size_ann_prem_0": float(base.SizeAnnPrem(0)),
        "size_exps_acq_0": float(base.SizeExpsAcq(0)),
        "size_exps_comm_init_0": float(base.SizeExpsCommInit(0)),
    }

    lt = model.LifeTable[pol.Sex(), pol.IntRate("PREM"), pol.TableID("PREM")]
    commutation = {
        col: [float(getattr(lt, col)(x)) for x in range(MAX_X + 1)]
        for col in ["qx", "lx", "dx", "Dx", "Cx", "Nx", "Mx"]
    }

    shared = {
        "att_age": padvec(base.AttAge, n),
        "base_mort_rate": padvec(lambda t: asmp.BaseMortRate(base.AttAge(t)), n),
        "mort_factor": padvec(asmp.MortFactor, n),
        "surr_rate_base": padvec(asmp.SurrRate, n),
        "reserve_nlp_prem": padvec(lambda t: pol.ReserveNLP_Rate("PREM", t), n),
        "surr_charge": padvec(pol.SurrCharge, n),
        "cash_value_rate": padvec(pol.CashValueRate, n),
        "size_benefit_surr": padvec(base.SizeBenefitSurr, n),
        "size_exps_comm_ren": padvec(base.SizeExpsCommRen, n),
        "disc_rate": padvec(base.DiscRate, n),
    }

    stresses = {}
    for key, sk in STRESSES.items():
        p = s.Projection[sk]
        vecs = {
            "infl_factor": padvec(p.InflFactor, n),
            "pols_if_end": padvec(p.PolsIF_End, n),
            "pols_maturity": padvec(p.PolsMaturity, n),
            "pols_if_beg": padvec(p.PolsIF_Beg, n),
            "pols_if_beg1": padvec(p.PolsIF_Beg1, n),
            "pols_death": padvec(p.PolsDeath, n),
            "pols_surr": padvec(p.PolsSurr, n),
            "prem_income": padvec(p.PremIncome, n),
            "benefit_death": padvec(p.BenefitDeath, n),
            "benefit_surr": padvec(p.BenefitSurr, n),
            "benefit_total": padvec(p.BenefitTotal, n),
            "exps_acq": padvec(p.ExpsAcq, n),
            "exps_comm_init": padvec(p.ExpsCommInit, n),
            "exps_comm_ren": padvec(p.ExpsCommRen, n),
            "exps_maint": padvec(p.ExpsMaint, n),
            "exps_total": padvec(p.ExpsTotal, n),
            "size_exps_maint": padvec(p.SizeExpsMaint, n),
            "pv_prem_income": padvec(p.PV_PremIncome, n),
            "pv_benefit_total": padvec(p.PV_BenefitTotal, n),
            "pv_exps_total": padvec(p.PV_ExpsTotal, n),
            "pv_net_cashflow": padvec(p.PV_NetCashflow, n),
        }
        cellnames = set(p.cells)
        if "MortRateFactor" in cellnames:
            vecs["mort_rate_factor"] = padvec(p.MortRateFactor, n)
        if "SurrRate" in cellnames:
            vecs["surr_rate"] = padvec(p.SurrRate, n)
        if "PolsSurrMass" in cellnames:
            vecs["pols_surr_mass"] = padvec(p.PolsSurrMass, n)
        stresses[key] = vecs

    scr = {
        "net_ast_value": {k: float(s.NetAstValue(*sk)) for k, sk in STRESSES.items()},
        "lapse_risk": {sh: float(s.LapseRisk(sh)) for sh in ["up", "down", "mass"]},
        "life": {r: float(s.Life(r)) for r in RISKS},
        "scr_life": float(s.SCR_life()),
    }

    return {"scalars": scalars, "commutation": commutation, "shared": shared,
            "stresses": stresses, "scr": scr}


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    model = mx.read_model(str(REPO / "models" / "solvency2" / "model"))
    inp = model.Input

    corr = {f"{r}|{c}": float(v) for (r, c), v in
            ((k, inp.CorrData[k]) for k in inp.CorrData.keys())}
    factors = {"|".join(str(x) for x in k): float(inp.FactorData[k])
               for k in inp.FactorData.keys()}

    profiles = {}
    for pid, scen, t0 in PROFILES:
        t = time.time()
        profiles[f"{pid}|{scen}|{t0}"] = dump_profile(model, pid, scen, t0)
        print(f"profile {pid}|{scen}|{t0} in {time.time() - t:.1f}s")

    JSON_OUT.write_text(json.dumps({
        "params": {"MAX_T": MAX_T, "MAX_X": MAX_X},
        "corr": corr, "factors": factors, "profiles": profiles,
    }, indent=None))
    print(f"wrote {JSON_OUT} ({JSON_OUT.stat().st_size // 1024} KB)")

    t = time.time()
    with CSV_OUT.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["policy_id", "scr_life", "life_mort", "life_longev",
                    "life_lapse", "life_exps", "lapse_up", "lapse_down",
                    "lapse_mass", "nav_base"])
        for pid in range(1, 301):
            s = model.SCR_life[0, pid, 1]
            w.writerow([pid, s.SCR_life(), s.Life("mort"), s.Life("longev"),
                        s.Life("lapse"), s.Life("exps"), s.LapseRisk("up"),
                        s.LapseRisk("down"), s.LapseRisk("mass"),
                        s.NetAstValue()])
            if pid % 50 == 0:
                print(f"  {pid}/300 ({time.time() - t:.0f}s)")
    print(f"wrote {CSV_OUT}")


if __name__ == "__main__":
    main()
