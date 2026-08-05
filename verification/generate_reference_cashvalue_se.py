"""Generate lifelib reference values for CashValue_SE (savings / universal life).

Two outputs:
  - cashvalue_se.json: full per-t vectors for a set of (point, scenario)
    profiles — the deep value-proof reference.
  - cashvalue_se_all_runs.csv: PV results for every (point, scenario)
    combination (4 x 10) — the in-workbook batch reconciliation target.
"""

import json
import sys
from pathlib import Path

import modelx as mx

REPO = Path(__file__).resolve().parents[1]
MODEL_PATH = REPO / "models" / "savings" / "CashValue_SE"
OUT_JSON = REPO / "verification" / "reference" / "cashvalue_se.json"
OUT_CSV = REPO / "verification" / "reference" / "cashvalue_se_all_runs.csv"

# (point_id, scen_id): all four products on scenario 1, plus scenario spread
DEEP_PROFILES = [(1, 1), (2, 1), (3, 1), (4, 1), (1, 5), (4, 10)]

TIME_CELLS = [
    "duration_mth", "duration", "age",
    "premium_pp", "mort_rate", "mort_rate_mth", "lapse_rate",
    "surr_charge_rate", "inflation_factor",
    "pols_if", "pols_maturity", "pols_new_biz", "pols_death", "pols_lapse",
    "inv_return_mth", "inv_income_pp", "inv_income",
    "net_amt_at_risk", "coi_rate", "coi_pp", "coi",
    "prem_to_av_pp", "prem_to_av", "maint_fee_pp", "maint_fee",
    "premiums", "commissions", "expenses",
    "claims_over_av", "surr_charge", "av_change", "net_cf",
    "margin_expense", "margin_mortality",
]
AV_PP_TIMINGS = ["BEF_PREM", "BEF_FEE", "BEF_INV", "MID_MTH"]
AV_TIMINGS = ["BEF_MAT", "BEF_NB", "BEF_FEE"]
POLS_TIMINGS = ["BEF_MAT", "BEF_NB", "BEF_DECR"]
CLAIM_KINDS = ["DEATH", "LAPSE", "MATURITY"]

SCALAR_CELLS = [
    "age_at_entry", "policy_term", "sum_assured", "proj_len",
    "pv_premiums", "pv_claims", "pv_expenses", "pv_commissions",
    "pv_inv_income", "pv_av_change", "pv_net_cf", "pv_pols_if",
    "check_av_roll_fwd", "check_margin", "check_pv_net_cf",
]

PV_COLS = ["pv_premiums", "pv_claims", "pv_expenses", "pv_commissions",
           "pv_inv_income", "pv_av_change", "pv_net_cf"]


def main():
    model = mx.read_model(str(MODEL_PATH))
    result = {"model": "CashValue_SE", "profiles": {}}
    csv_rows = ["point_id,scen_id," + ",".join(PV_COLS)]

    n_points = len(model.Projection.model_point_table)
    n_scens = int(model.Projection.std_norm_rand.index.get_level_values(0).max())
    print(f"{n_points} model points x {n_scens} scenarios")

    for scen in range(1, n_scens + 1):
        # scen_id is a Reference on the base space; assigning it invalidates
        # all cached ItemSpaces so each scenario recomputes from scratch
        model.Projection.scen_id = scen
        for pid in range(1, n_points + 1):
            proj = model.Projection[pid]
            pvs = [getattr(proj, c)() for c in PV_COLS]
            csv_rows.append(f"{pid},{scen}," + ",".join(f"{v:.10f}" for v in pvs))

            if (pid, scen) in DEEP_PROFILES:
                n = proj.proj_len()
                point = {"scalars": {}, "time_vectors": {}, "av_pp_at": {},
                         "av_at": {}, "pols_if_at": {}, "claims": {},
                         "disc_factors": list(map(float, proj.disc_factors()))}
                for name in SCALAR_CELLS:
                    v = getattr(proj, name)()
                    point["scalars"][name] = bool(v) if isinstance(v, bool) else float(v)
                mpt = proj.model_point()
                for extra in ["policy_count", "duration_mth", "premium_pp",
                              "av_pp_init", "load_prem_rate"]:
                    point["scalars"][f"mp_{extra}"] = float(mpt[extra])
                point["scalars"]["mp_has_surr_charge"] = bool(mpt["has_surr_charge"])
                point["scalars"]["mp_is_wl"] = bool(mpt["is_wl"])
                point["scalars"]["pols_if_init"] = float(proj.pols_if_init())

                for name in TIME_CELLS:
                    cell = getattr(proj, name)
                    point["time_vectors"][name] = [float(cell(t)) for t in range(n)]
                for tm in AV_PP_TIMINGS:
                    point["av_pp_at"][tm] = [float(proj.av_pp_at(t, tm)) for t in range(n)]
                for tm in AV_TIMINGS:
                    point["av_at"][tm] = [float(proj.av_at(t, tm)) for t in range(n)]
                for tm in POLS_TIMINGS:
                    point["pols_if_at"][tm] = [float(proj.pols_if_at(t, tm)) for t in range(n)]
                for kind in CLAIM_KINDS:
                    point["claims"][kind] = [float(proj.claims(t, kind)) for t in range(n)]
                result["profiles"][f"{pid}|{scen}"] = point
                print(f"deep profile point {pid} scen {scen}: proj_len={n} "
                      f"pv_net_cf={point['scalars']['pv_net_cf']:.2f} "
                      f"checks={point['scalars']['check_av_roll_fwd']}/"
                      f"{point['scalars']['check_margin']}/"
                      f"{point['scalars']['check_pv_net_cf']}")
            # reset model to free ItemSpace memory between scenarios
            model.Projection.clear_at(pid)

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(result, indent=1))
    OUT_CSV.write_text("\n".join(csv_rows) + "\n")
    print(f"wrote {OUT_JSON}")
    print(f"wrote {OUT_CSV} ({len(csv_rows) - 1} runs)")


if __name__ == "__main__":
    sys.exit(main())
