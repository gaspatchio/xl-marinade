"""Generate lifelib reference values for BasicTerm_SE.

Same pattern as generate_reference_basicterm_s.py, extended for SE's
in-force mechanics: duration_mth, is_active, pols_if_at timings, new
business, and the rate-table premium. Cells that lifelib itself cannot
evaluate at inactive durations (mortality lookups below the table's age
floor for not-yet-issued policies) are recorded as null and skipped by
the value proof.
"""

import json
import sys
from pathlib import Path

import modelx as mx

REPO = Path(__file__).resolve().parents[1]
MODEL_PATH = REPO / "models" / "basiclife" / "BasicTerm_SE"
OUT_PATH = REPO / "verification" / "reference" / "basicterm_se.json"

# Diverse in-force profile: negative durations (future new business),
# issue-month, near-maturity, zero and maximal policy counts, all terms.
TEST_POINTS = [1, 2, 8, 11, 12, 50, 65, 200, 404, 502]

TIME_CELLS = [
    "duration_mth",
    "duration",
    "age",
    "is_active",
    "mort_rate",        # guarded: may be unevaluable at negative durations
    "mort_rate_mth",    # guarded
    "lapse_rate",
    "inflation_factor",
    "pols_if",
    "pols_death",
    "pols_lapse",
    "pols_maturity",
    "pols_new_biz",
    "claims",
    "premiums",
    "commissions",
    "expenses",
    "net_cf",
]
GUARDED = {"mort_rate", "mort_rate_mth"}

TIMINGS = ["BEF_MAT", "BEF_NB", "BEF_DECR"]

SCALAR_CELLS = [
    "age_at_entry",
    "policy_term",
    "sum_assured",
    "proj_len",
    "pv_claims",
    "pv_premiums",
    "pv_expenses",
    "pv_commissions",
    "pv_net_cf",
    "pv_pols_if",
    "net_premium_pp",
    "premium_pp",
    "check_pv_net_cf",
]


def main():
    model = mx.read_model(str(MODEL_PATH))
    result = {"model": "BasicTerm_SE", "lifelib_points": {}}

    for pid in TEST_POINTS:
        proj = model.Projection[pid]
        n = proj.proj_len()
        point = {
            "scalars": {},
            "time_vectors": {},
            "pols_if_at": {},
            "disc_factors": list(map(float, proj.disc_factors())) if n else [],
        }
        for name in SCALAR_CELLS:
            v = getattr(proj, name)()
            point["scalars"][name] = bool(v) if isinstance(v, bool) else float(v)
        point["scalars"]["policy_count"] = float(proj.model_point()["policy_count"])
        point["scalars"]["duration_mth_0"] = float(proj.model_point()["duration_mth"])
        point["scalars"]["pols_if_init"] = float(proj.pols_if_init())

        for name in TIME_CELLS:
            cell = getattr(proj, name)
            vec = []
            for t in range(n):
                try:
                    v = cell(t)
                    vec.append(float(v))
                except (KeyError, ValueError):
                    if name not in GUARDED:
                        raise
                    vec.append(None)
            point["time_vectors"][name] = vec
        for timing in TIMINGS:
            point["pols_if_at"][timing] = [
                float(proj.pols_if_at(t, timing)) for t in range(n)]

        result["lifelib_points"][str(pid)] = point
        print(f"point {pid}: proj_len={n} count={point['scalars']['policy_count']:.0f} "
              f"dm0={point['scalars']['duration_mth_0']:.0f} "
              f"premium_pp={point['scalars']['premium_pp']} "
              f"pv_net_cf={point['scalars']['pv_net_cf']:.4f} "
              f"check={point['scalars']['check_pv_net_cf']}")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(result, indent=1))
    print(f"wrote {OUT_PATH}")


if __name__ == "__main__":
    sys.exit(main())
