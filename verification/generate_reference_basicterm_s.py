"""Generate lifelib reference values for BasicTerm_S.

Runs the BasicTerm_S model for a set of test model points and dumps
every cell (time-indexed vectors and scalars) to JSON. The JSON is the
"expected" side of the value proof for the Excel workbook.

Usage:
    .venv/bin/python verification/generate_reference_basicterm_s.py
"""

import json
import sys
from pathlib import Path

import modelx as mx

REPO = Path(__file__).resolve().parents[1]
MODEL_PATH = REPO / "models" / "basiclife" / "BasicTerm_S"
OUT_PATH = REPO / "verification" / "reference" / "basicterm_s.json"

# Test model points chosen to span age, sex, and all three policy terms
# (10/15/20 years). Point 1 is lifelib's documented default.
TEST_POINTS = [1, 2, 3, 4, 5, 172, 3001, 5555, 9999, 10000]

# Cells projected over time t = 0 .. proj_len-1
TIME_CELLS = [
    "duration",
    "age",
    "mort_rate",
    "mort_rate_mth",
    "lapse_rate",
    "inflation_factor",
    "pols_if",
    "pols_death",
    "pols_lapse",
    "pols_maturity",
    "claims",
    "premiums",
    "commissions",
    "expenses",
    "net_cf",
]

# Scalar cells (per model point)
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
    result = {"model": "BasicTerm_S", "lifelib_points": {}}

    for pid in TEST_POINTS:
        proj = model.Projection[pid]
        n = proj.proj_len()
        point = {
            "scalars": {},
            "time_vectors": {},
            "disc_factors": list(map(float, proj.disc_factors())),
            # 0-ary vector cell in lifelib (numpy array over t), like disc_factors
            "disc_rate_mth": list(map(float, proj.disc_rate_mth())),
        }
        for name in SCALAR_CELLS:
            v = getattr(proj, name)()
            point["scalars"][name] = (
                bool(v) if isinstance(v, bool) else float(v)
            )
        for name in TIME_CELLS:
            cell = getattr(proj, name)
            point["time_vectors"][name] = [float(cell(t)) for t in range(n)]
        result["lifelib_points"][str(pid)] = point
        print(f"point {pid}: proj_len={n} premium_pp={point['scalars']['premium_pp']}"
              f" pv_net_cf={point['scalars']['pv_net_cf']:.6f}"
              f" check={point['scalars']['check_pv_net_cf']}")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(result, indent=1))
    print(f"wrote {OUT_PATH}")


if __name__ == "__main__":
    sys.exit(main())
