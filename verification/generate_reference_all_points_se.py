"""Generate lifelib reference results for ALL 10,000 BasicTerm_SE model points.

Uses BasicTerm_ME (the vectorized twin of BasicTerm_SE — same in-force
mechanics over all model points at once) and cross-checks it against
BasicTerm_SE on the single-point test IDs before writing.
"""

import json
import math
import sys
from pathlib import Path

import modelx as mx
import pandas as pd

REPO = Path(__file__).resolve().parents[1]
OUT_CSV = REPO / "verification" / "reference" / "basicterm_se_all_points.csv"
SINGLE_REF = REPO / "verification" / "reference" / "basicterm_se.json"

CROSS_CHECK_RTOL = 1e-9


def main():
    model = mx.read_model(str(REPO / "models" / "basiclife" / "BasicTerm_ME"))
    p = model.Projection

    df = pd.DataFrame({
        "policy_id": p.model_point_table.index,
        "premium_pp": p.premium_pp(),
        "pv_premiums": p.pv_premiums(),
        "pv_claims": p.pv_claims(),
        "pv_expenses": p.pv_expenses(),
        "pv_commissions": p.pv_commissions(),
        "pv_net_cf": p.pv_net_cf(),
    }).set_index("policy_id")

    single = json.loads(SINGLE_REF.read_text())["lifelib_points"]
    worst = 0.0
    for pid, exp in single.items():
        for col in df.columns:
            want = exp["scalars"][col]
            got = float(df.loc[int(pid), col])
            if math.isnan(want) and math.isnan(got):
                continue
            rel = abs(got - want) / max(1.0, abs(want))
            worst = max(worst, rel)
            if rel > CROSS_CHECK_RTOL:
                print(f"MISMATCH point {pid} {col}: ME={got} SE={want}")
                return 1
    print(f"cross-check ME vs SE on {len(single)} points: worst rel diff {worst:.2e}")

    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUT_CSV, float_format="%.10f")
    print(f"wrote {OUT_CSV} ({len(df)} points)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
