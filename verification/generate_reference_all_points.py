"""Generate lifelib reference results for ALL 10,000 model points.

Uses BasicTerm_M (the vectorized twin of BasicTerm_S — same formulas over
all model points at once) and cross-checks it against BasicTerm_S on the
10 single-point test IDs before writing. Output feeds the workbook's
Lifelib_Reference sheet for the in-workbook VBA reconciliation.
"""

import json
import sys
from pathlib import Path

import modelx as mx
import pandas as pd

REPO = Path(__file__).resolve().parents[1]
OUT_CSV = REPO / "verification" / "reference" / "basicterm_all_points.csv"
SINGLE_REF = REPO / "verification" / "reference" / "basicterm_s.json"

CROSS_CHECK_RTOL = 1e-9


def main():
    model = mx.read_model(str(REPO / "models" / "basiclife" / "BasicTerm_M"))
    p = model.Projection

    df = pd.DataFrame({
        "point_id": p.model_point_table.index,
        "premium_pp": p.premium_pp(),
        "pv_premiums": p.pv_premiums(),
        "pv_claims": p.pv_claims(),
        "pv_expenses": p.pv_expenses(),
        "pv_commissions": p.pv_commissions(),
        "pv_net_cf": p.pv_net_cf(),
    }).set_index("point_id")

    # Cross-check the vectorized model against the single-point reference
    single = json.loads(SINGLE_REF.read_text())["lifelib_points"]
    worst = 0.0
    for pid, exp in single.items():
        for col in df.columns:
            want = exp["scalars"][col]
            got = float(df.loc[int(pid), col])
            rel = abs(got - want) / max(1.0, abs(want))
            worst = max(worst, rel)
            if rel > CROSS_CHECK_RTOL:
                print(f"MISMATCH point {pid} {col}: M={got} S={want}")
                return 1
    print(f"cross-check M vs S on {len(single)} points: worst rel diff {worst:.2e}")

    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUT_CSV, float_format="%.10f")
    print(f"wrote {OUT_CSV} ({len(df)} points)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
