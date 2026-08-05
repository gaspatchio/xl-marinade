"""Generate lifelib reference values for the SmithWilson workbook.

Dumps every cell of the smithwilson model (11 cells, no model points):
the observed inputs, the Wilson function grid W(i,j) for the full
extrapolation horizon, the fitted zeta vector, and the extrapolated
bond prices P and spot rates R out to T_MAX years.
"""

import json
from pathlib import Path

import modelx as mx

REPO = Path(__file__).resolve().parents[1]
OUT = REPO / "verification" / "reference" / "smithwilson.json"

T_MAX = 65  # extrapolation horizon in years (N=25 observed + 40 extrapolated)


def main():
    model = mx.read_model(str(REPO / "models" / "smithwilson" / "model"))
    sw = model.SmithWilson
    n = int(sw.N)

    ref = {
        "params": {
            "N": n,
            "UFR": float(sw.UFR),
            "alpha": float(sw.alpha),
            "T_MAX": T_MAX,
        },
        "spot_rates": [float(x) for x in sw.spot_rates],
        "u": [float(sw.u(i)) for i in range(1, T_MAX + 1)],
        "m": [float(sw.m(i)) for i in range(1, n + 1)],
        "mu": [float(sw.mu(i)) for i in range(1, T_MAX + 1)],
        # W grid rows i=1..T_MAX, cols j=1..N (first N rows = W_matrix)
        "W": [[float(sw.W(i, j)) for j in range(1, n + 1)]
              for i in range(1, T_MAX + 1)],
        "zeta": [float(sw.zeta(j)) for j in range(1, n + 1)],
        "P": [float(sw.P(i)) for i in range(1, T_MAX + 1)],
        "R": [float(sw.R(i)) for i in range(1, T_MAX + 1)],
    }

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(ref, indent=1))
    print(f"wrote {OUT}")
    print(f"N={n} T_MAX={T_MAX} R({n})={ref['R'][n-1]:.6f} "
          f"R({T_MAX})={ref['R'][T_MAX-1]:.6f}")


if __name__ == "__main__":
    main()
