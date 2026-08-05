"""Structure proof: SmithWilson workbook dependency graph vs lifelib call graph.

Same method as the other structure proofs; expected edges transcribed from
models/smithwilson/model/SmithWilson/__init__.py. Mapping notes:
  - u(i) appears on three sheets (Calibration!B, Wilson!A, Extrapolation!A);
    each use is asserted against the copy the consuming formula reads.
  - The vector cells are the same workbook cells as their scalar
    counterparts (m_vector == the m column, mu_vector == the mu column,
    W_matrix == the top 25 rows of the W grid, zeta == zeta_vector), so
    the vector/scalar edges are identities, not graph edges.
  - zeta_vector <- m_vector / mu_vector run through the explicit
    (m - mu) column on Calibration (lifelib computes the difference inline);
    the proof asserts both hops.
  - zeta is a single CSE array formula (MINVERSE/MMULT); its binding is
    the top-left cell Calibration!G4.
"""

import json
import re
import sqlite3
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
DB = REPO / "marinade" / "SmithWilson.ir.db"
REPORT = REPO / "verification" / "reports" / "smithwilson_structure_proof.json"

KNOWN_EXTRACTOR_GAPS: dict = {}

LOC = {
    # inputs
    "spot_rates": "Spot_Rates!C10",
    "UFR": "Assumptions!C4",
    "alpha": "Assumptions!C6",
    # Calibration (rows 4..28; representative row 10)
    "u@cal": "Calibration!B10",
    "spot_rates@cal": "Calibration!C10",
    "m": "Calibration!D10",
    "mu@cal": "Calibration!E10",
    "d": "Calibration!F10",            # m - mu, inline in lifelib's zeta_vector
    "zeta": "Calibration!G4",          # CSE array formula, top-left cell
    # Wilson grid (rows 4..68 x cols B..Z; representative interior cell)
    "u@wilson": "Wilson!A30",
    "W": "Wilson!M30",
    # Extrapolation (rows 4..68; representative row 30 = t 27, beyond N)
    "u@ext": "Extrapolation!A30",
    "mu@ext": "Extrapolation!B30",
    "P": "Extrapolation!C30",
    "R": "Extrapolation!D30",
}

EXPECTED_EDGES = [
    # m(i) = (1 + spot_rates[i-1]) ** -u(i)
    ("m", "u@cal"), ("m", "spot_rates@cal"),
    ("spot_rates@cal", "spot_rates"),
    # mu(i) = exp(-UFR * u(i))  — on both consuming sheets
    ("mu@cal", "UFR"), ("mu@cal", "u@cal"),
    ("mu@ext", "UFR"), ("mu@ext", "u@ext"),
    # W(i, j) = f(UFR, alpha, u)
    ("W", "UFR"), ("W", "alpha"), ("W", "u@wilson"),
    # zeta_vector = inv(W_matrix) @ (m_vector - mu_vector)
    ("zeta", "W"), ("zeta", "d"),
    ("d", "m"), ("d", "mu@cal"),
    # P(i) = mu(i) + sum(zeta(j) * W(i, j))
    ("P", "mu@ext"), ("P", "zeta"), ("P", "W"), ("P", "u@ext"),
    # R(i) = (1 / P(i)) ** (1 / u(i)) - 1
    ("R", "P"), ("R", "u@ext"),
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
        "workbook": "workbooks/SmithWilson/SmithWilson.xlsm",
        "ir_db": "marinade/SmithWilson.ir.db",
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
