"""Run every model point of a lifelib workbook through IronCalc and compare with
the workbook's own Excel batch results and the lifelib reference. Values stay local."""

import json
import sys
import time

import ironcalc
import openpyxl

wb_path, nocache, out, limit = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4])
wv = openpyxl.load_workbook(wb_path, data_only=True, read_only=True)
ref = {
    r[0]: r[1:]
    for r in wv["Lifelib_Reference"].iter_rows(min_row=3, values_only=True)
    if isinstance(r[0], int)
}
xl = {}
for r in wv["Batch_Results"].iter_rows(min_row=13, values_only=True):
    if isinstance(r[1], int):
        xl[r[1]] = (r[2], r[9])  # premium_pp (wb), pv_net_cf (wb): real Excel
m = ironcalc.load_from_xlsx(nocache, "en", "UTC", "en")
idx = {p["name"]: i for i, p in enumerate(m.get_worksheets_properties())}
c, s = idx["Control"], idx["Summary"]
ROWS = [int(x) for x in sys.argv[5].split(",")]
OUT = list(
    zip(
        ROWS,
        ["premium_pp", "pv_premiums", "pv_claims", "pv_expenses", "pv_commissions", "pv_net_cf"],
        strict=False,
    )
)
ids = sorted(ref)[:limit]
t0 = time.time()
res = {}
for pid in ids:
    m.set_user_input(c, 4, 3, str(pid))
    m.evaluate()
    res[pid] = [m.get_cell_value(s, row, 3) for row, _ in OUT]
el = time.time() - t0


def rel(a, b):
    return abs(a - b) / max(abs(a), abs(b), 1e-300)


prem_bad = [p for p in ids if res[p][0] != xl[p][0]]
ncf_max = max(rel(res[p][5], xl[p][1]) for p in ids)
ncf_exact = sum(res[p][5] == xl[p][1] for p in ids)
lif_max = {name: max(rel(res[p][i], ref[p][i]) for p in ids) for i, (_, name) in enumerate(OUT)}
distinct = len({tuple(v) for v in res.values()})
maxabs = max(abs(res[p][5] - xl[p][1]) for p in ids)
summary = dict(
    pv_net_cf_max_abs_vs_excel=maxabs,
    engine="ironcalc",
    policies=len(ids),
    seconds=round(el, 1),
    per_policy_ms=round(1000 * el / len(ids), 1),
    distinct_result_rows=distinct,
    premium_exact_vs_excel=len(ids) - len(prem_bad),
    premium_mismatch_ids=prem_bad[:20],
    pv_net_cf_bit_identical_vs_excel=ncf_exact,
    pv_net_cf_max_rel_vs_excel=ncf_max,
    max_rel_vs_lifelib=lif_max,
)
json.dump(
    dict(summary=summary, results={str(k): v for k, v in res.items()}),
    open(out, "w", encoding="utf-8"),
)
print(json.dumps(summary))
