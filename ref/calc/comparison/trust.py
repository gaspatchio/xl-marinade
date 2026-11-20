"""Trust test: does an engine reproduce a workbook's stored values by really recalculating?

usage: trust.py <engine: ironcalc|formulas> <workbook> <out.json>
The engine is given a copy with every formula cell's cached value deleted.
The JSON holds only counts, addresses, function names and reasons -- no cell values.
"""

import collections
import json
import sys
import time
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from common import ERR, functions, ground_truth, norm, same, strip_cached
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string


def run_ironcalc(nocache, keys):
    import ironcalc

    m = ironcalc.load_from_xlsx(str(nocache), "en", "UTC", "en")
    m.evaluate()
    idx = {p["name"]: i for i, p in enumerate(m.get_worksheets_properties())}
    got = {}
    for sheet, a1 in keys:
        col, row = coordinate_from_string(a1)
        try:
            got[(sheet, a1)] = m.get_cell_value(idx[sheet], row, column_index_from_string(col))
        except Exception as e:  # noqa: BLE001
            got[(sheet, a1)] = f"#EXC:{type(e).__name__}"
    return got


def run_formulas(nocache, keys):
    import formulas

    xl = formulas.ExcelModel().loads(str(nocache)).finish()
    sol = xl.calculate()
    book = nocache.name
    got = {}
    for sheet, a1 in keys:
        k = f"'[{book}]{sheet.upper()}'!{a1}"
        if k not in sol:
            got[(sheet, a1)] = "#MISSING"
            continue
        v = sol[k]
        v = getattr(v, "value", v)
        try:
            v = v[0, 0]
        except Exception:  # noqa: BLE001
            pass
        got[(sheet, a1)] = v
    return got


def main():
    engine, wb, out = sys.argv[1], Path(sys.argv[2]), Path(sys.argv[3])
    res = {"engine": engine, "workbook": wb.name}
    t0 = time.time()
    gt = ground_truth(wb)
    nocache = out.with_suffix(".nocache.xlsx")
    res["formula_cells"] = len(gt)
    res["stripped"] = strip_cached(wb, nocache)
    try:
        t1 = time.time()
        got = {"ironcalc": run_ironcalc, "formulas": run_formulas}[engine](nocache, list(gt))
        res["calc_seconds"] = round(time.time() - t1, 2)
    except BaseException as e:  # noqa: BLE001
        res["load_failure"] = f"{type(e).__name__}: {str(e)[:300]}"
        res["trace_tail"] = traceback.format_exc().splitlines()[-6:]
        out.write_text(json.dumps(res, indent=1))
        print(json.dumps(res)[:600])
        return
    ok = 0
    reasons = collections.Counter()
    by_fn = collections.Counter()
    examples = collections.defaultdict(list)
    bad = set()
    for key, (f, exp, arr) in gt.items():
        g = got[key]
        if same(exp, g):
            ok += 1
            continue
        bad.add(key)
        gs = str(g) if g is not None else ""
        if isinstance(exp, str) and exp in ERR and gs in ERR:
            reason = "different error code"
        elif gs.startswith("#") or any(e in gs for e in ERR):
            reason = f"engine error {gs[:12]}"
        elif g is None or gs == "":
            reason = "engine returned empty"
        elif isinstance(norm(exp), float) and isinstance(norm(g), float):
            reason = "number differs"
        else:
            reason = "value differs"
        if arr:
            reason += " (array formula)"
        reasons[reason] += 1
        for fn in functions(f):
            by_fn[fn] += 1
        if len(examples[reason]) < 5:
            examples[reason].append({"cell": f"{key[0]}!{key[1]}", "functions": functions(f)})
    res.update(
        reproduced=ok,
        share=round(ok / max(1, len(gt)), 6),
        reasons=dict(reasons.most_common()),
        functions_in_failing_cells=dict(by_fn.most_common()),
        examples=examples,
        total_seconds=round(time.time() - t0, 2),
    )
    res["_bad"] = [f"{s}!{a}" for s, a in sorted(bad)]
    out.write_text(json.dumps(res, indent=1))
    print(json.dumps({k: v for k, v in res.items() if k not in ("_bad", "examples")}))


main()
