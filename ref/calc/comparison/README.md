# Comparison-run scripts

Prototype scripts from the 2026-11-17 engine comparison (see `../README.md`). They contain
no workbook content. Raw results on the lifelib workbooks are in `results/`.

Setup (scratch environment, not the repo's):

    pip install ironcalc openpyxl            # IronCalc only
    pip install formulas numpy-financial scipy   # only for the formulas cross-check and the probe

| Script | What it does |
|---|---|
| `common.py` | Shared code: reads each formula cell's stored value, and writes a copy of a workbook with every formula cell's stored value deleted, so an engine has to compute. |
| `trust.py` | `python trust.py <ironcalc|formulas> <workbook> <out.json>`. Recalculates the cache-free copy and compares with the stored values (1e-9). Output: share reproduced, reasons, function names in failing cells, example addresses. No cell values. |
| `policies_ironcalc.py` | `python policies_ironcalc.py <workbook> <cache-free copy> <out.json> <limit> <rows>`. For the lifelib BasicTerm workbooks: runs every model point and compares with the workbook's Excel batch results. `<rows>` is the Summary rows for premium, pv_premiums, pv_claims, pv_expenses, pv_commissions, pv_net_cf: `13,17,11,19,20,21` for BasicTerm_S, `11,17,18,19,20,21` for BasicTerm_SE. |
| `probe.py` | `python probe.py <out dir>`. Builds a synthetic workbook of 96 checked cases (functions and rounding ties) and scores both engines. |

IronCalc prints "Unexpected type (empty) in ..." for each stripped cell when it loads a
cache-free copy. The lines are harmless; filter with `grep -v "^Unexpected type"`.

`trust.py` writes the cache-free copy next to the output file as `<out>.nocache.xlsx`.
That copy still holds the workbook's inputs and formulas, so treat it like the workbook.
