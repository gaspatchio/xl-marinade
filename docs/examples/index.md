# Examples

Three end-to-end how-tos, each self-contained and copy-pasteable. Every workbook
below is built inline with `openpyxl` — generic cell content, no client data — so
you can run each example verbatim and get the exact output shown.

## Extract a workbook and query its dependencies

First, build a tiny synthetic workbook: a 5-period amortization with two formula
columns that depend on each other — `Balance` this period depends on last period's
`Balance` and `Decay`, and `Decay` this period depends on this period's `Balance`.

Save as `make_book.py`:

```python
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sheet1"
ws["A1"] = "Month"
ws["B1"] = "Balance"
ws["C1"] = "Decay"
ws["A2"] = 0
ws["B2"] = 1000
ws["C2"] = 0
for row in range(3, 8):
    prev = row - 1
    ws[f"A{row}"] = row - 2
    ws[f"B{row}"] = f"=B{prev}-C{prev}"
    ws[f"C{row}"] = f"=B{row}*0.02"
wb.save("book.xlsx")
```

```bash
python make_book.py
marinade extract book.xlsx -o ir.db
```

Now query the binding-level dependency graph — [`agent_binding_dependencies`](../reference/schema.md#dependencies)
— with the `sqlite3` CLI:

```bash
sqlite3 -header -column ir.db \
  "SELECT from_label, to_label, kind, edge_count FROM agent_binding_dependencies
   ORDER BY from_label, to_label;"
```

```text
from_label  to_label      kind     edge_count
----------  ------------  -------  ----------
Balance     Decay         formula  1
Balance     Sheet1!A1:C2  formula  1
Decay       Balance       formula  1
```

Three edges: `Balance` reads `Decay` and the seed row (`A1:C2`, the header +
period-0 constants, which didn't get a label of its own), and `Decay` reads
`Balance` right back — the cyclic binding graph described in
[Bindings & the graph](../guide/bindings.md). (The `ORDER BY from_label, to_label`
gives a stable row order — edge rows are otherwise unordered.) Same query from Python:

```python
import sqlite3

con = sqlite3.connect("ir.db")
con.row_factory = sqlite3.Row
rows = con.execute(
    """
    SELECT from_label, to_label, kind, edge_count
    FROM agent_binding_dependencies
    ORDER BY from_label, to_label
    """
).fetchall()
for row in rows:
    print(dict(row))
```

```text
{'from_label': 'Balance', 'to_label': 'Decay', 'kind': 'formula', 'edge_count': 1}
{'from_label': 'Balance', 'to_label': 'Sheet1!A1:C2', 'kind': 'formula', 'edge_count': 1}
{'from_label': 'Decay', 'to_label': 'Balance', 'kind': 'formula', 'edge_count': 1}
```

## Diff two versions of a model

Build two versions of the same tiny model — v2 changes the `Decay` formula's
multiplier from `0.02` to `0.03`, nothing else:

Save as `make_versions.py`:

```python
import openpyxl


def build(decay_rate: float) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Month"
    ws["B1"] = "Balance"
    ws["C1"] = "Decay"
    ws["A2"] = 0
    ws["B2"] = 1000
    ws["C2"] = 0
    for row in range(3, 8):
        prev = row - 1
        ws[f"A{row}"] = row - 2
        ws[f"B{row}"] = f"=B{prev}-C{prev}"
        ws[f"C{row}"] = f"=B{row}*{decay_rate}"
    return wb


build(0.02).save("model_v1.xlsx")
build(0.03).save("model_v2.xlsx")
```

```bash
python make_versions.py
marinade extract model_v1.xlsx -o a.db
marinade extract model_v2.xlsx -o b.db
marinade diff a.db b.db -o changelist.json
```

The changelist (see [Diffing workbooks](../guide/diffing.md) for the full field
reference) reports the change once, at the binding level, not once per cell:

```json
{
  "summary": {
    "total_changes": 8,
    "bindings_formula_changed": 1,
    "cells_formula_changed": 5,
    "ir_inference_changes": 1
  },
  "changes": [
    {
      "seq": 3,
      "type": "BINDING_FORMULA_CHANGED",
      "sheet": "Sheet1",
      "address": "C3:C7",
      "old_formula": "RC[-1]*0.02",
      "new_formula": "RC[-1]*0.03",
      "modification_kind": "logic_change"
    }
  ]
}
```

`bindings_formula_changed: 1` is the number that matters for review — the five
`cells_formula_changed` entries are the same edit, once per cell in the `Decay`
column, rolled up into that one `BINDING_FORMULA_CHANGED` entry.
`modification_kind: "logic_change"` (rather than `reference_shift`) confirms this
is a real formula edit, not just references moving because of an insertion
elsewhere in the sheet. `ir_inference_changes: 1` is unrelated noise — a
table-candidate re-ranking that shifted because the underlying bindings changed,
not because you edited a cell — and is worth filtering out before you review a
changelist:

```python
import json
from pathlib import Path

changelist = json.loads(Path("changelist.json").read_text())

workbook_edits = [
    c for c in changelist["changes"] if c.get("layer") != "ir_inference"
]
for change in workbook_edits:
    if change["type"] == "BINDING_FORMULA_CHANGED":
        print(
            f"{change['address']}: {change['old_formula']} -> "
            f"{change['new_formula']} ({change['modification_kind']})"
        )
```

```text
C3:C7: RC[-1]*0.02 -> RC[-1]*0.03 (logic_change)
```

## Generate documentation (deterministic, then enriched)

Using the `ir.db` from the first example:

```bash
marinade document ir.db -o out/
```

Writes `out/documentation.md` (human-readable) and `out/model_spec.json`
(machine-readable) — no network call is made. Both describe the same two
calculated variables from the extraction above (`Balance`, `Decay`), their
inferred role (`Calculation`), and which of them need reconciliation against a
reference implementation.

To layer LLM-written narrative documentation on top, install the extra, set a
key, and pass `--enrich` (see [LLM add-on](../llm.md) for provider setup):

```bash
pip install 'xl-marinade[llm]'
export OPENAI_API_KEY="sk-..."
marinade document ir.db -o out/ --enrich
```

`--enrich` degrades gracefully rather than failing: without the `[llm]` extra
installed, `marinade document ... --enrich` prints
`note: xl-marinade[llm] not installed — deterministic documentation only` and
writes the exact same deterministic `documentation.md`/`model_spec.json` as the
plain command above. The same fallback applies if the extra is installed but no
key is configured, or the request errors — you always get the deterministic
output, and enrichment can only add to it, never replace it with something
worse or block the run.

## A complete worked example set: lifelib actuarial workbook twins

Beyond these inline how-tos, the
[`examples/lifelib-actuarial-workbooks`](https://github.com/gaspatchio/xl-marinade/tree/examples/lifelib-actuarial-workbooks)
branch carries five full Excel twins of open-source [lifelib](https://lifelib.io)
actuarial models (term assurance, in-force portfolio, universal life,
Smith-Wilson curve fitting, Solvency II SCR). Each is authored with the
conventions Marinade rewards — machine-key rows, uniform per-column formulas,
named ranges — and ships with committed value proofs against the Python model,
`marinade`-based structure proofs, and in-workbook VBA batch reconciliations.
Start with [that branch's README](https://github.com/gaspatchio/xl-marinade/blob/examples/lifelib-actuarial-workbooks/README.md).

## Next

- [Quickstart](../quickstart.md) — the same extract → document → diff workflow,
  narrated end to end.
- [CLI reference](../reference/cli.md) — every flag used above.
- [SQLite schema](../reference/schema.md) — the full set of `agent_*`/`marinade_*`
  views to query.
