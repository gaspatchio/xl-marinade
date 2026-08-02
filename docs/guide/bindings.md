# Bindings & the graph

A **binding** is a labelled group of cells that represents one variable in the
model — a whole column ("Policies In Force", one value per projection month), a
single input cell ("Discount Rate"), or a block of assumptions. Extraction groups
the raw cell-level graph into bindings so that auditing a model means reasoning
about its variables, not walking thousands of individual cell references one at a
time.

A binding is always a contiguous rectangular range on a single sheet, and it is
one of two kinds:

- **formula** — the cells share a repeated formula pattern (e.g. the same formula,
  position-adjusted, down every row of a column).
- **constant** — the cells hold input values with no formula.

## A tiny worked example

Take a two-column projection sheet, one row per month:

| Cell | Formula | Meaning |
|---|---|---|
| `B2` | `=B1-C1` | Policies in force this month = last month's in-force − last month's lapses |
| `B3` | `=B2-C2` | (same pattern, next row) |
| ⋮ | ⋮ | ⋮ |
| `C2` | `=B2*0.02` | Lapses this month = 2% of this month's in-force |
| `C3` | `=B3*0.02` | (same pattern, next row) |
| ⋮ | ⋮ | ⋮ |

Column B collapses into one binding — call it `pols_if` — and column C collapses
into another, `pols_lapse`. Instead of a graph with one node per cell, you now
have a graph with two nodes: a variable, not a spreadsheet coordinate.

## How bindings relate: binding edges

Bindings are connected by edges derived from the underlying cell references: if
any cell in one binding refers to any cell in another, an edge connects the two
bindings (with a count of how many underlying cell references it represents).
Extraction also recognizes VBA-driven dependencies the same way — a macro that
paste-as-values a template row into an output block produces a binding edge too,
tagged with the procedure responsible; see [VBA extraction](vba.md).

**Binding graphs can be cyclic.** In the example above, `pols_if` this month
depends on `pols_lapse` last month, and `pols_lapse` this month depends on
`pols_if` this month — cell by cell, that's a strict DAG (nothing depends on its
own future), but collapsed to bindings, `pols_if` and `pols_lapse` depend on each
other. This is normal for any recursive projection model, not a defect: extraction
handles cyclic binding graphs correctly and always produces a stable ordering over
them, so tooling built on top (documentation, lineage walks) doesn't need to
special-case recursive models.

## Why bindings are the useful audit unit

A model with 10,000 policy rows and 50 calculated columns has 500,000 formula
cells but on the order of 50 model variables. Auditing at the cell level means
drowning in near-duplicate edges; auditing at the binding level means answering
the question an actuary or analyst actually asks — "what does Net Cashflow depend
on, and what depends on it?" — with a handful of named variables instead of a wall
of cell coordinates.

## Next

- [How extraction works](how-extraction-works.md) — where bindings fit in the
  overall pipeline.
- [Diffing workbooks](diffing.md) — how bindings are matched across two versions
  of a workbook.
- [VBA extraction](vba.md) — how macro-driven edges connect to this same graph.
