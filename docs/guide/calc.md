# Recalculating with `marinade calc` (experimental)

!!! warning "Experimental"
    `marinade calc` and `marinade values` are experimental. They sit outside the SemVer
    contract that covers the SQLite output: file formats, status names and refusal rules can
    change in any release. Do not build a pipeline that cannot absorb that.

An extracted workbook stores results only for the inputs sitting in it, often a single
policy. `marinade calc` answers a different question: given other input values, what do
the output cells become? It recalculates the workbook with [IronCalc](https://github.com/ironcalc/ironcalc),
proves itself against the workbook's own stored values first, and puts a **status** on
every value it returns. `marinade values` returns the stored values in the same row shape,
without recalculating.

The rule behind the design: a missing value costs a missing comparison, a wrong value with
an "ok" label costs a wrong conclusion. When the calculator cannot vouch for a value, it
refuses it and says why.

## Install

```bash
pip install 'xl-marinade[calc]'
```

`marinade values` works on the base install. `marinade calc` needs the extra (it pins
`ironcalc==0.9.1`); without it the command exits with code 5.

## The files in

| Option | File | What it holds |
|---|---|---|
| `WORKBOOK` | `.xlsx` / `.xlsm` | The model, saved by Excel so it carries stored values. |
| `--policies` | `policies.csv` | One row per input set. An id column plus one column per mapped field. Ids stay text; a blank value or a duplicate id is an error. |
| `--map` | `map.json` | Which cell or defined name each field sets. |
| `--outputs` | repeatable | Output cells, ranges or defined names (`'Results!B5:B64'`). Ranges expand to cells. |
| `--variables` | `vars.json` | Optional. Named output variables for `rows.csv`. |
| `--check-against` | `check.json` | Optional. Stored results for other input sets, to widen the trust run. |

You must give at least one of `--outputs` and `--variables`.

### `map.json`

```json
{
  "id_field": "policy",
  "fields": {"age": "Inputs!C3", "term": "Inputs!C4", "rate": "Rate"}
}
```

`id_field` names the policies column that holds the id (default `id`). It may also appear
in `fields`, so the id is written into a workbook cell as well. Each target must be a cell
or defined name that holds a constant; mapping onto a formula is an error. Before each
input set, every mapped input is reset to its stored value. Values are converted to the
target cell's stored type, and ISO dates become Excel serials. A value outside the cell's
data-validation list refuses that input set's outputs.

### `vars.json`

```json
{
  "premium": {"range": "Proj!C11:C130", "period_start": 0},
  "reserve": {"range": "Proj!H11:H130", "period_start": 1},
  "npv":     {"cell": "Summary!B4"}
}
```

A range runs down one column or across one row; a 2-D range is an error.
**`period_start` is required on every range and has no default**, because workbooks differ
on whether periods count from 0 or 1. Periods count from it: the first cell of the range is
period `period_start`, the next one `period_start + 1`. A single `cell` has no period (the
column is empty) and must not carry `period_start`.

### `check.json`

Optional. It names results the workbook already holds for other input sets, such as a
batch-results sheet, so the trust run can recalculate those input sets too and compare.

```json
{
  "results": {"sheet": "Results", "id_column": "A", "first_row": 2,
              "outputs": {"Summary!B13": "C"}},
  "inputs": {"table": {"sheet": "Inputs", "id_column": "A", "first_row": 2,
                       "fields": {"age": "B", "term": "C"}}}
}
```

`results.outputs` maps an output cell to the results-sheet column holding its stored value
per input set. `inputs` has exactly one of `table` (a column per `map.json` field) or
`selector` (a single cell, e.g. `"Control!C4"`, that picks the input set by id). A mismatch
fails the outputs involved; only input sets the engine reproduced count towards the
exercised paths.

## How a run works

1. Read the workbook's calculation settings. 1904 dates or precision as displayed refuse
   the whole workbook (exit code 6).
2. Extract the formula graph into a temporary database, or reuse `--ir` if it matches the
   workbook.
3. Scan the formulas for [refused patterns](#refused-patterns).
4. Load a copy with every formula's stored value deleted, so the engine has to compute.
   The **trust run** compares each formula cell with what Excel stored, within
   `--tolerance` (default `1e-9`, relative or absolute), on the stored inputs and on every
   `--check-against` input set.
5. Add probe cells on a hidden sheet that report which branch, lookup row, rounding
   input and dynamic reference each input set reaches.
6. For each policy: reset, set the inputs, recalculate, read outputs and probes.
7. Walk the graph from each output back to its roots and assign a status.

The extraction database is never written by `calc`. The run clears its own known output
files from `-o` at the start (`results.csv`, `rows.csv`, `paths.csv`, `trust_report.json`,
`run.json`), so a failed run cannot leave stale ones behind. Nothing else in the directory
is touched.

## The files out

All in the `-o` directory, UTF-8, `\n` line endings, no timestamps.

| File | Contents |
|---|---|
| `results.csv` | `policy_id`, `output` (`Sheet!A1`), `value`, `status`, `reasons`. One row per policy and output cell. |
| `rows.csv` | With `--variables`: `policy`, `variable`, `period`, `value`, `status`, `reasons`. |
| `paths.csv` | `policy_id`, `cell`, `family`, `kind`, `ordinal`, `outcome`, `exercised`. The branch or lookup edge each policy took; `ordinal` tells apart several probes of one kind in one cell; `exercised` is `true` when the trust run took it too. |
| `trust_report.json` | `formula_cells`, `checked`, `reproduced`, `share`, `failures`, `refused_patterns`, `settings` (with `settings.refused`), `held_external_links`, `check_against`. |
| `run.json` | `ironcalc_version`, `marinade_version`, `workbook_sha256`, `tolerance`, `as_of`, `probe_cap`, `probes_unplaced`, and per variable its range, `period_start` and the cell period 0 falls on. |

`marinade values` writes `rows.csv` and `run.json` (variables only). Its rows carry status
`stored`, and the policy is `stored`, or the stored value of the id field when `--map`
maps it.

In `reasons`, each entry reads `kind@root:detail`, joined by `; `, in a stable order. The
root is the cell the problem starts at; it can be upstream of the output.

## Statuses

Every output gets a value. A value can carry several reasons; its status is the most
serious.

| Status | Meaning |
|---|---|
| `refused` | The output depends, through the graph, on a cell that failed the trust run or matches a refused pattern. The value is returned for inspection but must not be used. The reason names the root cell. |
| `uncertain` | The value rests on something implementation-dependent: a rounding function whose input sits within about 1e-12 relative of the point where the result flips, or a comparison or sorted lookup whose two sides are that close. Probes that could not be placed also land here ("not checked"). |
| `unverified_path` | The policy took a branch or lookup position the trust run never exercised, so those formulas were never checked against Excel. |
| `ok` | None of the above. |
| `stored` | `marinade values` only: the workbook's own stored value. |

Details worth knowing:

- Exact whole-number inputs to `INT`, `TRUNC`, `ROUNDUP`, `ROUNDDOWN`, `CEILING` and
  `FLOOR` are not uncertain (nothing can flip). Exact halves into `ROUND` and `MROUND`
  are.
- Branches are `IF`, `IFS`, `SWITCH`, `CHOOSE`, `IFERROR` and `IFNA`. A branch counts as
  exercised if any cell in the same canonical formula family took it in the trust run.
  Lookups count per family too, but the table's first row, last row and "beyond the end"
  must each have been hit.
- `OFFSET` and `INDIRECT` are followed to where they point for each input set; if the
  probe cannot be placed (cap hit, or inside `LET`/`LAMBDA`/an array context), the
  outputs behind it are never `ok`.
- A passing trust run proves only the paths the stored inputs take. That is what
  `unverified_path` and `--check-against` are for.

### Reason kinds

`silently_wrong` (engine value differs from Excel's, no error), `engine_error`,
`macro_function`, `unsupported_function`, `no_stored_value`, `refused_pattern`,
`check_against_mismatch` and `input_outside_allowed_values` give `refused`.
`near_rounding_flip`, `near_comparison` and `not_checked_value` give `uncertain`.
`unexercised_branch`, `unexercised_lookup_edge` and `not_checked_path` give
`unverified_path`. `external_link_held` is informational and leaves the status `ok`.

## Refused patterns

Some formulas reproduce Excel on the stored inputs only by luck and silently disagree on
the next. They are refused even when the trust run reproduces them. The `detail` is:

| Detail | What triggers it |
|---|---|
| `approximate lookup on unsorted column` | `VLOOKUP`/`HLOOKUP`/`MATCH`/`LOOKUP` in approximate mode over keys that are not sorted, or that are text outside plain a-z, 0-9 and space (Excel's collation differs from a plain sort). Also `XLOOKUP`/`XMATCH` with a binary-search mode, or a search mode the scan cannot read. |
| `approximate lookup on mixed text and numbers` | The key column mixes types. |
| `approximate lookup on a non-range table` | The table is not a plain reference. |
| `approximate lookup on whole column` | The table is a whole column or row. IronCalc counts empty cells; Excel skips them. |
| `implicit intersection on multi-cell name` | A single-cell use of a defined name that covers several cells. |
| `implicit intersection on multi-cell range` | A multi-cell range used outside a function argument. Ranges inside `SUMPRODUCT`, `MMULT` and similar are fine. |
| `name holds a formula` | A formula reads a defined name whose definition is a formula or a constant rather than cells (the detail text says "formula" for both). IronCalc 0.9.1 rejects such names. |
| `RAND`, `RANDBETWEEN`, `RANDARRAY` | Always. |
| `TODAY`, `NOW` | Unless `--as-of YYYY-MM-DD` is given; then they are replaced by that date. |
| `circular reference` | Always (iterative calculation is noted in the report, not honoured). |
| `data table` | What-If data tables are removed before loading and their cells refused. |
| `unparseable formula` | The formula could not be parsed. Error literals typed into a formula, such as `#N/A`, are a known source of these. |
| `external reference not held` | The formula reads an external reference the engine copy cannot hold: a range, or one with no cached value (IronCalc returns `#ERROR!` or `#REF!` for those). |

## Fixed points

| Case | Behaviour |
|---|---|
| External link with a cached single value | Replaced by the cached value. Outputs stay `ok`, with the reason `external_link_held`. |
| External link that cannot be held | Refused, see above. |
| 1904 date system | Whole workbook refused. |
| Precision as displayed | Whole workbook refused. |
| Iterative calculation | Circular cells refused as usual; the report records the setting. |
| Manual calculation mode | Recorded, not refused; stale stored values show as mismatches. |
| `TODAY`, `NOW` | Refused unless `--as-of`. |

Known IronCalc 0.9.1 gaps, each with a synthetic reproduction and the way `calc` handles
it, are listed in `ref/calc/upstream/` in the repository.

## Exit codes

| Code | Meaning |
|---|---|
| 0 | Success. |
| 1 | Error: bad input file, unknown sheet, mapping onto a formula, IR from another workbook. |
| 5 | The `[calc]` extra is not installed. |
| 6 | The workbook is refused outright (1904 dates, precision as displayed). `trust_report.json` and `run.json` are still written. |

## Example

```bash
marinade calc model.xlsx --policies policies.csv --map map.json \
    --outputs 'Summary!B4' --variables vars.json \
    --check-against check.json --as-of 2026-01-01 -o out/

marinade values model.xlsx --variables vars.json --map map.json -o stored/
```

Read `trust_report.json` first (`share`, `failures`), then filter `results.csv` or
`rows.csv` on `status`. Only `ok` rows can be used as they stand.

See the [CLI reference](../reference/cli.md) for every option.
