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
| `--policies` | `policies.csv` | One row per input set. An id column plus one column per mapped field. UTF-8 (save from Excel as "CSV UTF-8"). Ids stay text; a blank value, a duplicate id or a duplicate column name is an error. Columns with no name (the empty trailing columns some exports add) are ignored. |
| `--map` | `map.json` | Which cell or defined name each field sets. |
| `--outputs` | repeatable | Output cells, ranges or defined names (`'Results!B5:B64'`). Ranges expand to cells. |
| `--variables` | `vars.json` | Optional. Named output variables for `rows.csv`. |
| `--check-against` | `check.json` | Optional. Stored results for other input sets, to widen the trust run. |

You must give at least one of `--outputs` and `--variables`.

### Targets

Every place that names cells (`map.json` fields, `--outputs`, `vars.json`, and the
output, sheet and selector entries of `check.json`) takes a cell (`Inputs!C3`), a range
(`Proj!C11:C130`) or a defined name. Sheet names and defined names ignore case, as in
Excel.

A name can be defined for the whole workbook or for one sheet only (a sheet copied in
from another workbook often brings its own). Write a sheet's own name as
`'Sheet'!Name`. A bare name means the workbook's definition; when a sheet defines the
same name too, a bare name is refused as ambiguous, and the error says how to write each.

A whole column or row (`Calc!B:B`, `Calc!3:3`, or a name over one) covers the part of it
the sheet uses: up to the last row (or column) that holds a cell. A target that covers
no cells is an error that names it: a name that holds a formula (`OFFSET(...)` dynamic
ranges included), a constant, a broken reference (`#REF!`) or cells in another workbook,
and a whole column or row outside the range the sheet uses. Give such a target as a
plain range instead.

Inside formulas, names are read the way Excel reads them: a formula reads its own
sheet's definition of a name first, and the workbook's only when the sheet has none.

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

The list is the validation's literal items, or the stored values of its source: a range,
a whole column (the cells it holds) or a defined name over cells (the validation sheet's
own name first, as in a formula). As in Excel, the check ignores case, and numbers and
dates match by value: `2025-01-01` typed into a date cell matches a list holding that
date. What is not checked:

- A list fed by a formula (`INDIRECT`, `OFFSET`, a name holding one, as in cascading
  dropdowns) is not checked in this version: any value is accepted there.
- A list held in formula cells is read from the values the workbook stored, not
  recomputed for the input set.
- Only list validations are checked. Whole-number, decimal, date, length and custom
  validations are not.

### `vars.json`

```json
{
  "premium": {"range": "Proj!C11:C130", "period_start": 0},
  "reserve": {"range": "Proj!H11:H130", "period_start": 1},
  "npv":     {"cell": "Summary!B4"}
}
```

A range runs down one column or across one row of one sheet; a 2-D range, or a name
whose areas lie on more than one sheet, is an error.
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
`selector` (a single cell, e.g. `"Control!C4"`, that picks the input set by id). Only input
sets the engine reproduced count towards the exercised paths.

A **mismatch** (an input set whose recalculated output differs from the stored one by more
than `--tolerance`) refuses that output only for policies that took the same path to it:

- For the mismatching set, the run records its path to the output: the outcome of every
  branch and lookup-edge probe upstream of the output.
- A policy whose own outcomes for those probes are all the same gets
  `check_against_mismatch` on that output, and the output is `refused`.
- If any of those probes is unplaced, or its reading is unsure (for the set or for the
  policy), the paths count as the same and the output is refused. The rule errs toward
  refusing.
- An output with no branch or lookup probes upstream has only one path, so a mismatch on it
  refuses it for every policy.
- A policy on a different path is judged by the usual rules: `ok` if a matching input set
  (or the stored inputs) checked its path, `unverified_path` if nothing did.

`trust_report.json` lists every mismatch under `check_against.mismatches`: the input set
id (`input_set`), the output (`cell`), `stored`, `computed`, and `relative_difference`
(measured the way the tolerance is: relative, or absolute below 1).

#### When you need `--check-against`

When a workbook stores results for only one input set, give `--check-against` whenever it
holds results for others (a batch-results sheet, a scenario table). Without it, the trust
run sees only the paths the stored inputs take. A lookup that lands in the middle of a
table for a new policy (a different age row, a different year) and a branch the stored
inputs do not take stay unverified, so most new values come back `unverified_path`.
Measured on lifelib BasicTerm (10,000 model points stored in `Batch_Results`, selector
`Control!C4`): with no check set, 100% of new values came back `unverified_path`; with 200
checked points, 100% came back `ok`, and all of them were correct.

## How a run works

1. Read the workbook's calculation settings. 1904 dates, precision as displayed or a
   Strict Open XML file refuse the whole workbook (exit code 6).
2. Extract the formula graph into a temporary database, or reuse `--ir` if it matches the
   workbook.
3. Scan the formulas for [refused patterns](#refused-patterns).
4. Load a copy with every formula's stored value deleted, so the engine has to compute.
   The **trust run** compares each formula cell with what Excel stored, within
   `--tolerance` (default `1e-9`, relative or absolute), on the stored inputs and on every
   `--check-against` input set.
5. Add probe cells on a hidden sheet that report which branch, lookup row, rounding
   input and dynamic reference each input set reaches. Probes go only into cells upstream
   of the requested outputs, and never into a cell that is already refused. A probe is
   the part of the cell's formula it reports on, lifted onto that sheet so it computes
   what the cell computes: references keep their sheet (quoted where Excel needs it,
   as for `'2024'`), operators keep Excel's precedence (`A&B="x"` is `(A&B)="x"`), and
   under `--as-of` `TODAY()` and `NOW()` read the as-of date, as the cells do.
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
| `trust_report.json` | `formula_cells`, `checked`, `reproduced`, `share`, `failures`, `refused_patterns`, `settings` (with `settings.refused`), `held_external_links`, `check_against` (`input_sets`, and `mismatches` with `input_set`, `cell`, `detail`, `stored`, `computed`, `relative_difference`). |
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
| `refused` | The output depends, through the graph, on a cell that failed the trust run or matches a refused pattern, or a `--check-against` input set on the same path did not reproduce it. The value is returned for inspection but must not be used. The reason names the root cell. |
| `uncertain` | The value rests on something implementation-dependent: a rounding function whose input sits within about 1e-12 relative of the point where the result flips, or a comparison or sorted lookup whose two sides are that close (but not equal). A rounding or comparison probe that could not be placed also lands here (`not_checked_value`). |
| `unverified_path` | The policy took a branch or lookup position the trust run never exercised, so those formulas were never checked against Excel. A branch or lookup probe that could not be placed also lands here (`not_checked_path`). |
| `ok` | None of the above. |
| `stored` | `marinade values` only: the workbook's own stored value. |

Details worth knowing:

- Exact whole-number inputs to `INT`, `TRUNC`, `ROUNDUP`, `ROUNDDOWN`, `CEILING` and
  `FLOOR` are not uncertain (nothing can flip). Exact halves into `ROUND` and `MROUND`
  are.
- A comparison whose two sides are exactly equal is decided, not near: `IF(t=0, ...)`
  with `t` exactly `0` stays `ok`. Only sides that differ by more than nothing and at
  most about 1e-12 relative are `near_comparison`. The same holds for an approximate
  lookup whose value equals its matched key.
- Branches are `IF`, `IFS`, `SWITCH`, `CHOOSE`, `IFERROR` and `IFNA`, and an `XLOOKUP`
  with an `if_not_found` argument (outcomes `found` and `not_found`; the argument counts
  only when nothing was found). A branch counts as exercised if any cell in the same
  canonical formula family took it in the trust run. Lookups (`VLOOKUP`, `HLOOKUP`,
  `MATCH`, `LOOKUP`, `XLOOKUP` and `XMATCH`, in every match and search mode calc accepts)
  count per family too, but the table's first row, last row and "beyond the end" must
  each have been hit. For `XLOOKUP` and `XMATCH` first and last are positions in the
  lookup array, whichever way the search runs.
- Only evidence that was checked counts. A cell that failed the trust run, or is a
  refused pattern, exercises nothing for its family. A branch or lookup inside an arm
  the formula did not evaluate (the `FALSE` arm of an `IF` that went `TRUE`) is not
  exercised by the trust run and is not a path a policy took. A `--check-against` input
  set exercises only the branches and lookups upstream of the outputs it was compared on.
- Because probes are placed only upstream of the requested outputs, asking for fewer
  outputs can make a status more conservative: a family member that is not upstream of
  any requested output is not probed, so it cannot vouch for the branches it took.
- A probe that cannot be placed (cap hit, or inside `LET`/`LAMBDA`/an array context)
  never yields `ok`: a branch or lookup probe gives `unverified_path`, a rounding or
  comparison probe `uncertain`, and a dynamic-reference probe `refused`.
- `OFFSET` and `INDIRECT` are followed to where they point for each input set; if that
  cannot be read, the outputs behind it are refused (`dynamic reference not checked`).
  An `INDIRECT` text that addresses a cell off the grid (row 0, below row 1,048,576 or
  right of column `XFD`) counts as unread: Excel gives `#REF!` there, while IronCalc
  0.9.1 reads `Rates!B0` as 0.
- An error code Excel does not have (IronCalc's `#N/IMPL!`, `#ERROR!`, `#CIRC!` and the
  like) in an output or in a probe is an `engine_error` for that policy, so the output is
  `refused`, even when an `IFERROR` around it would have turned it into a plausible
  value. Excel's own error codes (`#N/A`, `#DIV/0!`, `#VALUE!` and the rest) are values
  like any other, with one exception: a `#NAME?` in an output for a policy, where the
  trust run had none in that cell, is an `engine_error`. Excel's `#NAME?` comes from
  the formula's text, not its inputs, so one only a new policy reaches is a function
  the engine lacks on an arm the trust run never took.
- A passing trust run proves only the paths the stored inputs take. That is what
  `unverified_path` and `--check-against` are for.

### Reason kinds

`silently_wrong` (engine value differs from Excel's, no error), `engine_error`,
`macro_function`, `unsupported_function`, `no_stored_value`, `refused_pattern`,
`check_against_mismatch` and `input_outside_allowed_values` give `refused`.
`near_rounding_flip`, `near_comparison` and `not_checked_value` give `uncertain`.
`unexercised_branch`, `unexercised_lookup_edge` and `not_checked_path` give
`unverified_path`. `external_link_held` is informational and leaves the status `ok`.

A `#NAME?` from the engine is `macro_function` when the cell, or any cell upstream of it,
calls a VBA function (the engine cannot run VBA); the root is that UDF cell, the smallest
one if several feed it. It is `unsupported_function` only when no UDF is involved: the
engine does not know a function the cell uses.

## What `ok` means, and what it does not

`ok` means every formula and every path behind the value was checked against values Excel
stored: on the stored inputs, or on a `--check-against` input set that took the same
branches and lookup edges. It does not guarantee that a new input set's value is within
`--tolerance` of what Excel would compute for it. Floating-point differences between the
engine and Excel are tiny, but cancellation can amplify them: a net figure that is a small
difference of large present values can move by more than 1e-9 relative even when every
formula reproduced. On lifelib BasicTerm_SE one model point's net cash-flow present value
came out at 1.003e-9 relative to Excel's, against the default tolerance of 1e-9. Choose
`--tolerance` for the outputs you compare, and treat agreement with Excel beyond it as
likely, not proven.

## Memory and output size

`results.csv`, `paths.csv` and `rows.csv` are written one policy at a time, flushed as each
policy finishes, so memory does not grow with the number of policies. Nor does it grow
with the number of `--check-against` input sets: each matching set's branch and lookup
outcomes are folded into the exercised set as it finishes and its readings dropped; a
mismatching set keeps only its path to the output it missed. Peak memory is set by the
workbook (the formula graph, the engine's copy of the workbook and the probe plan) and by
the output cells you request: a range is expanded cell by cell, so `Calc!B1:B1048576`
costs memory for a million cells. A whole column (`Calc!B:B`) expands only over the rows
the sheet uses.

The files on disk do grow with the policies. `paths.csv` has one row per policy per
branch or lookup probe it reached, which for an actuarial projection can be over a
thousand rows per policy (about 1.2 million rows and 126 MB for 1,000 lifelib BasicTerm_S
policies). Leave room in `-o` for it.

A run that stops part-way (a bad value in `policies.csv`, for instance) deletes
`results.csv`, `paths.csv` and `rows.csv`, so partial results never look complete.

## Refused patterns

Some formulas reproduce Excel on the stored inputs only by luck and silently disagree on
the next. They are refused even when the trust run reproduces them. The `detail` is:

| Detail | What triggers it |
|---|---|
| `approximate lookup on unsorted column` | `VLOOKUP`/`HLOOKUP`/`MATCH`/`LOOKUP` in approximate mode over keys that are not sorted, or that are text outside plain a-z, 0-9 and space (Excel's collation differs from a plain sort). Also `XLOOKUP`/`XMATCH` with a binary-search mode, or a search mode the scan cannot read. |
| `approximate lookup on mixed text and numbers` | The key column mixes types. |
| `approximate lookup on a non-range table` | The table is not a plain reference. |
| `approximate lookup on whole column` | The table is a whole column or row, or a name over one. IronCalc counts empty cells; Excel skips them. |
| `implicit intersection on multi-cell name` | A single-cell use of a defined name that covers several cells (a name over a whole column or row always does). |
| `implicit intersection on multi-cell range` | A multi-cell range used outside a function argument. Ranges inside `SUMPRODUCT`, `MMULT` and similar are fine. |
| `name holds a formula` | A formula reads a defined name whose definition is a formula or a constant rather than cells (the detail text says "formula" for both). IronCalc 0.9.1 rejects such names. |
| `RAND`, `RANDBETWEEN`, `RANDARRAY` | Always. |
| `TODAY`, `NOW` | Unless `--as-of YYYY-MM-DD` is given; then they are replaced by that date. |
| `circular reference` | Always (iterative calculation is noted in the report, not honoured). |
| `data table` | What-If data tables are removed before loading and their cells refused. |
| `unparseable formula` | The formula could not be parsed. Error literals typed into a formula, such as `#N/A`, are a known source of these. |
| `external reference not held` | The formula reads an external reference the engine copy cannot hold: a range, one with no cached value (IronCalc returns `#ERROR!` or `#REF!` for those), or a defined name that points into another workbook. |
| `approximate lookup on input-dependent keys` | An approximate lookup whose key cells change with a mapped input (an input cell, or a formula that depends on one). The keys may be sorted for the stored inputs and not for the next policy. |
| `function IronCalc does not implement: NAME` | The formula calls an Excel function IronCalc 0.9.1 does not have: `AGGREGATE`, `BAHTTEXT`, `CALL`, the `CUBE…` functions, `DBCS`, `DETECTLANGUAGE`, `EUROCONVERT`, `FIELDVALUE`, `FILTERXML`, `GETPIVOTDATA`, `GROUPBY`, `IMAGE`, `JIS`, `PHONETIC`, `PIVOTBY`, `REGISTER.ID`, `RTD`, `SQL.REQUEST`, `STOCKHISTORY`, `TRANSLATE`, `WEBSERVICE`. The engine answers `#NAME?`, an Excel error code, so an `IFERROR`, `ISERROR` or `ISNUMBER` around the call would turn it into a plausible value. |

One more refusal is per policy, read from the probes: `lookup value type differs from
keys`, when an approximate lookup matches a key of a different type than its lookup value
(text against numbers). Excel finds nothing there (`#N/A`); IronCalc 0.9.1 returns a row.

## Fixed points

| Case | Behaviour |
|---|---|
| External link with a cached single value | Replaced by the cached value. Outputs stay `ok`, with the reason `external_link_held`. |
| External link that cannot be held | Refused, see above. |
| 1904 date system | Whole workbook refused. |
| Precision as displayed | Whole workbook refused. |
| Strict Open XML (`.xlsx` saved as "Strict Open XML Spreadsheet") | Whole workbook refused (`strict OOXML`). |
| Iterative calculation | Circular cells refused as usual; the report records the setting. |
| Manual calculation mode | Recorded, not refused; stale stored values show as mismatches. |
| `TODAY`, `NOW` | Refused unless `--as-of`. |

Known IronCalc 0.9.1 gaps, each with a synthetic reproduction and the way `calc` handles
it, are listed in `ref/calc/upstream/` in the repository.

## Exit codes

| Code | Meaning |
|---|---|
| 0 | Success. |
| 1 | Error: bad input file, unknown sheet or name, a target that covers no cells, mapping onto a formula, IR from another workbook. |
| 5 | The `[calc]` extra is not installed. |
| 6 | The workbook is refused outright (1904 dates, precision as displayed, Strict Open XML). `trust_report.json` and `run.json` are still written. |

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
