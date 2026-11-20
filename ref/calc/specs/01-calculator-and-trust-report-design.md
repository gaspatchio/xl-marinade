# `marinade calc`, PR 1: calculator and trust report — design

Status: **design, approved 2026-11-19. Nothing is built.** Background, engine comparison and
the three-PR plan are in [`../README.md`](../README.md).

`marinade calc` recalculates a workbook for new input sets with IronCalc, proves itself
against the workbook's own stored values first, and puts a status on every output value.
`marinade values` returns a workbook's stored values in the same shape, without
recalculating.

## Why

A workbook stores results only for the inputs sitting in it, often a single policy.
Checking anything against "what the workbook says" is limited to those inputs. `calc`
returns the workbook's answer for input sets it does not hold, and says how far each
answer can be trusted.

## Decisions

| Question | Options considered | Decision |
|---|---|---|
| Engine | write our own evaluator; wrap `formulas`/`pycel`; wrap IronCalc | IronCalc (`ironcalc` on PyPI, Apache-2.0/MIT), as the optional extra `xl-marinade[calc]`, labelled experimental |
| Interface | command line first; Python library first; both equal | Command line first, files in and out; a Python `calc()` underneath, as with `extract()` and `document()` |
| What a trust failure refuses | each affected output; the whole workbook; per output plus a minimum share | Each output that depends on a failing or refused cell |
| An input set that takes a branch the trust run never exercised | mark and return; refuse; refuse some cases | Mark it "unverified path" and return the value |
| What counts as an exercised branch | per cell; per formula | Per formula (see Statuses) |
| What `calc` takes | workbook; extraction database; both | The workbook; it extracts its own graph, or reuses `--ir` if the file hash matches |
| Values as variable/period rows | cells only; a variable map | A variable map, used by both `calc` and `values` |
| Extra stored results for other input sets | later; PR 1 | PR 1, as `--check-against` |

## Goals, non-goals, anti-goals

**Goals.** Recalculate a list of input sets offline and deterministically. A trust report.
A status on every output value. A record of which branches and lookup rows each input set
took. Variable/period rows from recalculated and stored values through one map.

**Non-goals (later PRs).** Suggesting input values that reach each branch and table edge
(PR 2). Stand-ins for a workbook's macro functions (PR 3); PR 1 refuses outputs that depend
on them. Writing any downstream tool's expected-values file.

**Anti-goals.**
- A wrong number with an "ok" status. A refusal costs a missing comparison; a silent wrong
  value costs a wrong conclusion.
- Writing to the extraction database, or changing its versioned contract. `calc` stays
  outside SemVer while experimental.
- Crossing the tier boundary. `calc` depends on `core`, never the reverse; `core` and `docs`
  stay importable without IronCalc.

## Trust run

IronCalc recalculates a copy of the workbook with every formula cell's stored value
deleted, so it cannot return a stored value. Each formula cell is compared with what Excel
stored, within a tolerance (default 1e-9 relative or absolute). A cell that does not match
is **failing**, with a reason, silently wrong first:

1. wrong value, no error (silently wrong)
2. engine error (`#VALUE!`, `#ERROR!` and the like)
3. macro function (`#NAME?` on a function the workbook defines in VBA)
4. function IronCalc lacks
5. no stored value (the cell cannot be checked)

**Refused patterns.** Some cells can match on the stored inputs only by luck. They are
refused even when they reproduce. Marinade's parser finds them without running anything:

- approximate `VLOOKUP` / `MATCH` / `LOOKUP` on a column that is not sorted, or that mixes
  text and numbers
- a single-cell reference to a defined name covering several cells (implicit intersection)
- `RAND`, `RANDBETWEEN`; `TODAY`, `NOW` without `--as-of`
- circular references
- What-If data tables: removed before loading (IronCalc will not load a file that has
  them), and their cells refused

**Extra stored results.** `--check-against check.json` names results a workbook already
holds for other input sets, such as a batch-results sheet: the results sheet, its id column,
which result column matches which output, and where each input set's inputs come from
(an input table with a column per map field, or a selector cell that picks the input set).
The trust run recalculates each of those input sets and compares. Branches and lookup rows
they reach count as exercised. A mismatch fails the outputs involved for every policy that
took the mismatching input set's path to them: the same outcome at every branch and
lookup-edge probe upstream of the output. An unplaced or unsure probe counts as the same
path, and an output with no such probes has one path, so the mismatch fails it for all
policies (agreed 2026-11-20).

## Statuses

Every output value is filled in. A value can carry several reasons; its status is the
most serious.

| Status | When |
|---|---|
| refused | It depends, through the graph, on a failing or refused cell. The reason names the root cell. |
| uncertain | It passes through a rounding function (`ROUND`, `ROUNDUP`, `ROUNDDOWN`, `MROUND`, `INT`, `TRUNC`, `CEILING`, `FLOOR`) whose input is within about 1e-12 relative of the point where the rounding flips; or through a comparison in a branch, or a lookup on a sorted column, whose two sides are that close. |
| unverified path | The input set took a branch or lookup position the trust run never exercised. The reason lists them. |
| ok | None of the above. |

**Branches** are `IF`, `IFS`, `SWITCH`, `CHOOSE`, `IFERROR` and `IFNA`. A branch counts as
exercised if any cell in the same canonical formula family took it in the trust run
(including `--check-against` input sets). Lookups count per formula family too, except
that the table's first row, last row and "beyond the end" must each have been hit to count.

**Dynamic references.** `OFFSET` and `INDIRECT` can point elsewhere for new inputs than the
static graph records. A probe cell reports where each one points for each input set, and
the status walk follows the reference actually used. If a probe cannot be placed, the
outputs behind it are refused. `INDEX` needs no probe: its result stays inside a range the
graph already holds.

**Probes that cannot be placed** (a cap is hit, or the expression sits inside `LET`,
`LAMBDA` or an array context) never yield "ok": a branch probe gives "unverified path", a
rounding or comparison probe gives "uncertain", with the reason "not checked".

## Fixed points

| Case | Behaviour |
|---|---|
| `RAND`, `RANDBETWEEN` | Always refused. |
| `TODAY`, `NOW` | Refused unless `--as-of` is given; then replaced by that date. |
| External links | Only the external reference is replaced, by the value Excel cached for it inside the workbook. Outputs behind it stay "ok" with the reason "external link held at cached value". |
| 1904 date system | Whole workbook refused. |
| Precision as displayed | Whole workbook refused. |
| Iterative calculation | Circular cells refused as usual; the report notes the setting. |
| Manual calculation mode | Noted in the report, not refused; stale stored values show up as mismatches. |

A whole-workbook refusal is a typed error with a non-zero exit. Only `trust_report.json`
and `run.json` are written.

Marinade does not yet record the date system, iterative calculation, precision as
displayed or calculation mode; PR 1 reads them from the workbook.

Tested on 2026-11-19 against synthetic workbooks with IronCalc 0.9.1:

| Case | Excel | IronCalc |
|---|---|---|
| 1904 date system, `DATE(2020,1,1)` | 42369 | 43831 (setting ignored, no error) |
| Precision as displayed, `=1/3` shown as `0.00`, times 3 | 0.99 | 1.0 (setting ignored, no error) |
| Iterative calculation, `A1 = 0.5*A1+1` | about 2 | `#CIRC!` |
| `RAND` | random | a new value on every recalculation |
| `TODAY` replaced by `DATE(2026,1,1)` | 46023 | 46023 |
| External reference with no cached value | cached value | `#REF!` |
| Approximate `VLOOKUP` of text against numbers | `#N/A` | a number |
| Single-cell reference to a column-range name | the cell on the same row | `#VALUE!` |
| `SUM(S!$A$1:S!$A3)` | 6 | `#ERROR!` |

## Inputs

```
marinade calc book.xlsx --policies policies.csv --map map.json \
    [--outputs 'Results!B5:B64' ...] [--variables vars.json] [--check-against check.json] \
    [--ir book.db] [--as-of 2026-01-01] [--tolerance 1e-9] -o out/

marinade values book.xlsx --variables vars.json [--map map.json] [--ir book.db] -o out/
```

- `policies.csv`: one row per input set, with an id column.
- `map.json`: each field mapped to a cell or a defined name. The target must hold a constant;
  mapping onto a formula is an error.
- Before each input set, every mapped input is reset to its stored value. A blank field is
  an error. Values are converted to the target cell's stored type; ISO dates become Excel
  serials. A value outside a cell's data-validation list refuses that input set's outputs.
- `--outputs`: cells, ranges or defined names; ranges are expanded to cells.

**Variable map.** One file serves `calc` and `values`, so recalculated and stored rows
carry identical keys:

```json
{
  "premium":     {"range": "Proj!C11:C130", "period_start": 0},
  "claims":      {"range": "Proj!D11:D130", "period_start": 0},
  "reserve_end": {"range": "Proj!H11:H130", "period_start": 1},
  "npv":         {"cell":  "Summary!B4"}
}
```

A range runs down one column or across one row; a 2-D range is an error. `period_start` is
required on every range, with no default, because workbooks differ on whether periods count
from 0 or 1. A single cell has an empty period. `values` reads stored values from the
extraction, needs no IronCalc and works on the base install. The stored input set's id comes
from `map.json`'s id field if mapped, else `stored`.

## Outputs (in `out/`)

- `results.csv`: one row per input set and output cell: `policy_id`, `output` (`Sheet!A1`),
  `value`, `status`, `reasons`.
- `rows.csv` (with `--variables`): `policy`, `variable`, `period`, `value`, `status`,
  `reasons`. From `values`, status is `stored`. Rows only: no downstream file format.
- `trust_report.json`: share reproduced, every failing cell with its reason (silently wrong
  first), refused patterns found, the root cell behind each refusal, workbook settings.
- `paths.csv`: per input set, the formula family and the branch or lookup edge taken, each
  marked exercised or not.
- `run.json`: IronCalc and Marinade versions, workbook hash, tolerance, as-of date, and for
  each variable its range, `period_start` and the cell period 0 falls on.

## How a run works

1. Read workbook settings; refuse the whole workbook if needed.
2. Extract the graph into a temporary database, or reuse `--ir`.
3. Scan formulas with the parser for refused patterns.
4. Write a copy with stored values deleted, data tables removed and external references
   replaced by cached values.
5. Add probe cells on a hidden sheet, per formula family where possible: branch conditions,
   lookup positions, rounding inputs and comparisons, dynamic references. IronCalc computes
   them, so there is no evaluator of our own. (IronCalc's Python interface can add a sheet and
   formulas, and `CELL("address", OFFSET(...))` and `ROWS`/`COLUMNS` report where a dynamic
   reference points.)
6. Trust run on the stored inputs and any `--check-against` input sets.
7. For each input set: reset and set inputs, recalculate, read outputs and probes.
8. Assign statuses by walking the graph, plus probed dynamic references, from each output to
   its roots.
9. Write the output files.

Loading once and recalculating per input set took 6-8 ms per input set on the lifelib
workbooks, before probes.

## Left to the implementation

- Module layout inside `xl_marinade.calc`; the exact JSON shape of `trust_report.json` and
  `run.json`.
- How "sorted" and "mixed-type" are tested for an approximate-lookup column.
- Probe naming, placement, and the cap for very large workbooks.
- Error types and exit codes, in the `errors.py` hierarchy.
- Tests on synthetic workbooks only (openpyxl or `test_workbook_generator/`), one per
  refused pattern, status and fixed point.
- Upstream IronCalc issues for the silently wrong cases, each with a synthetic reproduction.
- Packaging: the `[calc]` extra, the licence check, the changelog, the bundled skill section,
  and `docs/`.
