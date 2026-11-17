# Recalculating a workbook for new inputs (`marinade calc`): engine comparison

Status: **investigation only. Nothing is built.** This note records the evidence behind
choosing an engine, and the plan it points to.

## The question

XL Marinade extracts a workbook's formulas, stored values and dependency graph. It does
not recalculate. A recalculator would answer: given new values for some input cells,
what do the output cells become?

A workbook stores results only for the inputs sitting in it. Comparing anything against
"what the workbook says" is therefore limited to those inputs, unless the workbook can be
recalculated for others.

## What the tool would have to do

1. **Recalculate.** Given a workbook, a list of input sets, a mapping from fields to
   input cells, and the output cells or ranges wanted, return the outputs per input set
   as a plain table. Offline, deterministic, no spreadsheet application, no network.
2. **Trust itself first.** Before use on a workbook, reproduce that workbook's own stored
   values from the original inputs by a recalculation it really performs. Report the
   share reproduced and every cell it could not compute and why. Refuse, or mark the
   affected outputs, where it cannot.
3. **Say what it is unsure of.** Mark an output as uncertain where its value rests on
   something implementation-dependent, rounding ties above all.
4. **Help choose inputs.** List the input cells that feed a condition or a lookup, and
   the values that reach each branch or table edge.

## Decision

Wrap [IronCalc](https://github.com/ironcalc/ironcalc) (Apache-2.0 / MIT) behind
Marinade's interface. Do not write our own evaluator, and do not change Marinade's
licence.

- `xl_marinade.calc` module and a `marinade calc` command, depending on `core` and never
  the reverse.
- IronCalc as an optional extra (`xl-marinade[calc]`); the feature labelled experimental
  and outside the versioned output contract until it settles.
- The calculator never writes to the extraction database.
- Refuse by default: a wrong number does more harm than a missing edge.

The two other evaluators tried, `formulas` (EUPL-1.1+) and `pycel` (GPLv3), are copyleft
and cannot be in the install closure (`.github/workflows/licence-check.yml`). `pycel` has
had no release since 2021 and was not run here.

## Evidence

Run on 2026-11-17 with `ironcalc` 0.9.1 and `formulas` 1.3.4, on the five lifelib example
workbooks. They were saved by Excel, so their stored values are true Excel results. The
workbooks are not in this repository; the scripts and raw results are in `comparison/`.

**Trust test.** Each engine was given a copy of the workbook with every formula cell's
stored value deleted, so it had to compute. Reproduced means equal within 1e-9.

| Workbook | Formula cells | IronCalc | IronCalc time | `formulas` | `formulas` time |
|---|---|---|---|---|---|
| BasicTerm_S | 4,605 | 100% | 0.2 s | 100% | 38 s |
| BasicTerm_SE | 6,679 | 100% | 0.2 s | 100% | 44 s |
| CashValue_SE | 64,094 | 100% | 0.5 s | 100% | 386 s |
| SmithWilson (array formulas, `MMULT`, `MINVERSE`) | 2,209 | 100% | 0.03 s | 92.8% | 15 s |
| Solvency2 | 22,123 | 100% | 0.4 s | 19.7% | 73 s |

**New inputs against real Excel.** BasicTerm_S and BasicTerm_SE each hold 10,000 model
points, a selector cell, and a sheet of results that Excel's own batch macro produced for
every one. IronCalc, on all 20,000:

- the rounded premium is exactly equal to Excel's on all 20,000;
- the net present value agrees within 1e-9 relative (it is a difference of large numbers;
  the component present values match the lifelib Python model to about 1e-12);
- 6-8 ms per model point.

**Function probe.** A synthetic workbook of 96 checked cases: conditionals, lookups,
sums and counts, rounding, `NPV`/`IRR`/`PMT`, the normal distribution, dates and text.
Expected values are documented Excel results or computed independently.

| | IronCalc | `formulas` |
|---|---|---|
| Correct | 95 of 96 | 89 of 96 |
| Unsupported | `AGGREGATE` | `OFFSET`, `INDIRECT`, `SUBTOTAL`, `AGGREGATE` |
| Documented rounding ties (11 cases, e.g. `ROUND(1.005,2)` = 1.01) | all correct | 2 wrong |
| Circular reference | `#CIRC!` | 0 when told to allow circular references |
| Call to a macro function | `#NAME?`, carried to dependants | same |

## What the evidence does not show

- **Computed rounding ties are unverified.** None of the 20,000 premiums landed on a tie
  (the closest was 0.00005 cents from a half-cent). The uncertainty mark is still needed.
- **The lifelib workbooks use a narrow set of functions.** The probe widens that, but its
  expected values are ours.
- **Passing the trust test proves only the branches the stored inputs take.** A new input
  that takes the other branch of an `IF`, or lands on a different lookup row, runs
  formulas the test never checked. Outputs on such paths need their own mark.
- **External links with no cached value** come back as `#REF!` from IronCalc. A wrapper
  has to supply the cached value itself.
- **IronCalc is pre-1.0 and has no hook for custom functions** in its Python interface.
  A function it lacks means refusing the affected outputs or contributing upstream.

## Plan: three pull requests

1. **Calculator and trust report.** Recalculate a list of input sets into a plain table;
   the trust report; refusal downstream of macro functions, `RAND`, circular references
   and unsupported functions; a mark on outputs that pass through a `ROUND` near a tie; a
   mark on outputs that leave the paths the trust run exercised; a report of which
   branches and lookup rows each input set exercised.
2. **Input-selection helper.** Requirement 4 in full: suggest input values that reach
   each branch and table edge. Needs argument-level parsing (`core/parser.py`).
3. **Stand-ins for macro functions.** Compute a workbook's macro function in Python, write
   the result into the cell, and recalculate in dependency order using Marinade's graph.

## What Marinade already has to build on

- `core/parser.py`: formula parser that builds a syntax tree.
- `core/levels_cycles.py`: evaluation levels and cycle detection.
- `agent_dependencies`: every cell, range and external edge, including reads of empty
  cells (since 0.7.1).
- `cell_udf_calls` and `udfs`: which cells call macro functions.
- `is_array_formula` (set since 0.7.1). `is_spilled` is not populated.
- `data_validations`, `defined_names`, `agent_table_candidates`.
- Not recorded yet, and a calculator may need them: the 1900/1904 date system, whether
  iterative calculation is on, and "precision as displayed".

## Open design questions

- Command line first, Python library first, or both as documented interfaces?
- Does `marinade calc` take the workbook, the extracted database, or both? IronCalc needs
  the workbook file; refusal and path marking need the graph.
- The shape of the input mapping and of the output selection (cells, ranges, names).
- The shape of the trust report and of per-output marks, and whether a failed trust test
  blocks a run or only marks it.
- How "near a tie" is defined for the `ROUND` mark.
- What as-of date `TODAY`/`NOW` use.
