---
name: using-xl-marinade
description: Use XL Marinade (deterministic Excel formula-graph extraction to SQLite, binding-level dependency queries, structural diff) when authoring, auditing, modifying, or verifying Excel workbooks — especially actuarial/financial models. Triggers - building a workbook that must be provably correct, auditing or changing an existing (human-built) workbook, tracing what drives a cell/output, attributing a numeric change, comparing two workbook versions, checking a workbook's calculation structure against a source model (Python/lifelib), or working with VBA-bearing xlsm files.
---

# Using XL Marinade — agent guide

XL Marinade extracts a workbook's calculation structure into a SQLite "IR"
database: cells, formulas, binding-level groups, a dependency graph, and
parsed VBA. It is an **auditor/tracer** — `extract` snapshots cached values and
never recalculates. The one exception is the experimental `marinade calc` (below),
which recalculates with a separate engine and labels how far to trust each value.

Install and run it with `uv` (or any venv where `xl-marinade` is installed):

```bash
uv run marinade --help          # extract · document · diff · calc · values
```

Pick the reference for your task BEFORE starting:

- **Authoring a new provably-correct workbook** → `references/authoring.md`
  (author-for-the-extractor conventions, value/structure proof patterns, VBA).
- **Auditing or modifying an existing workbook** → `references/auditing.md`
  (recon SQL, version-chain diffs, mechanism-vs-data separation, attribution).
- **Anything that opens Excel or edits a file via openpyxl** →
  `references/excel-ops.md` (recalc traps, invisible modals, corruption
  hazards, single-instance discipline).

## Extract

```bash
marinade extract path/to/book.xlsx -o out.ir.db      # .xlsx or .xlsm
marinade extract path/to/book.xlsm -o out.ir.db --max-memory-mb 4000
```

- Works on `.xlsx` and `.xlsm`; VBA is parsed into the IR by every install
  (0.7.0+, no extra needed). Empty VBA tables mean "no macros" ONLY if
  `SELECT value FROM ir_metadata WHERE key = 'vba_extraction'` is
  `no_vba_project`; `failed` means they may be partial. A db stamped
  `skipped_no_oletools` came from 0.5.0–0.6.x without the `[vba]` extra: the
  workbook HAS macros that were never read — re-extract with 0.7.0+ before
  auditing or diffing VBA.
- **Recalculate in Excel and save before extracting** if you will read
  cached *values* — openpyxl saves STRIP cached values. Formula/edge
  *structure* does not need a recalc.
- Extraction aborts above `--max-memory-mb` (default 1800) to prevent OOM;
  raise it for large books. Output is one SQLite file — query with `sqlite3`.
- Runtime scales with the book: a generated ~65k-cell workbook extracts in
  seconds; a large human-built model (thousands of bindings) takes minutes —
  run it in the background, keep one `.ir.db` per version in scratch, and
  don't commit ad-hoc dbs.

## Query

**Query the `agent_*` / `marinade_*` VIEWS, never the base tables.** The views
are the extractor's versioned public contract. Read the database's own version
rather than assuming one —
`SELECT value FROM ir_metadata WHERE key = 'schema_version'` — and pin to the
major. It is currently `"3.0"`; `marinade_nodes` was called `atlas_nodes` below
3.0, so a database stamped `"2.0"` has the old view name. Base tables (`cells`,
`formulas`, `cell_edges_internal`, `range_edges`, …) are internal storage and
can change shape between releases without notice. If a query you need isn't covered by a view, treat that as a
gap to report, not a reason to reach into base tables.

**Direction convention:** `from_*` = the formula (the dependent); `to_*` =
what it reads (the precedent). "What drives X" = rows WHERE `from_*` is in X.

On workbooks **authored to the conventions** (machine-key row etc.) the
`agent_bindings` / `agent_binding_dependencies` views are the entry point:
65k cells collapse to ~100–250 labeled bindings. On human-built books labels
degrade and bindings number in the thousands — use `agent_bindings` for
orientation only and do audit work at cell level (`agent_cells_light`,
`agent_dependencies`).

```sql
-- model map: one row per detected column/block with auto-derived label
SELECT sheet, address, formula_pattern, label FROM agent_bindings;

-- what drives a given output (precedents; DISTINCT — one row per edge)
SELECT DISTINCT to_label, to_address FROM agent_binding_dependencies
WHERE from_address LIKE 'Summary!%';

-- ground truth when an edge looks wrong: the stored formula for a cell
SELECT cell_address, formula, formula_r1c1 FROM agent_cells_light
WHERE cell_address = 'Projection!C3';

-- cross-sheet consumer census ("is this tab dead? what breaks if I replace
-- it?"). agent_dependencies already UNIONs cell + range + external edges,
-- so one query does it — no manual union of edge tables:
SELECT substr(from_cell, 1, instr(from_cell, '!') - 1) AS from_sheet,
       count(*) AS edges
FROM agent_dependencies
WHERE to_cell LIKE 'Projection!%' AND from_cell NOT LIKE 'Projection!%'
GROUP BY 1 ORDER BY 2 DESC;

-- find every consumer of a defined name / anchor before a wide edit
-- (defined names survive verbatim in the formula text):
SELECT cell_address, formula FROM agent_cells_light
WHERE formula_r1c1 LIKE '%MortRates%';

-- a range edge with its resolved bounds (whole-row/col refs arrive
-- pre-expanded, e.g. A1:A1048576):
SELECT from_cell, to_cell, cell_count, to_r1, to_c1, to_r2, to_c2
FROM agent_dependencies WHERE dependency_type = 'range';

-- VBA is structure too: procedures appear in the unified node surface
SELECT node_id, display_name FROM marinade_nodes WHERE node_kind = 'procedure';
```

Schema notes — when a query errors, list the view's columns with
`PRAGMA table_info(<view>)`; never guess twice:

- `agent_cells_light` (cheap default) / `agent_cells` (adds `value`/`format`
  JSON) — key columns `cell_address` (`Sheet!A1`), `sheet`, `row`, `col`,
  `formula` (A1), `formula_r1c1`, `data_type`, `is_array_formula` (1 on
  the top-left cell of an array formula; always 0 in databases extracted
  before 0.7.1, and `is_spilled` is always 0).
- `agent_bindings` — `sheet`, `address` (`Sheet!A1:A10`), `formula_pattern`
  (R1C1; NULL for constant/heterogeneous blocks), `label`, `binding_type`,
  `classification`, `confidence`.
- `agent_dependencies` — `from_cell`, `to_cell`, `dependency_type`
  (`cell` | `range` | `external`), `cell_count` (ranges), `to_r1/c1/r2/c2`.
  A `to_cell` can be an empty cell with no row in `agent_cells_light`: the
  read is real (a blank input, a placeholder sheet). Databases extracted
  before 0.7.1 omit those cell edges; re-extract if a read seems missing.
- `agent_binding_dependencies` — binding-level graph with `from_label`,
  `from_address`, `to_label`, `to_address`, `edge_count`, `kind`
  (`formula` | `via_vba_paste`).
- `marinade_nodes` — bindings AND VBA procedures in one list: `node_kind`
  (`cell` | `procedure`), `display_name`. VBA IDs are
  `vba::<module>::<name>::<kind>` — split on `::`.
- Formula-text search lives in `agent_cells_light.formula` /
  `.formula_r1c1`. Defined names appear verbatim (not resolved to refs).

## Diff two workbook versions

```bash
marinade diff v1.ir.db v2.ir.db -o diff.json     # earlier db first, later second
```

One JSON document, three layers to read in order:

- **`summary`** — the count vector (`bindings_formula_changed`,
  `cells_formula_changed`, `cols_inserted`, `ir_inference_changes`, …).
  Read this FIRST; it's the changelist overview.
- **`binding_map`** — one row per binding with `diff_state`
  (`unchanged` | `modified` | `added` | `removed` | `moved`), `address_a/b`,
  `label_a/b`. This labeled map is the agent-consumable changelist. A
  binding that only shifted because rows/columns were inserted or deleted is
  `unchanged` with `shifted: true` and differing `address_a/b` — not an
  edit; filter on `diff_state`, not on address. One deleted with its rows/columns
  is `removed` (`BINDING_REMOVED` carries `with_axis_deletion: true`).
  `summary.{rows,cols}_{inserted,deleted}` count rows/columns, not events.
- **`changes`** — the per-event detail. Each carries `seq`, `type`, and a
  **`layer`** (always present — filter on it, never on its absence):
  - `layer = "workbook"` — real workbook edits (`FORMULA_CHANGED`,
    `VALUE_CHANGED`, `BINDING_FORMULA_CHANGED`, `BINDING_ADDED`,
    `COLS_INSERTED`, and `*_EDGE_*` out of a cell whose formula changed).
    `BINDING_FORMULA_CHANGED` carries `modification_kind` (`logic_change` vs
    a reference shift) — the logic changes are your review set.
  - `layer = "ir_inference"` — extractor-derived, NOT workbook edits
    (`TABLE_CANDIDATE_*`, `BINDING_LABEL_EVIDENCE_CHANGED`,
    `BINDING_METADATA_CHANGED`, and `*_EDGE_*` out of a cell whose formula
    did NOT change — e.g. a lookup re-resolved because its key's value
    moved). These churn around a real edit even though the cells they name
    were not touched; `summary` counts them as `ir_inference_changes`.
    **Never count them as edits in an audit.**
  - VBA edits are `layer = "workbook"` too, after everything else:
    `VBA_MODULE_ADDED|REMOVED`, `VBA_PROCEDURE_ADDED|REMOVED|RENAMED|
    RENAMED_AND_MODIFIED|LOGIC_CHANGED|SIGNATURE_CHANGED|COSMETIC_CHANGED`,
    `VBA_DECLARATIONS_CHANGED|COSMETIC_CHANGED` (module code outside
    procedures), counted in `summary.vba_*`. Each names `module`,
    `procedure`, `kind`, `compile_branch` and `lines` (`lines_a`/`lines_b`
    for a matched pair) — read the
    bodies from `vba_procedures` in each db. `*_COSMETIC_CHANGED` is
    comments/whitespace/case only.

Practical rules:

- Diffing large dbs prints nothing until done and can take minutes at full
  CPU — run it in the background; a silent diff is not a hung diff.
- Diff is for ADJACENT versions of one lineage. Two diverged workbooks
  produce a flood of "changes" — noise; compare with targeted SQL instead.
- Bucket `changes` by `type`/`sheet` in Python and classify EVERY workbook-
  layer change into an expected bucket — anything outside the list is a
  finding. This is the regression tool after any builder change or version
  bump; don't re-prove from scratch.

## Recalculating with `marinade calc` (experimental)

Needs `pip install 'xl-marinade[calc]'` (exit 5 if missing); outside SemVer, so
formats may change. `marinade values` returns stored values and needs no extra.

```bash
marinade calc book.xlsx --policies policies.csv --map map.json \
    --outputs 'Summary!B4' --variables vars.json \
    [--check-against check.json] [--as-of 2026-01-01] -o out/
marinade values book.xlsx --variables vars.json [--map map.json] -o stored/
```

- **Files in.** `policies.csv` (id column + one column per field), `map.json`
  (`{"id_field": "policy", "fields": {"age": "Inputs!C3"}}`; targets must be
  constants), `vars.json` (`{"prem": {"range": "Proj!C11:C130", "period_start": 0}}`
  or `{"npv": {"cell": "Summary!B4"}}`). **`period_start` is required on every
  range, no default; periods count from it.** `check.json` names stored results
  for other input sets (results sheet + id column + output-to-column map, and an
  input `table` or a `selector` cell) so the trust run covers more paths.
  Targets are cells, ranges or defined names, in any case; a sheet's own
  (sheet-scoped) name is written `'Sheet'!Name`, and a bare name that a sheet
  also defines is refused as ambiguous. A whole column (`Calc!B:B`) covers the
  rows the sheet uses; a target covering no cells (a formula or `OFFSET` name, a
  constant) is an error. Every cell of an array formula's range (legacy or
  dynamic/spill) is a formula cell: compared in the trust run, carrying the
  formula's status, and an error as a map or selector target.
- **Files out** in `-o`: `results.csv`, `rows.csv`, `paths.csv`,
  `trust_report.json`, `run.json`. Known outputs are cleared at the start of a run.
- **Read `trust_report.json` first** (`share` reproduced, `failures`,
  `refused_patterns`, `settings`), then filter on `status`:
  `refused` (depends on a failing/refused root cell, or the engine returned an
  error code Excel does not have — do not use), `uncertain` (rounding flip or
  near-but-not-equal comparison; an unplaced rounding/comparison probe),
  `unverified_path` (a branch or lookup row the trust run never took; an unplaced
  branch/lookup probe), `ok`. Only `ok` can be used as is; the `reasons` column
  names the root cell. `values` rows are `stored`.
- Only reproduced cells vouch for their formula family's branches; branches inside
  an arm not evaluated count for nothing. Lookups (`XLOOKUP`/`XMATCH` included)
  must have hit the first row, last row and not-found in the trust run to be `ok`
  there. Probes go upstream of the requested outputs only, so fewer outputs can
  mean a more conservative status.
- `OFFSET`/`INDIRECT` (and ranges ending in a call, `C5:OFFSET(...)`) are followed per
  policy; their argument cells are precedents; every formula cell on the sheet an
  `OFFSET` starts from or an `INDIRECT` text names literally is probed. An `INDIRECT`
  naming no sheet that lands on unprobed formulas gives `not_checked_path` (and
  `not_checked_value` for their roundings/comparisons) — never `ok`.
- **Refused even when they reproduce:** approximate lookups on unsorted, mixed,
  whole-column or non-plain-text keys (and `XLOOKUP`/`XMATCH` binary search);
  implicit intersection on a multi-cell name or on a range outside a function
  argument; names holding a formula or constant; `RAND`/`RANDBETWEEN`/`RANDARRAY`;
  `TODAY`/`NOW` without `--as-of`; circular refs; data tables; unparseable
  formulas (an error literal such as `#N/A` typed into a formula is a known
  cause); external references that cannot be held (ranges, no cached value,
  names pointing into another workbook); approximate lookups on input-dependent
  keys; a call to a function IronCalc lacks (`AGGREGATE`, `BAHTTEXT`, the `CUBE…`
  functions and a few more; the detail names it). Per policy: `lookup value type
  differs from keys` (text against numbers); an `INDIRECT` address off the grid
  (row 0) is `dynamic reference not checked`; number-to-text where the engine's
  text can differ from Excel's (`number-to-text at a decimal tie` for
  `TEXT`/`FIXED`/`DOLLAR`; `… of a negative value that rounds to zero or one unit`
  for `TEXT`; `… where Excel's 15-digit text differs` for `&`, `CONCAT`/`TEXTJOIN`
  (ranges read cell by cell) and text functions; `criterion built with & from a
  non-integer number` for `SUMIF`-family criteria). Static, for every policy:
  `number-to-text not checked` (unreadable or non-literal `TEXT` format, a range
  given to a text function, array formulas, `LET`, `--probe-cap`).
- **Exit codes:** 0 ok, 1 error, 5 `[calc]` missing, 6 workbook refused outright
  (1904 dates, precision as displayed, Strict Open XML).
- A pass on the stored inputs proves only the paths those inputs take; never
  call a value verified because the share is 100%.
- **Give `--check-against` whenever the workbook stores results for other input
  sets** (a batch-results sheet): with results stored for one input set only,
  lookups landing mid-table and branches the stored inputs skip stay unverified.
  On lifelib BasicTerm, no check set gave 100% `unverified_path`; 200 checked
  points gave 100% `ok`, all correct.
- A check-set mismatch (`check_against_mismatch`) refuses an output only for
  policies on the same path (same branch and lookup-edge outcomes upstream; an
  unplaced or unsure probe counts as the same; a probe-free output has one path).
  `trust_report.json` `check_against.mismatches` gives `stored`, `computed`,
  `relative_difference`.
- **`ok` means every formula and path was checked against Excel's stored values,
  not that a new value is within tolerance of what Excel would give**:
  cancellation (a net figure as a small difference of large present values) can
  amplify floating-point differences past 1e-9 (one lifelib point: 1.003e-9).
- `#NAME?` downstream of a VBA UDF is `macro_function` rooted at the UDF cell;
  `unsupported_function` only when no UDF is involved.
- Results stream to disk per policy: memory stays flat in policies and check
  sets, but `paths.csv` can run to ~1,000+ rows per policy. A run that stops
  part-way deletes `results.csv`/`paths.csv`/`rows.csv`.

## Verification core rules (templates in references/authoring.md)

- Two proofs, in this order: **structure proof** (pure SQLite, no Excel —
  the fast gate), then **value proof** (Excel recalc — the slow gate).
- Structure-proof FAIL triage: check whether the "missing" reference exists
  verbatim in the stored `agent_cells_light.formula` — if yes it's an
  extractor gap (record it in a known-gaps list), not a workbook error.
  Don't ask the edge tables to verify themselves.
- In-sheet identity checks: per-row RELATIVE residuals with divisor = the
  LARGEST side (`MAX(1, |lhs|, |rhs|)`); count violations, never sum absolute
  tolerances. Diagnose a failing check by DUMPING THE VIOLATING ROWS first —
  and consider that a faithful twin reproduces the source model's own
  defects: the source's consistency check may itself be the bug.
- Matrix solves: measure `cond()` before choosing tolerances — Excel
  MINVERSE vs numpy agree only to eps·cond.

## What passing proofs still miss (adversarial-audit checklist)

Value and structure proofs certify only what they read. Audit these classes
separately:

1. **Evidence chain** — commit all artifacts from ONE final pass in
   dependency order (build → VBA → macro run saved in-file → proofs →
   re-extract IR → diff); hash-tie IR dbs and reports to committed files.
2. **Unverified visible regions** — Excel `OR()`/`AND()` do NOT
   short-circuit: an out-of-range INDEX inside a guard poisons it with
   #REF!. Nest IFs (only IF is lazy). Everything a reviewer can SEE must be
   verified or guarded, including beyond-horizon rows.
3. **Right value, wrong mechanism** — replicate the source's resolution
   MECHANISM (dict get vs match-cascade), not just the resolved value on
   current data.
4. **Provenance** — verify labels/cover claims like numbers.
5. **Reconciliation scope** — state the exact quantities reconciled, then
   reconcile ALL of them (not a silent subset).

Minors worth a line: Excel ROUND (half-away-from-zero) ≠ Python round()
(banker's); chart series over ""-returning IFs plot as ZEROS (chart the exact
observed range or return NA()); dead input columns mislead auditors — wire
them or drop them.

## Operational rules (read references/excel-ops.md before ANY Excel work)

- Marinade never needs Excel; value proofs and recalcs do. On macOS, Excel is
  ONE shared instance machine-wide: coordinate ALL Excel work (hold a
  machine-wide lock) and NEVER `pkill` Excel. An Excel "hang" is another
  process's workbook until proven otherwise — ask Excel for its ACTIVE
  WORKBOOK NAME before diagnosing anything.
- After any recalc-and-save, probe one known computed cell for a real number
  before trusting a tie-out — `None == None` "passes", and a suspiciously
  fast recalc means calculation never ran.
- openpyxl hazards: close the file in Excel before headless edits; saves
  strip cached values (next open = full cold recalc); ArrayFormula/spill
  cells must be translated with a corrected `ref` — verbatim copies corrupt
  the file (Excel repair dialog).
- Long recalcs/macros exceed foreground timeouts: run them in the background;
  launch macros via `osascript` inside `with timeout of 3600 seconds` —
  xlwings `wb.macro()` has a fixed ~60s Apple-event timeout.
