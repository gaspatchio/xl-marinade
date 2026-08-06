---
name: using-xl-marinade
description: Use XL Marinade (deterministic Excel formula-graph extraction to SQLite, binding-level dependency queries, structural diff) when authoring, auditing, modifying, or verifying Excel workbooks — especially actuarial/financial models. Triggers - building a workbook that must be provably correct, auditing or changing an existing (human-built) workbook, tracing what drives a cell/output, attributing a numeric change, comparing two workbook versions, checking a workbook's calculation structure against a source model (Python/lifelib), or working with VBA-bearing xlsm files.
---

# Using XL Marinade — agent guide

XL Marinade extracts a workbook's calculation structure into a SQLite "IR"
database: cells, formulas, binding-level groups, a dependency graph, and
parsed VBA. It is an **auditor/tracer, not a calculator** — it snapshots
cached values and never recalculates.

Install and run it with `uv` (or any venv where `xl-marinade` is installed):

```bash
uv run marinade --help          # extract · document · diff
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

- Works on `.xlsx` and `.xlsm` (VBA is parsed into the IR).
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
  `formula` (A1), `formula_r1c1`, `data_type`, `is_array_formula`.
- `agent_bindings` — `sheet`, `address` (`Sheet!A1:A10`), `formula_pattern`
  (R1C1; NULL for constant/heterogeneous blocks), `label`, `binding_type`,
  `classification`, `confidence`.
- `agent_dependencies` — `from_cell`, `to_cell`, `dependency_type`
  (`cell` | `range` | `external`), `cell_count` (ranges), `to_r1/c1/r2/c2`.
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
  `label_a/b`. This labeled map is the agent-consumable changelist.
- **`changes`** — the per-event detail. Each carries `seq`, `type`, and a
  **`layer`**:
  - `layer = "workbook"` — real workbook edits (`FORMULA_CHANGED`,
    `BINDING_FORMULA_CHANGED`, `BINDING_ADDED`, `COLS_INSERTED`,
    `*_EDGE_ADDED`). `BINDING_FORMULA_CHANGED` carries `modification_kind`
    (`logic_change` vs a reference shift) — the logic changes are your
    review set.
  - `layer = "ir_inference"` — extractor-derived, NOT workbook edits
    (`TABLE_CANDIDATE_*`, `BINDING_LABEL_EVIDENCE_CHANGED`). These churn as
    IR segmentation shifts even when no cell changed; `summary` counts them
    as `ir_inference_changes`. **Never count them as edits in an audit.**

Practical rules:

- Diffing large dbs prints nothing until done and can take minutes at full
  CPU — run it in the background; a silent diff is not a hung diff.
- Diff is for ADJACENT versions of one lineage. Two diverged workbooks
  produce a flood of "changes" — noise; compare with targeted SQL instead.
- Bucket `changes` by `type`/`sheet` in Python and classify EVERY workbook-
  layer change into an expected bucket — anything outside the list is a
  finding. This is the regression tool after any builder change or version
  bump; don't re-prove from scratch.

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
