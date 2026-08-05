---
name: xl-marinade
description: Use XL Marinade (formula-graph IR extraction, binding-level dependency queries, structural diff) when authoring, auditing, modifying, or verifying Excel workbooks — especially actuarial/financial models. Triggers - building a workbook that must be provably correct, auditing or changing an existing (human-built) workbook, tracing what drives a cell/output, attributing a numeric change, comparing two workbook versions, checking a workbook's calculation structure against a source model (Python/lifelib), or working with VBA-bearing xlsm files.
---

# XL Marinade — agent guide

Marinade extracts a workbook's calculation structure into a SQLite "IR"
database: cells, formulas, binding-level groups, a dependency graph, and
parsed VBA. It is an **auditor/tracer, not a calculator** — it snapshots
cached values and never recalculates.

Pick the reference for your task BEFORE starting:

- **Authoring a new provably-correct workbook** → `references/authoring.md`
  (author-for-the-extractor conventions, value/structure proof patterns, VBA).
- **Auditing or modifying an existing workbook** → `references/auditing.md`
  (recon SQL, version-chain diffs, mechanism-vs-data separation, attribution).
- **Anything that opens Excel or edits a file via openpyxl** →
  `references/excel-ops.md` (mandatory: lock discipline, recalc traps,
  invisible modals, corruption hazards).

## Extract

```bash
pip install xl-marinade   # CLI: marinade
marinade extract path/to/book.xlsx -o out.ir.db
```

- Works on .xlsx and .xlsm. VBA is parsed into the IR when oletools is
  present (`pip install "xl-marinade[vba]"`); without it extraction still
  succeeds and prints a notice.
- **Recalculate in Excel and save before extracting** if you will read
  cached values (openpyxl saves STRIP cached values); formula/edge
  structure does not need recalc.
- Runtime scales with the book: a generated ~65k-cell workbook extracts in
  seconds; a large human-built model (thousands of bindings) takes
  minutes — run it in the background, keep one `.ir.db` per version in
  scratch, and don't commit ad-hoc dbs.

## Query

**Direction convention:** `from_*` = the formula (the dependent); `to_*` =
what it reads (the precedent). "What drives X" = rows WHERE from_address
is in X.

On workbooks **authored to the conventions** (machine-key row etc.) the
`agent_*` views are the entry point — 65k cells collapse to ~100–250
labeled bindings. On human-built books labels degrade and bindings number
in the thousands: use `agent_bindings` for orientation only and do audit
work at cell level (`cells`/`formulas`/`cell_edges_internal`/`range_edges`).

```sql
-- model map: one row per detected column/block with auto-derived label
SELECT sheet, address, formula_pattern, label FROM agent_bindings;

-- what drives a given output (precedents; DISTINCT — the view has one
-- row per contributing edge)
SELECT DISTINCT to_label, to_address FROM agent_binding_dependencies
WHERE from_address LIKE 'Summary!%';

-- ground truth when an edge looks wrong: the stored formula
SELECT s.sheet_name || '!' || c.a1, c.formula_a1
FROM cells c JOIN sheets s ON c.sheet_id = s.sheet_id
WHERE s.sheet_name = 'Projection' AND c.a1 = 'D10';

-- cross-sheet consumer census ("is this tab dead? what breaks if I
-- replace it?") — a reference arrives as a cell edge OR a range edge,
-- so UNION both:
SELECT s_from.sheet_name, COUNT(*) FROM (
  SELECT e.from_cell_id, c_to.sheet_id AS to_sheet
    FROM cell_edges_internal e JOIN cells c_to ON c_to.cell_id = e.to_cell_id
  UNION ALL
  SELECT r.from_cell_id, r.to_sheet_id FROM range_edges r
) u JOIN cells c_from ON c_from.cell_id = u.from_cell_id
JOIN sheets s_from ON s_from.sheet_id = c_from.sheet_id
JOIN sheets s_to   ON s_to.sheet_id   = u.to_sheet
WHERE s_to.sheet_name = '<sheet>' AND s_from.sheet_name != '<sheet>'  -- fill BOTH
GROUP BY 1 ORDER BY 2 DESC;

-- find every consumer of a defined name / anchor before assuming a wide
-- edit (the defined_names table maps each name to its destination ref)
SELECT s.sheet_name, c.a1 FROM cells c
JOIN formulas f ON f.formula_id = c.formula_id
JOIN sheets s ON s.sheet_id = c.sheet_id
WHERE f.formula_r1c1 LIKE '%my_defined_name%';

-- VBA is structure too
SELECT m.name, p.name, p.kind FROM vba_procedures p
JOIN vba_modules m ON p.module_id = m.module_id;
```

Schema gotchas — when a query errors, run `.schema <table>`; never guess
twice:

- `sheets.sheet_name` (not name); `cells.a1` (not address); `cells.formula_a1`
  holds the per-cell A1 formula.
- Formula-text search lives in `formulas` (`formula_r1c1`,
  `formula_a1_example`) joined via `cells.formula_id` — there is NO
  `formula_text` column.
- `bindings` has no sheet/address columns — it's `sheet_id` + `address_a1`;
  the `agent_*` views flatten these for you.
- Single-cell refs land in `cell_edges_internal`, multi-cell in
  `range_edges` (from_cell_id → to_sheet_id + to_r1/c1/r2/c2).
- `cell_comments` carries its own `sheet_id` + `a1` — join to `sheets`
  directly, not through `cells`.

## Diff two workbook versions

```bash
marinade diff a.ir.db b.ir.db -o diff.json
# JSON top keys: "summary" (counts per change type), "binding_map"
# (per-binding diff_state unchanged/modified/added/removed/moved, with
# labels), "changes" (full changelist: [{"type": "FORMULA_CHANGED"|…}]).
```

- Read `summary` + `binding_map` first — that pair is the agent-consumable
  changelist; open `changes` only to drill in. Formula changes are
  classified reference_shift vs logic_change.
- Output is split into **workbook edits** vs **IR INFERENCE
  (extractor-derived; not workbook edits)** — table candidates, label
  evidence, time annotations. Inference events churn with IR segmentation
  even when no cell changed: never count them as edits in an audit. In
  full JSON they carry `"layer": "ir_inference"`.
- Diffing large dbs (tens of MB) takes tens of minutes at full CPU and
  prints NOTHING until done — run it in the background; a silent diff is
  not a hung diff.
- Diff is for ADJACENT versions of one lineage. Two diverged workbooks
  produce hundreds of thousands of "changes" — noise; compare with
  targeted SQL instead.
- Bucket full-JSON output in Python
  (`Counter(c['sheet'] for c in d['changes'] if c['type']=='FORMULA_CHANGED')`)
  and classify EVERY change into expected buckets — anything outside the
  list is a finding. This is the regression tool after any builder change
  or version bump; don't re-prove from scratch.

## Verification core rules (templates in references/authoring.md)

- Two proofs, in this order: **structure proof** (pure SQLite, no Excel —
  the fast gate), then **value proof** (Excel recalc — the slow gate).
- Structure-proof FAIL triage: check whether the "missing" reference
  exists verbatim in the stored `formula_a1` — if yes it's an extractor
  gap (record in a KNOWN_EXTRACTOR_GAPS dict), not a workbook error. Don't ask
  the edge tables to verify themselves.
- In-sheet identity checks: per-row RELATIVE residuals with divisor = the
  LARGEST side (`MAX(1,|lhs|,|rhs|)`); count violations, never sum
  absolute tolerances. Diagnose a failing check by DUMPING THE VIOLATING
  ROWS first — and consider that a faithful twin reproduces the source
  model's own defects: the source's consistency check may itself be the
  bug.
- Matrix solves: measure cond() before choosing tolerances — Excel
  MINVERSE vs numpy agree only to eps·cond.

## What passing proofs still miss (adversarial-audit checklist)

Value and structure proofs certify only what they read. Audit these
classes separately:

1. **Evidence chain** — commit all artifacts from ONE final pass in
   dependency order (build → VBA → macro run saved in-file → proofs →
   re-extract IR → diff); hash-tie IR dbs and reports to committed files.
2. **Unverified visible regions** — Excel `OR()`/`AND()` do NOT
   short-circuit: an out-of-range INDEX inside a guard poisons it with
   #REF!. Nest IFs (only IF is lazy). Everything a reviewer can SEE must
   be verified or guarded, including beyond-horizon rows.
3. **Right value, wrong mechanism** — replicate the source's resolution
   MECHANISM (dict get vs match-cascade), not just the resolved value on
   current data.
4. **Provenance** — verify labels/cover claims like numbers.
5. **Reconciliation scope** — state the exact quantities reconciled, then
   reconcile ALL of them (not a silent subset).

Minors worth a line: Excel ROUND (half-away-from-zero) ≠ Python round()
(banker's); chart series over ""-returning IFs plot as ZEROS — chart the
exact observed range or return NA(); dead input columns mislead auditors —
wire them or drop them.

## Operational rules (read references/excel-ops.md before ANY Excel work)

- Marinade never needs Excel; value proofs and recalcs do. On macOS, Excel is
  ONE shared instance machine-wide: hold a machine-wide lock for ALL Excel
  work (this repo provides `verification/excel_lock.py`) and NEVER pkill
  Excel. An Excel "hang" is another process's workbook until proven
  otherwise — ask Excel for its ACTIVE WORKBOOK NAME before diagnosing
  anything.
- After any recalc-and-save, probe one known computed cell for a real
  number before trusting a tie-out — None==None "passes", and a
  suspiciously fast recalc means calculation never ran.
- openpyxl hazards: close the file in Excel before headless edits; saves
  strip cached values (next open = full cold recalc); ArrayFormula/spill
  cells must be translated with a corrected `ref` — verbatim copies
  corrupt the file (Excel repair dialog).
- Long recalcs/macros exceed foreground timeouts: run them in the
  background; launch macros via `osascript` inside
  `with timeout of 3600 seconds` — xlwings `wb.macro()` has a fixed ~60s
  Apple-event timeout.
