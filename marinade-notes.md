# XL Marinade — usage notes from building lifelib workbook twins

Running log of how Marinade behaves when used to support agent-authored actuarial
workbooks. Feeds the future `.claude/skills/xl-marinade` skill. Findings
tagged **A#** are Marinade issues/limitations; **P#** are usage patterns that work.

## Session 2026-07-13 — BasicTerm_S

### Invocation

```bash
# extract (pip install xl-marinade; needs ≥ 0.1.0 — earlier internal
# builds lacked the lookup-transparency and name-token fixes)
marinade extract workbooks/BasicTerm_S/BasicTerm_S.xlsx \
    -o marinade/BasicTerm_S.ir.db
```

- Extraction of a 7-sheet, ~65k-cell workbook (10,000-row model point table)
  takes a few seconds. Output is a single SQLite file; query with `sqlite3`.
- **Recalc before extract**: Marinade snapshots *cached* values, it never
  recalculates. Our harness saves the workbook from Excel (xlwings) right
  after a full recalc so cached values are trustworthy.

### P1 — The `agent_*` views are the right entry point

The live schema is newer than `docs/PROJECT_VISION.md` suggests. There is no
`consistency_report`/`cycles`/`levels` table in current output; instead there
is a family of agent-oriented views. The two that matter most for authoring:

- `agent_bindings` — one row per detected column/block: sheet, A1 range,
  R1C1 `formula_pattern`, auto-derived `label`, label-candidate evidence.
- `agent_binding_dependencies` — labeled binding-level dependency graph
  (`from_label`, `from_address` → `to_label`, `to_address`, `edge_count`).

For BasicTerm_S: 88 bindings, 90 binding edges, 101 labeled dependency rows —
a compact, readable model map (vs ~65k raw cells). This is the granularity an
authoring agent should reason at. (Counts are from the first extraction that
session; the committed IR db reflects the final workbook and differs slightly.)

### P2 — Author for the grouper and it labels your model for free

Marinade groups a formula column into ONE binding when every row shares the same
R1C1 pattern. Design consequences that all paid off:

- Keep each projection column's formula uniform across all rows (use IF
  guards rather than different formulas per region). Marinade's "init-merger"
  tolerates a different t=0 row (it merged our `pols_if` / `pols_maturity`
  seed rows into the column bindings).
- Put a machine-key row (the lifelib cell name, small grey italic text)
  directly under the human header. Marinade's label scan picked these up
  verbatim: bindings came out labeled `t`, `duration`, `age`, `pols_if`, …
  which makes the IR graph directly comparable to the Python model.
- Named ranges (`Proj_NetCF`, `PV_Claims`, …) surface in `defined_names`
  and as `named_exact` label candidates — name anything you'll verify.

### P3 — Structure proof pattern

Value agreement doesn't prove the calculation *structure* matches. Pattern:
transcribe the source model's call graph as (dependent, precedent) pairs,
map each model cell to a representative workbook cell, and assert an IR edge
exists whose from/to binding ranges contain those cells
(`verification/structure_proof_basicterm_s.py`). Found real Marinade gaps on
first use:

> **Update (same day):** A1 and A2 are FIXED in the extractor (included
> in xl-marinade ≥ 0.1.0): INDEX/XLOOKUP/CHOOSE are now transparent for
> argument-ref extraction (still emit their DYNAMIC marker). Structure proof
> went 57/66 -> 66/66. Notes kept for the record.

### A1 — Cell refs in MATCH/INDEX arguments produce no edges (FIXED — PR #1)

`Projection!D10 = INDEX(MortRates, MATCH(C10, MortAges, 0), MIN(B10, 5) + 1)`
yields range edges to `Mortality!B2:G104` and `Mortality!A2:A104` but **no
edges to C10 or B10**. Plain refs outside lookup functions (e.g.
`E10 = 1-(1-D10)^(1/12)`) edge correctly. Impact: precedent tracing through
lookup-driven cells silently loses the lookup *keys* — for an actuarial model
that means losing the age/duration drivers of every table lookup.

### A2 — INDEX/MATCH over defined names into an all-constant sheet: no edges at all (FIXED — PR #1)

`Assumptions!C5 = INDEX(MP_Age, MATCH(PointID, MP_ID, 0))` produced **zero**
outgoing edges (no cell edges, no range edges), and there are zero edges of
any kind into the 10,001-row constant `Model_Points` sheet, even though the
`MP_*` defined names are correctly registered in `defined_names`. The same
formula shape against the `Mortality` sheet (A1) does produce range edges, so
the difference is plausibly the constant-binding filtering of the big sheet
("Filtered 61219 cells already in constant bindings" in extractor output).
Impact: the entire model-point data sheet is invisible in the dependency
graph.

### Handy queries

```sql
-- model map: every labeled dependency
SELECT from_label, from_address, to_label, to_address
FROM agent_binding_dependencies;

-- what drives a given output (precedents of a labeled binding)
SELECT to_label, to_address FROM agent_binding_dependencies
WHERE from_address LIKE 'Summary!%';

-- stored formula of a cell (ground truth when edges look wrong)
SELECT s.sheet_name || '!' || c.a1, c.formula_a1
FROM cells c JOIN sheets s ON c.sheet_id = s.sheet_id
WHERE s.sheet_name = 'Projection' AND c.a1 = 'D10';
```

Gotchas: `sheets.sheet_name` (not `name`); `cells.a1` (not `address`);
bindings reference sheets via `sheet_id` — the `agent_*` views join these
for you, prefer them.

### P4 — Shipping VBA in generated workbooks (macOS reality)

Requirement: every workbook carries a RunAllModelPoints macro. Findings:

- openpyxl cannot author VBA. An `.xlsm` is just a zip: inject
  `xl/vbaProject.bin`, switch the workbook content type to
  `macroEnabled.main+xml`, add the vbaProject relationship — see
  `make_xlsm()` in `workbooks/BasicTerm_S/build_workbook.py`. Deterministic
  and CI-friendly once the bin exists.
- Producing `vbaProject.bin` from a `.bas` is the hard part on macOS: no
  `do Visual Basic` AppleScript event, no VBE in Excel's sdef, no
  `vb_project` via xlwings/appscript, no PyPI compiler. Options: UI-script
  the VBE import (needs Accessibility permission), or a one-time manual
  import (⌥F11 ▸ File ▸ Import File… ▸ save as xlsm) then
  `bootstrap_vba.py --harvest` to extract the bin for the repo.
- Write the macro against **named ranges** (`Names("PointID").RefersToRange`),
  never sheet codenames — survives rebuilds, and keeps the macro source
  readable next to the lifelib mapping.
- Marinade parses VBA (`vba_modules` / `vba_procedures` / `vba_procedure_edges`
  tables via oletools) — the injected module is visible to the IR, which we
  can use to structure-check the macro itself. Not yet exercised.

### P5 — VBA batch loops on Mac Excel: guard against stale calculation

First full 10,000-point RunAllModelPoints run FAILED with 505 mismatches in
a striking pattern: every ~11th point from #4388 onward, and 98/100 sampled
failures held exactly the PREVIOUS point's results. Diagnosis:
`Application.Calculate` under `xlCalculationManual` can return before the
dependency chain finishes when VBA loops thousands of set-value/calculate
cycles — the macro then reads last iteration's values. The external
xlwings-driven loop never hit this (each Apple-event round trip is
synchronous), which is why the Python mirror passed 300/300 while the VBA
failed. Fix (now standard for all workbook macros):

1. After `Application.Calculate`, spin `Do While
   Application.CalculationState <> xlDone : DoEvents : Loop` (probe the
   property once — availability differs across Excel builds).
2. Belt-and-braces sentinel: a formula cell that echoes the loop input
   (Summary!C4 `=PointID`) must equal the point just written; retry
   Calculate until it does, hard-error after N retries. A stale chain shows
   the previous id, so this catches exactly the failure mode.

**Outcome:** with the guards in place the full run PASSES — 10,000/10,000
reconciled, premium diff exactly 0, max PV rel diff 5.2e-11, 119s. Two more
operational lessons: (1) never `MsgBox` in an unattended macro — an error
dialog in a hidden Excel is an invisible modal that blocks every Apple event
(next .bas revision should write errors to Batch_Results instead); (2) only
ONE agent may drive Mac Excel at a time — it is a single shared instance, and
a concurrent session force-killing Excel severs connections mid-run
(-609/-1712) and corrupts add-in state. Coordinate via a lockfile, never
pkill Excel. **Implemented:** `verification/excel_lock.py` (machine-wide
lock at /tmp/excel-agent.lock — blocking queue, dead-PID self-heal; drop
the same helper into any other workspace that drives Excel, with the
rules in each repo's CLAUDE.md).

Marinade VBA visibility confirmed: the injected module appears in the IR as
`vba_modules` (name `BatchRunner.bas`, kind standard, full source_text +
sha256) with `vba_procedures` (RunAllModelPoints, sub) and 10 static cell
refs in `vba_procedure_cell_refs` — the macro is auditable structure, not a
black box.

Related invisible modal (2026-07-14): Excel's own "Enable Macros" security
prompt on every .xlsm open. Accessibility/Automation grants don't remove it —
it's an Excel preference, not a TCC permission, and with `visible=False` it
stalls runs silently. Fix at source (Excel quit first):
`defaults write com.microsoft.Excel VisualBasicMacroExecutionState -string
"EnabledWithoutWarnings"`. Trade-off: ALL workbooks' macros run unprompted.

Also learned: zip-surgery .xlsm assembly requires codeName plumbing —
`<workbookPr codeName="ThisWorkbook"/>` (openpyxl `wb.code_name`) and per-
sheet `sheetPr codeName` matching the vbaProject.bin doc modules, else the
first `ThisWorkbook` reference dies with run-time error 429. And VBA writing
"TRUE"/"FALSE" strings via `Range.Value` lands as native booleans.

## Session 2026-07-13 (later) — BasicTerm_SE

### P6 — ir_diff across related models is credible at binding granularity

First real IR-diff exercise: `marinade diff marinade/BasicTerm_S.ir.db
marinade/BasicTerm_SE.ir.db` (summarize the JSON changelist). The two
workbooks share layout conventions but SE adds in-force mechanics. Result
reads like a sensible model changelist, not noise:

- BINDING MAP: 105 unchanged / 58 modified / **17 added** / 2 removed /
  1 moved — the added bindings are exactly SE's new engine columns
  (duration_mth, is_active, the BEF_NB/BEF_DECR timing columns,
  new business) plus the Premium_Rates table.
- sheets_added 1 (Premium_Rates), cols_inserted 2, names_added 11
  (PolicyCount, Duration0, Prem* …), bindings_formula_changed 28 (the S
  formulas that gained is_active guards / switched exposure basis).
- Full JSON output is large (edges_changed ~28k at cell level) — for agent
  consumption stay at the binding-level sections (`summary` +
  `binding_map` in the JSON) first.

Practical: this is the regression tool for workbook iterations — after any
builder change, extract + diff old vs new and read the changelist instead
of re-proving everything from scratch.

### SE-specific authoring notes

- 2D rate-table lookups: keep header cells NUMERIC with a display format
  (`'"Term "0'`) so MATCH works — a text header ("Term 10") makes every
  dependent formula #N/A. Cost me one build cycle.
- Cells lifelib can't evaluate (mortality lookups at negative durations)
  need is_active guards in Excel AND null-markers in the reference JSON;
  the value proof compares guarded cells only where active.
- lifelib's `pols_if_at(t, timing)` maps cleanly to one column per timing —
  the within-month decrement order becomes visually explicit
  (BEF_MAT → maturities → BEF_NB → new business → BEF_DECR → deaths/lapses),
  which an actuary can audit line by line. Structure proof: 88/88 edges.
- The BatchRunner macro is workbook-agnostic (named ranges only): the SAME
  .bas served S and SE; each workbook still needs its own vbaProject.bin
  (doc modules bind per sheet count), one automated bootstrap each.

## Session 2026-07-13/14 — CashValue_SE

### P7 — A faithful twin reproduces the source model's own defects

CashValue_SE value proof initially FAILED only on its Checks sheet for the
single-premium products — while all 153,042 numeric comparisons matched
lifelib to 1e-9. Root cause, established by dumping the violating rows:
the per-month AV roll-forward identity **does not hold across the maturity
boundary in lifelib itself** (the maturity release is booked one row after
BEF_NB empties; the identity only telescopes over the two rows), which is
why lifelib's own `check_av_roll_fwd` returns False for those products.
The workbook was so faithful it replicated the broken check. Resolution:
the in-sheet check excludes maturity-boundary rows with the reason
documented on the sheet — and this is worth an upstream note to lifelib.

Two generalizable check-authoring rules came out of the same debugging:
- Per-row identity residuals must be RELATIVE, and the divisor must be
  the LARGEST side of the identity (`MAX(1, |lhs|, |rhs|)`), else rows
  where one side is exactly zero measure float residue against 1.
- Diagnose check failures by dumping the violating rows before theorizing
  — the first two theories (isclose-on-zero, summed tolerance) were both
  wrong; the row dump settled it in one look.

Operational: on macOS, long xlwings runs die if the machine sleeps
mid-run (wedged windowless Excel, hung Apple events). Chunk long value
proofs into per-profile invocations (verify script takes profile args)
and re-probe Excel scriptability before each chunk.

### Open questions for next session

- Does `--skip-evidence` speed up extraction materially on bigger books?
- VBA parsing path — exercised lightly (BatchRunner visible in IR); deeper
  use in the solvency2 capstone (scenario orchestration in VBA).
- Draft `.claude/skills/xl-marinade` from P1–P6 (enough material now).

## Session 2026-07-14 — SmithWilson

### P8 — Array formulas: what Marinade sees, and how Excel sabotages them

SmithWilson is the matrix-math test case (Wilson grid, ζ = W⁻¹(m−μ) via
MINVERSE/MMULT, curve extrapolation). Findings:

- **2D grids collapse cleanly.** The 65×25 Wilson grid (one uniform R1C1
  formula) became ONE binding, auto-labeled `W` from the machine-key row —
  the key-row convention works for blocks, not just columns. The whole
  11-cell model reads out of `agent_binding_dependencies` exactly like the
  lifelib call graph (structure proof 20/20, incl. `zeta ← W` from inside
  the CSE array and `P ← W` through the INDEX row-slice, PR #1 at work).
- **CSE array formulas parse fine.** A legacy `{=MMULT(MINVERSE(WMatrix),
  DVec)}` spanning 25 rows appears as a single-cell binding at its top-left
  anchor; edges are extracted from inside it. Map vector cells to that
  anchor in structure proofs.
- **THE TRAP — implicit intersection.** Plain (non-CSE) formulas written
  by openpyxl get legacy semantics: on open, Excel 365 silently inserts
  `@` before range names used as array arguments — `TRANSPOSE(@ZetaVec)` —
  collapsing the vector to a scalar and yielding #VALUE!. Not visible in
  openpyxl; only FORMULATEXT shows the `@`. RULE: any generated formula
  that feeds a whole range into array machinery (TRANSPOSE, MMULT,
  MINVERSE, or INDEX-slice arithmetic) must be written as an openpyxl
  `ArrayFormula` (CSE), even single-cell. Plain SUMPRODUCT over same-shape
  ranges is safe; anything needing reorientation is not.
- **Probe errors via a scratch cell** (`=FORMULATEXT(cell)` +
  `=ERROR.TYPE(cell)` written with xlwings, never saved): fastest way to
  see what Excel actually parsed vs what openpyxl wrote.

VBA reconciliation adapted to a model with no model points: the macro
reconciles the full result vectors (ζ, P, R = 155 values) instead of
looping a batch — the "reconcile everything against Lifelib_Reference"
convention carries over unchanged. Numerics: Excel MINVERSE vs numpy inv
agree to eps·cond(W) ≈ 1.7e-9 absolute on ζ (cond(W) ≈ 7e5) and 1e-12 on
P/R; measure cond BEFORE choosing check tolerances on matrix solves.

Bootstrap hardening (all four copies): AppleScript `quit` is
asynchronous — poll pgrep until Excel exits before relaunching, or the
dying instance tears down the new one; the VBE import open panel is a
window named "Import File" (not a sheet); an empty Excel reports workbook
list `missing value`, which must not be parsed as a foreign workbook.

## Session 2026-07-14 (later) — Solvency2 capstone

### P9 — The capstone: 22 sheets, inheritance via parameterized scenario sheets

solvency2 is the pattern-inheritance test: lifelib derives six stressed
projections from one engine via Override spaces. The Excel translation
that worked: ONE engine layout instantiated as seven scenario sheets, each
with a parameter block (risk, shock, sign, factors looked up from the
FactorData input table); the override formulas branch on those parameters
exactly where lifelib's Override cells replace base cells. Uniform
formulas per column survive (551 bindings for a 22-sheet workbook), and
the machine-key row makes all seven engines read as lifelib cells.

Other patterns that carried the build:
- **modelx match() → 8-step IFERROR/MATCH cascade** on concatenated
  lookup keys, with None-valued rows OMITTED from the workbook tables
  (match() skips them; presence-with-None ≡ absence). Verified
  byte-identical fallback order against modelx CellsImpl.find_match.
- **Shared schedule sheet** for pol/asmp/scen cells (reserves via
  commutation, cash values, base rates, discount rates) — lifelib holds
  these once per policy; so does the workbook. The stress sheets read them
  cross-sheet, which keeps the 7 engines thin.
- **Commutation identity**: ReserveNLP_Rate(t) collapses to one uniform
  formula because Axn(x+t, n−t) and the gamma2 annuity have constant
  numerators (Mx/Nx at x+n, x+m) — only Dx(x+t) and Nx(x+t) move with t.
- Structure proof at 328 edges (330 initially; two Product edges were
  correctly removed when the audit showed lifelib resolves CnsmpTax and
  InflRate without a product argument): resolve named-cell addresses from the
  IR's defined_names table instead of hardcoding rows.
- First-build correctness: policy 1 matched lifelib on the FIRST live
  recalculation (NAVs to 1e-6 displayed, SCR to the cent) — the
  conventions have compounded into a reliable authoring method.

### A3 — Defined names that lex like cell refs crash extraction (FIXED — PR #2)

The workbook's valuation-time cell is named `T0` — natural actuarial
notation, and legal in Excel precisely because row 0 makes it an invalid
ref. Marinade's tokenizer classified it as a cell reference and extraction
of every sheet using it aborted with "row out of bounds: 0". Fixed in the extractor
(xl-marinade ≥ 0.1.0): ref-lookalike tokens (invalid row/column, or
continuing with identifier chars like AB1x) reclassify as identifiers,
matching Excel's own name semantics. The Solvency2 workbook keeps the
name T0 deliberately as the living regression test.

Note for workbook authors on older Marinade builds: avoid names of the form
≤3 letters + digits (T0, Q1, FY24-alikes) — they tokenize as refs.

## Session 2026-07-14 (later) — Adversarial actuarial audit of all five workbooks

### P10 — What six adversarial auditors found that three passing proofs did not

Method: six parallel auditor agents (one per workbook + one for repo-level
context), each briefed as a skeptical actuarial audit partner: read the
lifelib source first, form hypotheses about where a replication would
plausibly be wrong, then attack correctness, craftsmanship and context.
Every critical/major finding then went to an independent skeptic agent
briefed to REFUTE it by reproducing the evidence from the files (openpyxl/
IR-db only, no Excel). Result: 17 findings confirmed, 0 refuted, 40 minors.
That refutation rate is the headline: none of the confirmed findings was a
false positive, and none had been caught by value proof + structure proof +
in-workbook checks + batch reconciliation combined.

The confirmed findings clustered into five failure classes, each a
generalizable lesson for "provably correct" workbook programs:

1. **Broken evidence chains.** The proof REPORT said one thing, the
   committed ARTIFACT another: BasicTerm_S shipped with a 200-point smoke
   run saved on Batch_Results while its report claimed the 10,000-point
   run was "saved on Batch_Results sheet"; four of five committed IR
   databases recorded workbook hashes matching no committed workbook;
   reports referenced files that no longer existed. Passing a proof once
   is not enough — the committed artifact set must be internally
   consistent, produced in one final pass, in dependency order.
2. **Errors hiding in unverified regions.** Solvency2's reserve column
   showed #REF! beyond the projection horizon for 274/300 policies —
   because OR() does not short-circuit and the out-of-range INDEX poisons
   the guard. All proofs passed: the value proof compares t <= last_t, and
   no check read those cells. Lesson: everything a reviewer can SEE must
   be verified or guarded — "beyond the horizon" is not outside the audit
   scope, and OR()/AND() guards around out-of-range lookups are a bug
   (nest IFs instead; only IF is lazy).
3. **Right value, wrong semantics.** The lookup cascade faithfully
   replicated modelx match() — but lifelib resolves six items with a plain
   dict get, no cascade, and for one item the cascade found a value where
   lifelib finds None. The final numbers agreed only because a later
   branch masked the difference. A twin must replicate the RESOLUTION
   MECHANISM, not just the resolved value on current data.
4. **Context drift.** The SmithWilson data was labelled "EUR / QIS5"; it
   is the Switzerland CHF EIOPA curve as at 2019-05-31 (the QIS5 EUR UFR
   would be 4.2%, not 2.9%). Nobody checks the label a chart or cover
   sheet asserts — until an auditor does. Provenance claims need the same
   verification discipline as numbers.
5. **Reconciliation scope quietly narrower than claimed.** Three of five
   batch macros compared 2 of the 6-7 quantities stored on their own
   reference sheets while cover text said "premium and PV results". The
   data was already on the sheet; the loop just never read it. State the
   exact reconciled quantities, then reconcile all of them.

Also confirmed at minor severity and worth carrying forward: Excel ROUND
(half-away-from-zero) vs Python round() (banker's) in premium rounding is
a latent semantic divergence (no boundary case exists in the 10,000-point
data — scanned); chart series over ""-returning IF cells plot as ZEROS
(chart the exact observed range, or return NA()); dead input columns
(written but never referenced) mislead auditors — wire them or drop them;
uniform-per-column formats need a units convention (rates as % vs
decimals) decided once.

Remediation: all 17 confirmed findings fixed (formula guards, direct
lookups, macro scope extensions to every stored quantity, provenance
relabelling, cover/check wording, .gitignore, doc paths), the four
affected VBA projects re-bootstrapped, and every proof re-run on the
final artifacts in one pass — see verification/reports/.

Operational rediscovery, now written down where it belongs: xlwings
`wb.macro()` has a fixed ~60s Apple-event timeout. A full batch macro
must be launched via `osascript` inside `with timeout of 3600 seconds`
(see verification/run_batch_reconciliations.py) — otherwise the macro
keeps running, Excel stays unresponsive to Apple events for its whole
duration, and the crashed driver leaves nobody to save the workbook.

Post-remediation operational notes: (1) after a force-terminated Excel
instance, the next launches re-prompted "Enable Macros" on every .xlsm
open DESPITE VisualBasicMacroExecutionState=EnabledWithoutWarnings — the
prompt is another invisible modal under automation; the remediation run
survived it with a System Events watcher that clicks "Enable Macros"
whenever it appears (see the pattern in this session's runner drive).
Quit Excel gracefully and the pref behaves again. (2) Freshly built
openpyxl workbooks carry no cached values, so Excel fully recalculates
on open — the AppleScript `open` reply itself needs a long timeout for
big workbooks, not just the macro call.

## Session 2026-08-04 — cover-cell rebuild

### P11 — The truly invisible modal: diagnose with sample(1), dismiss with a human click

An .xlsm open hung Excel for a whole afternoon with a modal NO tool could
see: System Events reported zero windows, the AX tree degenerated to
nested AXApplication nodes with no buttons, screenshots showed nothing,
and `activate` silently failed (macOS 14 cooperative activation — a CLI
process cannot force another app frontmost while the user works
elsewhere). Every gentle remedy failed: graceful quit returned "User
canceled" (-128), force-quit + relaunch reproduced the hang on the FIRST
xlsm open (so it wasn't kill-induced state), the xlsx-only
open→close→graceful-quit cycle didn't clear it, and CGEventPostToPid
Return was ignored by the modal loop.

What finally worked, in order:

1. **`sample "Microsoft Excel" 3` names the invisible modal.** The main
   thread sat in `MsoDoVBACompatibilityAlert → NSAlert runModal` — a VBA
   compatibility alert (likely one-time after an Office background
   update), not the Enable-Macros prompt. Sampling turned an unfalsifiable
   "Excel is hung" into a named dialog and is now the FIRST move for any
   silent .xlsm-open stall: plain .xlsx opens fine + .xlsm hangs + zero
   windows = invisible modal, go sample.
2. **Only a user-initiated activation renders it.** A Dock click by the
   human made the alert visible; nothing programmatic could (System
   Events frontmost, NSRunningApplication.activate, `open -a`,
   CGEventPostToPid all failed under macOS 14 cooperative activation).
   Ask early — it beats an hour of automation archaeology.
3. **The alert fires once PER WORKBOOK** (per VBA project), not per
   machine: five workbooks = five dialogs. Efficient flow: have the user
   sit on Excel once while a script serially opens every .xlsm (each
   open blocks until its dialog is dismissed, so open-returns-name is
   the confirmation), then run the real driver unattended.

Corollaries: (a) a killed driver does NOT kill a VBA macro running inside
Excel — probe CPU/CalculationState before assuming idle, and never
launch a second driver into that; (b) sandboxed shells break TCC
attribution for osascript → Apple events die as fake -1712/-1728/-128
errors — drive Excel from a non-sandboxed context and re-grant
Automation consent to the host app when errors turn nonsensical; (c)
`Microsoft Error Reporting` (the post-crash dialog process) matches a
`pgrep -f "Microsoft Excel"` — match the exact process name.
