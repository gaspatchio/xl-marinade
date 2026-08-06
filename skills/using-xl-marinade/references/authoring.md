# Authoring a new provably-correct workbook (marinade-first)

The concrete builder/proof scripts referenced below (`build_workbook.py`,
`structure_proof_*.py`, `verify_*.py`) live in the companion examples repo —
here they name a **pattern to transcribe**, not files to expect.

## Workflow for a new twin

1. **Read the source model first** — grep the source (e.g. a lifelib
   model's cell definitions) before building anything; when building a
   variant, diff it against the model you already twinned to scope the
   delta.
2. **Generate reference values** — run the source model, dump every cell's
   vector + scalars to JSON. Cells the source can't evaluate (e.g.
   mortality at negative duration) get null-markers; the value proof
   compares guarded cells only where active.
3. **Write a builder script** (headless openpyxl). Workbooks are
   GENERATED — never hand-edit the xlsx/xlsm.
4. **Chain the iteration loop as one command** (structure proof before
   value proof — it needs no Excel and is the fast gate):

   ```bash
   python build_workbook.py && \
   rm -f out.ir.db && \
   marinade extract book.xlsx -o out.ir.db && \
   python structure_proof.py && \
   python verify.py
   ```
5. **VBA layer**: edit `.bas` → bootstrap `vbaProject.bin` once (see
   excel-ops.md) → builder injects the bin via zip surgery on every
   rebuild → re-extract IR (the macro is auditable structure — procedures
   surface in `marinade_nodes`, with detail in the `vba_*` tables) → run the
   in-workbook all-points reconciliation.
6. **Regression between versions**: `marinade diff`, not re-proving (see
   SKILL.md).
7. **Adversarial audit last** — passing proofs still leave the failure
   classes in SKILL.md's checklist. Audit subagents work file-level only
   (openpyxl + sqlite + source model, never Excel) so the Excel lock stays
   uncontended; send every critical finding to an independent skeptic
   agent briefed to REFUTE it from the files.
8. **Final pass**: rebuild and re-verify all committed artifacts in one
   run, in dependency order, so reports/IR dbs/workbooks hash-tie.

## Author FOR the extractor

These conventions make the IR graph self-labeling and directly comparable
to the source model:

1. **Uniform column formulas.** One R1C1 pattern per projection column
   (IF-guards, not per-region formulas) → one binding per column. Works
   for 2D blocks too: a uniform 65×25 grid collapses to ONE binding. The
   init-merger tolerates a different t=0 seed row; if it stays a separate
   binding, that is where seed edges (e.g. `pols_if(0)=init`) live.
2. **Machine-key row.** Put the source model's cell name (small grey
   italic) directly under each human column header — the extractor's label
   scan adopts it verbatim, so `agent_binding_dependencies` reads like the
   source model's call graph. Labels are a legibility aid for whoever
   writes the proof maps — proof scripts themselves must match on binding
   ADDRESSES and `defined_names`, never on labels.
3. **Named ranges** for anything you will verify or let VBA touch; they
   surface in `defined_names` and as label candidates. Resolve named-cell
   addresses in proofs FROM `defined_names`, never hardcode rows.
4. Keep lookup-table headers NUMERIC with display formats (`'"Term "0'`),
   never text — a text header makes every dependent MATCH #N/A. Know that
   INDEX/XLOOKUP/CHOOSE/VLOOKUP/MATCH emit a `DYNAMIC:` marker plus their
   argument refs; INDIRECT/OFFSET stay opaque — avoid them in generated
   models.
5. **Array math must be CSE.** Plain openpyxl-written formulas get legacy
   semantics: Excel 365 silently inserts implicit intersection (`@`)
   before range names used as array arguments (`TRANSPOSE(@Vec)` →
   #VALUE!), invisible except via FORMULATEXT in live Excel. Write any
   formula feeding a whole range into TRANSPOSE/MMULT/MINVERSE or
   INDEX-slice arithmetic as an openpyxl `ArrayFormula` (even
   single-cell). Plain SUMPRODUCT over same-shape ranges is safe. Marinade
   handles CSE fine: a multi-row array formula is a single-cell binding at
   its top-left anchor (map vector cells to that anchor in structure
   proofs), with edges extracted from inside.
6. **Guards must nest IFs.** `OR()`/`AND()` evaluate all arguments — an
   out-of-range INDEX inside a guard yields #REF! in every "guarded" cell
   beyond the horizon. Only IF is lazy.
7. **Style conventions**: a Control sheet for run settings (model point
   selection is NOT an assumption), blue inputs / black formulas / green
   cross-sheet links, machine-key row under headers. When modifying someone
   else's workbook, THEIR conventions win.
8. **Inheritance/stress patterns**: parameterize one engine layout into N
   scenario sheets, each with a parameter block; override formulas branch
   on those parameters exactly where the source's override cells replace
   base cells. Shared schedule sheets keep the engines thin. modelx
   `match()` translates to an IFERROR/MATCH cascade — but ONLY for items
   the source actually resolves via match(); items resolved by plain dict
   get must use a direct lookup (right value, wrong mechanism is a
   finding).

## The two proofs

- **Value proof**: recalc in Excel (xlwings, under the lock), compare
  every mapped cell against source-model vectors for diverse profiles;
  1e-9 relative. Chunk long proofs into per-profile foreground runs — long
  background Excel jobs die, and machine sleep mid-run wedges windowless
  Excel.
- **Structure proof**: transcribe the source model's call graph as
  (dependent, precedent) pairs, map each source cell to a representative
  workbook cell, assert an `agent_binding_dependencies` edge whose from/to
  ranges contain them. Map table-joins per consumed column, seed rows to
  t=0 cells. Keep a known-gaps dict to classify extractor gaps separately
  from workbook errors — and before trusting a FAIL, check the stored
  `agent_cells_light.formula`: if the reference is present verbatim, it's
  an extractor gap, not a workbook error. Expected edges can be wrong too —
  audit the proof script against the source before "fixing" the workbook.

## VBA batch reconciliation macros

- Write the macro against **named ranges** only — the same `.bas` then
  serves sibling workbooks; each workbook still needs its own
  `vbaProject.bin` bootstrap (doc modules bind per sheet count).
- After `Application.Calculate`, spin
  `Do While Application.CalculationState <> xlDone : DoEvents : Loop`,
  then verify a sentinel formula cell echoes the point id just written;
  retry Calculate, hard-error after N retries. Without this, sustained
  set-value/calculate loops intermittently read the PREVIOUS point's
  results — Calculate can return before the dependency chain finishes.
- Never MsgBox — write errors to a results sheet (invisible modal
  otherwise). Reconcile ALL quantities stored on the reference sheet and
  state them.
- Models without model points: reconcile the full result vectors instead
  of looping a batch — the "reconcile everything against the reference
  sheet" convention carries over unchanged.
