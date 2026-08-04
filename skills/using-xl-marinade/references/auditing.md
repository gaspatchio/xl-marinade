# Auditing / modifying an existing workbook (marinade-first)

On human-built workbooks the `agent_*` labels degrade (no machine-key
row), so work at cell level; expect thousands of bindings and minutes per
extraction.

## Per-version workflow (repeat for every change bundle)

1. `cp` previous version → new version file. The pristine baseline is
   immutable; one version file per change bundle; document each version's
   intent (the change doc doubles as the diff's expected-change spec).
2. **Reconnoiter** with marinade SQL (dependency seams, anchors — recipes
   below) + openpyxl (values, styles, layout). Print `wb.sheetnames`
   before guessing sheet names (Excel truncates to 31 chars).
3. **Edit headlessly** via a persisted `build_vNN.py` (openpyxl) in
   scratch — rerunnable if the session dies. Close the file in Excel
   first; see excel-ops.md for the ArrayFormula-translation hazard.
4. **Recalc in real Excel** via osascript (open → calculate → save →
   close), under the Excel lock. Never trust cached values after a
   headless edit; probe one known computed cell for a real number.
5. **Value tie-out** (openpyxl `data_only=True` on the recalced file):
   cell-by-cell diff of all downstream sheets vs the prior version;
   expect ZERO diff outside the intended change surface.
6. **Structural proof**: extract the IR of the new version
   (`marinade extract`), `marinade diff` vs prior, classify every change
   into expected buckets; anything outside the list is a finding.
7. Update the change docs / version log.

## Find the seam before you cut

- **Consumer census** (SKILL.md has the query): one `agent_dependencies`
  query gives cross-sheet edge counts and identifies the interface sheet
  whose output layout you must preserve; everything behind it is
  replaceable. A large blast radius with a boundary is tractable surgery;
  without the boundary it isn't.
- **Anchor discovery**: search `agent_cells_light.formula_r1c1` (or
  `.formula`) for a defined name before assuming a wide edit — one anchor
  name can mean repointing two cells and touching nothing downstream.
- **Dead-tab check**: an `agent_dependencies` query showing zero
  cross-sheet consumers of a tab → safe to delete; nonzero → escalate as a
  decision, don't decide unilaterally.
- **Comment inventory**: the `cell_comments` base table before/after a
  rebuild — the owner's cell comments must survive. (It's version-gated —
  capability-check `sqlite_master` before querying, older dbs lack it.)

## Separate mechanism from data (the stage-1 twin)

When replacing a methodology AND its data: first build a twin with the
NEW mechanism on the OLD data, recalc, compare cell-by-cell across all
downstream sheets. Zero mismatches proves the mechanism swap changes
nothing; every difference in the final version is then attributable to
the data update alone.

## Attribute numeric changes with the graph, not by eyeballing

When a number moves and someone asks why (or is skeptical):

1. Enumerate the changed cells' direct consumers from `agent_dependencies`.
2. Independently replicate the affected line from raw inputs, outside
   Excel, and reconcile the deltas to the cent.
3. Enumerate the flow-on paths (tax, fees) and verify each.
4. Prove NON-paths: "sheet B has zero references to sheet A" is a marinade
   query, and it separates coincident movers from consequences.

Eyeball attribution routinely blames the wrong driver; the graph trace
finds coincident second drivers.

## Version-chain / adversarial audit

- Extract every version once (background; cache `.ir.db`s in scratch),
  then `marinade diff` on each consecutive pair; read `summary` +
  `binding_map` first, drill into `changes` only for pairs whose summary is
  unexpected. Review every formula/value change against documented intent —
  anything not traceable to a documented change is flagged.
- A zero-unexplained-changes claim covers the **`layer: "workbook"`**
  changes only. The **`layer: "ir_inference"`** changes (`TABLE_CANDIDATE_*`,
  `BINDING_LABEL_EVIDENCE_CHANGED`; counted in `summary.ir_inference_changes`)
  are extractor inference that churns benignly across versions — they need
  no per-event adjudication and must not be counted as edits.
- Diff only adjacent versions of one lineage; diverged workbooks produce
  noise at the scale of the whole book — use targeted SQL instead.
- **Subagent fan-out contract**: marinade never needs Excel, so IR
  extraction/diff/SQL can run in parallel read-only subagents while Excel
  is busy. Rules to put in every audit-agent prompt: read-only; headless
  openpyxl (`data_only=True` only if Excel recalced and saved the file;
  `data_only=False` for formulas); NEVER open Excel; artifacts to scratch
  only. Give the final diff agent the expected change surface enumerated
  inline. The parent re-verifies each load-bearing claim itself.
- Typical catches for such a fan-out that proofs miss: approximate-match
  VLOOKUP on an unsorted table; fixed-size ranges (MIN/SUM over N rows)
  left stale after a widening.

## Widening / structural edits

- **Test with POPULATED slots, not empty ones.** Clone real records into
  every widened slot — hardcoded per-row literals compute silently wrong
  values against hidden stale IDs, and every empty-slot test passes over
  them.
- Widening scripts must handle ArrayFormula cells explicitly (translate
  text AND `ref`; widen ranges inside them) — skipping them leaves stale
  spill ranges that corrupt the file.
- Build permanent in-workbook guard cells for silent failure modes
  (`OK / MISMATCH / INCOMPLETE PASTE`) — e.g. XLOOKUP defaulting to 0 on a
  short paste is invisible without one.

## Working with the owner's conventions

Their conventions win over generic ones (e.g. an existing paste-zone fill
convention beats your house input style). Generic spreadsheet-audit
checklists are useful as method; marinade supplies the two capabilities
they lack: dependency mapping and version diffing.
