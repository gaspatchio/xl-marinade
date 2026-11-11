# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres
to [Semantic Versioning](https://semver.org/spec/v2.0.0.html). The SQLite output
schema is a versioned public contract.

## [Unreleased]

## [0.6.1] - 2026-11-11

A fix to the VBA status reported without the `[vba]` extra (#57): a workbook
with no macros no longer reads as one whose macros were skipped. Only
`ir_metadata.vba_extraction` and the extraction progress line change; the
output schema is unchanged at 3.0.

### Fixed
- Without the `[vba]` extra, a workbook with no macros no longer reports its
  macros as unread (#57). The oletools check ran before the workbook was
  opened, so every plain `.xlsx` recorded `vba_extraction =
  skipped_no_oletools` and printed `VBA: 0 procs, … (skipped_no_oletools)` --
  exactly what an `.xlsm` whose macros were never read produced, so the two
  could not be told apart, and AI agents read both as "no macros". Now a
  workbook without a VBA project records `no_vba_project` whether or not
  oletools is installed, and only one that has a VBA project, extracted
  without oletools, records `skipped_no_oletools`. Its progress line reads
  `VBA: present, NOT READ -- … install 'xl-marinade[vba]' …` instead of a row
  of zero counts, and the warning names the workbook.

## [0.6.0] - 2026-11-10

A diff-correctness release for structural edits. Inserting or deleting a row
or column inside a sheet is now reported where it happened, with only its real
consequences (#53). Before, an interior column deletion was placed at the
sheet's tail, every column after it read as renamed or edited, and the deleted
columns' bindings were never reported. Diff output changes for any pair with
such an edit: `summary.{rows,cols}_{inserted,deleted}` now count rows and
columns, removed bindings are reported, and every `binding_map` row gains
`shifted`. Extraction output and the schema (3.0) are unchanged.

### Changed
- `summary.rows_inserted`, `rows_deleted`, `cols_inserted` and `cols_deleted`
  in a `marinade diff` changelist now count rows and columns, not
  `ROWS_*`/`COLS_*` events (#53). Deleting seven columns in three blocks used
  to read `cols_deleted: 3`. Each event's `at_row`/`at_col` and `count` now
  also describe one contiguous block: a run of deleted occupied columns
  separated by empty ones used to be one event whose `count` overstated the
  block at `at_col`.

### Fixed
- `marinade diff` placed interior row and column deletions at the sheet's
  tail (#53). Deleting columns from the middle of a formula block changes
  the relative references of every column to their right (`RC[-5]` becomes
  `RC[-3]`), so none of those columns matched themselves, and a fallback
  paired each with whatever now sat at its old address. On a lifelib
  CashValue_SE workbook with four `Projection` columns deleted, the diff
  reported the deletions at the wrong columns, 19 label changes that never
  happened, 21,698 formula changes and no removed bindings: 154,521 changes
  in all. It now reports the deletions at AD, AP:AQ and AU, the four deleted
  columns' bindings as removed, no label changes, and formula changes only
  for the four columns whose formulas became `#REF!`: 7,007 changes. In
  detail:
  - rows and columns that no longer match on their full contents are
    re-aligned with reference offsets ignored, then by their labels and
    constants, and the position-based fallback can no longer pair across
    other matches;
  - relative references that span an inserted or deleted row or column, and
    defined names whose ranges shifted, are compared as Excel rewrites them;
  - a binding that lay entirely in deleted rows or columns is reported
    `BINDING_REMOVED` (with `with_axis_deletion: true`) instead of vanishing,
    and one that lost only its first row or column is still compared;
  - every `binding_map` row gains `shifted`: true when a binding's address
    changed only because rows or columns were inserted or deleted before it.
    Such a binding stays `unchanged`, so filtering on `diff_state` still
    yields only what was edited;
  - `RC[-1]`, `R[-1]C` and `RC5` are recognised as references when formulas
    are canonicalized; only references with a number on both parts were.

## [0.5.0] - 2026-11-09

A VBA release. `marinade diff` now reports macro changes (#47); before this, a
VBA edit, even deleting every module, diffed to nothing. The extracted
database also records whether its VBA was read at all (#48), so a macro
workbook extracted without the `[vba]` extra no longer looks like one with no
macros. Diff output gains `VBA_*` change types and eight `summary.vba_*`
counts, and `ir_metadata` gains `vba_extraction`. The output schema is
unchanged at 3.0.

### Added
- The extracted database records whether its VBA was read (#48):
  `ir_metadata` key `vba_extraction` is `extracted`, `no_vba_project`,
  `skipped_no_oletools` or `failed`. Previously, extracting a macro workbook
  without the `[vba]` extra left every VBA table empty and nothing in the
  database said so, so it read exactly like a workbook with no macros.
  `ir_metadata` is a base table, so `schema_version` is unchanged at 3.0.

### Fixed
- `marinade diff` now reports VBA changes (#47). It never read the VBA tables,
  so a macro edit diffed to nothing: deleting all 13 modules from a lifelib
  SmithWilson workbook produced only `IR_METADATA_CHANGED`. The changelist now
  carries `VBA_MODULE_*`, `VBA_PROCEDURE_*` and `VBA_DECLARATIONS_*` entries
  (all `layer: "workbook"`, ordered after every other change), and `summary`
  gains eight `vba_*` counts, which are zero for workbooks without VBA. Along
  the way, three defects in the previously unused VBA comparison are fixed:
  - the two branches of an `#If` block that define the same procedure
    overwrote each other, so an edit to one branch could go unreported;
  - a procedure with several equally likely rename targets was dropped from
    the result instead of being reported as removed;
  - a rename with an unchanged body was always classified as renamed *and*
    modified, because the comparison included the line naming the procedure.

  Code outside procedures (`Option` statements, `Declare`s, module-level
  variables and constants) is compared too, so any edit to a module's code shows
  up as at least one change.
- A damaged `xl/vbaProject.bin` is now reported as a failed VBA extraction.
  olevba fell back to reading it as plain text, which found no modules, so the
  workbook was silently treated as having none.

## [0.4.0] - 2026-10-03

A correctness release for lookups and diffs. Lookups written against defined
names or `LET` variables now resolve into the graph (#43), and a one-cell edit
no longer makes `marinade diff` report unedited cells as edits (#42). The
first is a **breaking change to extraction output**: the same name-based
workbook extracts to different edges under 0.3.0 and 0.4.0. The output schema
is unchanged at 3.0.

### Changed
- **Breaking (extraction output): lookups written against defined names or
  `LET` variables now resolve like their explicit-range equivalents** (split
  from #42). `INDEX(MP_Age, MATCH(PointID, MP_ID, 0))` previously stopped at
  `partial_resolved` and left `UNRESOLVED:MP_Age` in `cell_edges_external`,
  while the same lookup written with ranges resolved to the row its key
  selects. For every such lookup:
  - it gains a cell edge to the cell its key selects in the cached values;
  - its `UNRESOLVED:<name>` external edge is removed;
  - `resolution_metrics` counts it `resolved`, not `partial_resolved`;
  - the binding edge between the lookup and its table changes from
    `kind = 'range_static'` (weight = table size) to `kind = 'formula'`
    (weight = resolved cells), as it already did for explicit-range lookups.

  Which binding depends on which is unchanged, and workbooks without name-based
  lookups extract identically. Across the five lifelib example workbooks the
  change adds 8–1,155 cell edges and removes 216–4,582 `UNRESOLVED` edges per
  workbook, with no measurable change in extraction time. Formula-valued,
  external and multi-area names keep the previous fallback.
- **Every entry in a `marinade diff` changelist now carries a `layer`** —
  `"workbook"` for an edit, `"ir_inference"` for extractor output that moved as
  a consequence. Previously only inference entries were tagged, so an untagged
  entry had to be read as an edit. `summary.ir_inference_changes` counts the
  inference layer.
- `BINDING_METADATA_CHANGED` is now in the `ir_inference` layer: every field it
  compares (confidence, orphan flag, extraction source, evidence and label-scan
  hashes) is extractor output.

### Fixed
- `marinade diff` reported cells nobody edited as changed (#42). A one-cell
  edit to a lifelib BasicTerm_S workbook produced 13–29 extra entries: every
  table candidate, the bindings around the edited cell, and — once lookups
  resolve — the edges of every lookup keyed by the edited cell. Now:
  - table-candidate members are compared by binding position, not by
    `binding_id` (which is seeded with the workbook's hash, so it differs
    between any two versions);
  - neighbouring bindings whose label-scan context picked up the edited
    cell's value are reported as `ir_inference`, not as edits;
  - an edge change out of a cell whose formula did not change — a lookup
    re-resolved because its key changed value — is `ir_inference`; edges that
    moved because their formula changed stay in the `workbook` layer.

## [0.3.0] - 2026-08-15

A platform-hardening release. 0.2.0 could not complete a single extraction on a
non-UTF-8 Windows console, and several outputs were not byte-reproducible
between two runs of the same input — against determinism being the point of the
artifact. The output schema is unchanged at 3.0.

### Added
- `marinade --version` (and `-V`) report the installed version. Previously the
  only way to learn it was package metadata.

### Changed
- **`ir_metadata` key `extractor_git_sha` is replaced by `extractor_version`.**
  `ir_metadata` is a base table and carries no compatibility guarantee, so
  `schema_version` is unchanged — but anything reading that key by name needs
  updating.
- Extraction telemetry is written to `<database>.telemetry.json` beside the
  output rather than a fixed `telemetry.json`. The old fixed name silently
  overwrote any same-named file in the chosen directory, and two extractions
  sharing an output directory clobbered each other's telemetry.

### Fixed
- Two extractions of the same workbook could disagree. The IR stamped
  `extractor_git_sha` from `git rev-parse` run in the caller's working
  directory with a 5-second timeout, so under load the timeout fired on one
  run and not the next (and any run outside a git checkout recorded
  `unknown`), producing a non-empty self-diff. The stamp is now
  `extractor_version`, taken from package metadata: constant per installed
  version, no subprocess, no dependence on the caller's directory.
- `marinade document` is now byte-reproducible: the generation timestamp in
  `documentation.md` and `model_spec.json` honours `SOURCE_DATE_EPOCH` (the
  reproducible-builds convention), so two runs over one IR database produce
  identical artifacts. Without the variable the wall clock is used, as before.
- `agent_bindings.formula_pattern` now reports the binding's **dominant**
  formula instead of its top-left cell's. On a binding the init merger created
  by folding an initialisation cell into the range it seeds, the top-left is the
  one-off — observed reporting a formula holding 1 cell of 111 on a real UL
  projection. Worse on a recurrence: the init formula carries no self-reference,
  so a consumer reading `formula_pattern` to detect recursion saw none. Ties
  keep the top-left, so single-formula bindings are unchanged.
- `marinade extract -o <dir>/<file>` now creates a missing output directory
  instead of failing with SQLite's bare `unable to open database file`, which
  named neither the cause nor the path and arrived after the CLI had already
  printed the output path as though it were fine. If the directory cannot be
  created, the error now names it.
- A database path containing `#`, `%HH` or `'` no longer opens the wrong file.
  Read-only opens built a SQLite URI by interpolating the path into
  `file:{path}?mode=ro`, so SQLite parsed filename characters as URI syntax: a
  `#` truncated the path at a fragment marker and took `?mode=ro` with it,
  silently opening — and creating — a different, empty database with no
  read-only guarantee, which surfaced as "no bindings" rather than an error.
  Paths now go through `Path.as_uri()`, and ATTACH binds the filename as a
  parameter.
- Generated text artifacts are now byte-identical across platforms.
  `documentation.md`, `model_spec.json`, the `marinade diff` changelist,
  `telemetry.json` and the mutation/usage logs were written in text mode
  without `newline=`, so Python translated every `\n` to `\r\n` on Windows —
  spurious diffs for anything hashing or comparing two runs, against a tool
  whose stated contract is determinism. All text writes now pin `newline="\n"`.
- `apply_actuarial_classification` no longer reports failure for a run that
  succeeded. Its completion banner was the package's only non-ASCII write to
  stdout, which Windows encodes with the ANSI code page under `errors='strict'`
  when redirected — and the banner fires after the overlay database is
  committed, so the `UnicodeEncodeError` turned a finished run into a non-zero
  exit. The banner is now ASCII and goes to stderr.
- A failed or interrupted `marinade extract` no longer destroys the previous
  output database. The final VACUUM wrote to the target after unlinking it, so
  a full disk or a Ctrl-C left the user with neither the old database nor a new
  one; the write now goes to a sibling temp file moved into place atomically.
  Note the tradeoff: the previous database is retained until the new one is
  complete, so extraction now needs room for both at once and fails cleanly on
  a nearly-full disk where it previously succeeded destructively.
- `marinade extract` crashed on ANY workbook under a non-UTF-8 locale
  (`'charmap' codec can't decode byte 0x90` on Windows): text files were
  read and written with the platform's default encoding. All text I/O now
  passes `encoding="utf-8"` explicitly, enforced by ruff's PLW1514.

## [0.2.0] - 2026-08-06

### Changed
- **Breaking (output schema 2.0 → 3.0):** the unified node view `atlas_nodes`
  is renamed to `marinade_nodes`. Same columns and semantics; only the view
  name changes. Databases stamped `schema_version` 3.0 no longer contain an
  `atlas_nodes` view. The bundled `using-xl-marinade` skill queries the new
  name.

### Fixed
- Extraction of rolling-window models no longer spills unbounded SQLite temp
  space ("database or disk is full") in the range-collapse step; replaced by
  chunked breadth aggregation + disjoint-box union counting with identical
  output.
- Lookup-dense sheets (INDEX/MATCH ledgers) no longer re-scan the lookup array
  once per formula, and no longer pay a `runtime_checkable` Protocol check per
  cell read. Bounded-range MATCH scans are memoized per immutable snapshot —
  never from a live `Workbook` source, whose mutations must stay visible — and
  the value source is classified once at construction. A 15k-formula
  INDEX/MATCH probe goes 111s → 5s and a ledger-style benchmark 231s → 17s,
  with byte-identical extraction verified on 8 real workbooks. The memo costs
  ~194 MB peak RSS on a 2.3M-formula model.

## [0.1.0] - 2026-08-02

### Added
- Deterministic Excel formula-graph extraction to SQLite (`marinade extract`).
- Deterministic documentation generation (`marinade document`); optional
  bring-your-own-key LLM enrichment via the `xl-marinade[llm]` extra.
- IR diff (`marinade diff`) emitting a JSON changelist.
- Importable library API: `xl_marinade.extract`, `xl_marinade.diff`,
  `xl_marinade.docs.document`, `xl_marinade.llm.document`.
