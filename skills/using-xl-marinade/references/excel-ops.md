# macOS Excel + openpyxl operations — hard rules

Read this before ANY work that opens Excel (xlwings, osascript) or edits
a workbook file headlessly. Marinade itself never needs Excel — these rules
are for the value-proof / recalc / authoring steps around it.

## The shared instance

macOS Excel is ONE instance machine-wide; other processes/agents may
drive it too.

- **Hold a machine-wide lock for ALL Excel work.** Use a lockfile-based
  helper (acquire before / release after; a `with hold(...)` context
  manager is the clean shape) so concurrent agents queue instead of
  colliding.
- **NEVER pkill/force-quit Excel.** It severs other processes'
  Apple-event connections (-609/-1712), corrupts add-in state, and
  re-arms the "Enable Macros" prompt.
- **An Excel "hang" is someone else's workbook until proven otherwise.**
  Diagnostic order: (1) check the lock's holder; (2) ask Excel for its
  ACTIVE WORKBOOK NAME via osascript — a foreign name means wait; (3) with
  Accessibility granted, inspect for dialogs via System Events.
- Apple events TIME OUT (not fail) while any VBA macro runs — a deaf
  Excel may be healthy and mid-macro.

## Invisible modals (hidden Excel blocks silently)

- `MsgBox` in an unattended macro; VBA runtime-error dialogs.
- Excel's "Enable Macros" prompt on .xlsm open. **Trust boundary: only
  bypass this for workbooks you generated or fully trust — never for an
  untrusted or user-supplied `.xlsm`.** Marinade's `extract` parses VBA
  **statically and never executes it**, so auditing or documenting an
  arbitrary workbook needs no macro execution at all — enabling macros is
  only for running a value-proof/reconciliation macro in a workbook you
  built. If you genuinely need it (trusted workbook, unattended run), relax
  the prompt only with a **guaranteed cleanup that survives timeouts, modals,
  and interrupts** — a shell `trap … EXIT` (or Python `try/finally`), never a
  bare write-then-delete that a hang can skip:

  ```bash
  # Excel quit first
  restore() { defaults delete com.microsoft.Excel VisualBasicMacroExecutionState 2>/dev/null; }
  trap restore EXIT
  defaults write com.microsoft.Excel VisualBasicMacroExecutionState -string "EnabledWithoutWarnings"
  # ... run the trusted macro ...
  ```

  Then **verify the pref is back to prompting before opening any other
  workbook** (`defaults read com.microsoft.Excel VisualBasicMacroExecutionState`
  should error or not return `EnabledWithoutWarnings`). Do **not** auto-approve
  the prompt with a System Events clicker on any workbook whose macros you did
  not author — that executes arbitrary embedded VBA with your Excel privileges.
- Detection requires Accessibility granted to the HOST app (the terminal
  or IDE), and the host restarted after granting — otherwise
  `osascript is not allowed assistive access (-1728)` and you are blind.

## Recalc and cached values

- **openpyxl saves strip ALL cached values.** The next Excel open does a
  full cold recalc (minutes on a large book) — give even the `open` call
  a long timeout, and never read values (openpyxl `data_only=True`,
  marinade's cached values, tie-outs) from a file Excel hasn't recalced
  AND saved since the last headless edit.
- **Probe a computed cell after every recalc-and-save**: a save under
  manual calculation can produce a value-less file, and a tie-out then
  "passes" on None==None. A suspiciously FAST recalc of a fuller model is
  evidence calculation never ran.
- Close the workbook in Excel before any headless edit — an open Excel
  copy saved later silently clobbers disk edits.
- Probe formula-vs-parse problems via a scratch cell (`=FORMULATEXT(cell)`
  + `=ERROR.TYPE(cell)` written with xlwings, never saved) — the only way
  to see Excel 365's silently-inserted `@` implicit intersection.

## Long-running work

- xlwings `wb.macro()` has a fixed ~60s Apple-event timeout. Launch batch
  macros via `osascript` inside `with timeout of 3600 seconds` wrapping
  open → run → save → close.
- Anything touching every formula of a big workbook exceeds foreground
  command timeouts — run it in the background and poll artifacts. But very
  long background Excel jobs are fragile — prefer chunked runs that each
  finish in minutes (give value-proof scripts per-profile arguments for
  exactly this).
- Machine sleep mid-run wedges windowless Excel (hung Apple events);
  re-probe Excel scriptability before each chunk.

## openpyxl object-model traps

- Merged ranges: only the ANCHOR cell holds the value; writing a
  non-anchor `MergedCell` raises read-only. `EmptyCell` lacks
  `column_letter`.
- Print `wb.sheetnames` before indexing; sheet names truncate at 31 chars
  and change after reorgs.
- **ArrayFormula/spill cells (LET/XLOOKUP dynamic arrays)**: copying or
  translating them as plain strings keeps a stale `ref` and CORRUPTS the
  file — Excel shows a repair dialog. Use
  `openpyxl.worksheet.formula.ArrayFormula` with a corrected `ref` (and
  `openpyxl.formula.translate.Translator` for the text). Widen ranges
  INSIDE array formulas too.
- Repair-dialog forensics: the report names `xl/worksheets/sheetN.xml`;
  unzip the xlsx and map N → sheet name via `xl/workbook.xml` +
  `xl/_rels/workbook.xml.rels`. Regenerate cleanly; never keep the
  repaired copy.
- VBA writing "TRUE"/"FALSE" strings via `Range.Value` lands native
  booleans.

## xlsm assembly + VBA bootstrap (generated workbooks)

- An .xlsm is a zip: inject `xl/vbaProject.bin`, switch content type to
  `macroEnabled.main+xml`, add the relationship. Stamp
  `wb.code_name = "ThisWorkbook"` + per-sheet codeNames matching the bin's
  doc modules, or the first `ThisWorkbook` reference dies with error 429.
- Producing the bin needs ONE GUI-scripted VBE import (Accessibility
  required). Gotchas: the menu item is literally `Import File...` (three
  ASCII dots, not `…`); the import panel is a WINDOW named "Import File",
  not a sheet; AppleScript `quit` is asynchronous — poll pgrep until Excel
  exits before relaunching; an empty Excel reports workbook list
  `missing value` (don't parse it as a foreign workbook).

## Shell/environment friction

- zsh: no word-splitting of unquoted vars, `=` not `==` in `[ ]` tests,
  failed globs abort compound commands (`setopt null_glob`).
- A lock helper must run on the stock system `python3` — keep it
  dependency-free and old-syntax-safe (no PEP 604 `X | None`).
