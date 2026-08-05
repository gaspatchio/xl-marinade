# Provably-correct lifelib workbook twins — agent instructions

Provably-correct Excel twins of lifelib actuarial models. See README.md for
the hypothesis and roadmap, marinade-notes.md for XL Marinade usage
patterns and hard-won operational lessons.

## Excel coordination (MANDATORY)

macOS Excel is one shared instance for the whole machine, and agents in
other workspaces may also drive it. Rules:

1. **Hold the machine-wide lock for ALL Excel work** — xlwings, AppleScript/
   `osascript`, or anything that opens/recalculates a workbook:
   - ad-hoc commands: `python3 verification/excel_lock.py run --purpose "<what>" -- <command…>`
   - multi-step sessions: `python3 verification/excel_lock.py acquire --purpose "<what>"` … `release`
   - Python scripts: `from excel_lock import hold` / `with hold("<what>"):`
   The repo's verification and bootstrap scripts already take it themselves.
2. **NEVER `pkill`/force-quit Excel** — that kills other agents' runs and
   corrupts Excel's add-in state. If Excel seems stuck, first run
   `python3 verification/excel_lock.py status` to see who holds it; a live
   holder means WAIT. Ask the user before any force-quit.
3. If a run seems hung, remember VBA `MsgBox`/dialogs in a hidden Excel are
   invisible modals — with Accessibility granted, inspect via
   `System Events` (marinade-notes.md P5); if that shows NOTHING and .xlsm
   opens hang while .xlsx opens work, run `sample "Microsoft Excel" 3` to
   name the invisible modal, and ask the user for one Dock-click to render
   and dismiss it (P11).

## Key facts

- For any workbook job, start with the `actuarial-excel-modeling` skill
  (`.claude/skills/`) — it routes between Anthropic's financial-services
  skills (`financial-analysis` plugin) and Marinade per use case; rationale in
  MARINADE_AND_FINANCIAL_SERVICES_SKILLS.md. If the plugin is not installed,
  suggest it to the user
  (`claude plugin install financial-analysis@claude-for-financial-services`).
- Python env: `.venv` (lifelib, modelx, openpyxl, xlwings, pandas).
- Marinade extractor: `pip install xl-marinade`, then
  `marinade extract <xlsx/xlsm> -o <db>`; query the `agent_*` views.
- Workbooks are GENERATED — edit `workbooks/<Model>/build_workbook.py`, never
  the xlsx/xlsm by hand. VBA lives in `<Model>/*.bas`; after editing, run
  `bootstrap_vba.py` (needs Accessibility) then `build_workbook.py`.
- Every workbook must pass: value proof, structure proof, in-workbook VBA
  all-points reconciliation (reports committed under verification/reports/).
- Workbook style: Control sheet for run settings (model point selection is
  NOT an assumption), blue inputs / black formulas / green cross-sheet links,
  machine-key row under headers (Marinade auto-labels bindings from it).
