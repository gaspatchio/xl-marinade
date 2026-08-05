---
name: actuarial-excel-modeling
description: Job-level playbook for actuarial/financial Excel modelling — combines Anthropic's financial-services skills (xlsx-author, audit-xls, clean-data-xls) with XL Marinade. Use at the START of any such job, before the per-tool skills - building an actuarial or financial workbook (especially replicating a Python/source model provably), changing or enhancing a live human-built workbook, auditing one, reconciling or diffing workbook versions, running a release gate on a model change, or setting up a recurring data-update process. Routes to the right skill/tool at each step; the per-tool skills carry the mechanics.
---

# Actuarial Excel modelling — the combined method

Two toolsets, two roles. **Anthropic's financial-services skills** are the
author and the reviewer's checklist: `xlsx-author` (workbook craft:
colour-coded inputs/formulas/links, cover sheet, Checks tab, no hardcodes
in calc cells, headless openpyxl), `audit-xls` (error censuses, hardcode
detection, silent-failure hunting), `clean-data-xls` (messy feed prep).
**XL Marinade** is the instrument: extracts a workbook's formula
graph to SQLite — binding-level dependency queries, structure comparison
against a source model, deterministic version diffs. The skills make a
workbook a credible practitioner artifact; Marinade makes it a provable one.
Use them together; neither substitutes for the other.

## Before starting any job

1. If the `financial-analysis` plugin is not installed, suggest it to the
   user: `claude plugin install financial-analysis@claude-for-financial-services`.
2. Install the Marinade extractor — `pip install xl-marinade` (the `xl-marinade` skill has the
   mechanics: extraction, `agent_*` views, query patterns, Excel traps).
3. Apply the house conventions in every recipe below:
   - `xlsx-author` colour code — blue inputs, black formulas, green
     cross-sheet links; cover sheet with version control; a Checks sheet
     rolling up to one ALL CHECKS PASS cell; no hardcodes in calc cells.
   - Run settings (model point / scenario selection, batch limits) live on
     a **Control sheet**, never mixed into assumptions — selecting what to
     run is not an assumption.
   - **Author for Marinade**: a machine-key row (source-model variable names,
     small grey italic) directly under human column headers — Marinade labels
     bindings from it, making the extracted graph readable against the
     source model. Uniform per-column formulas and named ranges collapse
     grids into clean bindings.

## Recipe: build a new model workbook

1. Invoke `xlsx-author`; author headlessly (a `build_workbook.py`-style
   generator, never hand-edited artifacts) with the conventions above.
2. Extract with Marinade after each build; query the binding graph to confirm
   the structure you meant to build is the structure you built.
3. Add in-workbook checks mirroring whatever validation the model's domain
   demands (balance ties, decrement roll-forwards, PV identities).
4. Recalculate in Excel before reading values — never trust freshly
   authored cached values.

## Recipe: replicate a source model (Python etc.) provably

Extends the build recipe with three proofs; ship all three:

1. **Value proof** — recalculate in Excel, compare every mapped cell
   against the source model across diverse model points (tight tolerance,
   e.g. 1e-9); commit the report.
2. **Structure proof** — transcribe the source model's call graph as
   (dependent, precedent) pairs; assert each edge exists in Marinade's
   binding-dependency graph. This is what machine-key rows buy you: the
   extracted graph carries source-model names, so the proof is writable.
   State honestly that it is a presence proof (edges exist), not an
   exhaustiveness proof.
3. **Full-population reconciliation** — an in-workbook macro (or batch
   driver) recomputing every model point/policy/run against stored source
   results, verdict on a results sheet, evidence saved in the artifact.

## Recipe: change a live, human-built workbook

1. **Blast radius before any edit**: extract IR, query cross-sheet edges
   to find the interface seams the change touches and how many references
   sit downstream. Plan the surgery to preserve interface layouts.
2. Make the change with `xlsx-author` discipline; add guard cells
   (`audit-xls` silent-failure ethos) at the seams Marinade identified —
   paste-completeness, key-match, splice-point checks that fail LOUDLY.
   Respect the incumbent workbook's own conventions where they conflict
   with skill defaults; the owner's muscle memory wins.
3. **Fused release gate**, one pass/fail: (a) Marinade IR diff before vs
   after — pass = formula changes confined to the intended surface, zero
   changes elsewhere, zero dangling references; (b) full recalc + error
   census in the `audit-xls` mold — pass = no new `#REF!`/`#N/A`/
   `#VALUE!` anywhere, all check cells green.

## Recipe: audit an existing workbook

1. Extract IR; use the binding graph to scope the audit — trace what
   drives the outputs under review, rank seams by downstream reference
   count, and aim `audit-xls` methods (hardcode scan, error census,
   silent-failure hunting: lookups that default rather than fail, approx-
   match lookups on unsorted tables) at those seams.
2. For multi-reviewer or agentic audits, keep a deterministic layer: the
   census of facts reviewers classify (formula changes, error cells,
   dependency edges) must come from Marinade/file inspection, not from a
   model's reading of the workbook — no LLM in the evidence chain.

## Recipe: reconcile or attribute versions

1. Extract IR for both versions; run the Marinade diff. Triage its
   classification first: reference shifts (rows/columns moved) are noise,
   logic changes are the review set.
2. Review every logic change against stated intent (a change register,
   a change request list); classify requirement / legitimate consequence /
   unexplained. The target is zero unexplained.
3. For consecutive-version histories, diff pairwise and gate each pair.

## Recipe: recurring data updates (monthly feeds)

1. `clean-data-xls` for the incoming feed; paste/import into a dedicated
   raw-data sheet behind a stable interface layout.
2. Permanent guard cells at the interface (completeness, alignment,
   splice checks) so a bad paste fails loudly on sight.
3. Marinade diff of last month's vs this month's workbook to surface silently
   restated historic values in the feed.

## Honest scope

The financial-services skills usually contribute as absorbed method —
invoked or read once, their conventions carried forward in builder
scripts and checklists. Marinade runs on every build, every version, every
audit. Only the spreadsheet skills matter here; the valuation and deck
skills (DCF/LBO/comps/pptx) are for other jobs. Rationale and the
engagements behind these recipes: see MARINADE_AND_FINANCIAL_SERVICES_SKILLS.md
at the repo root.
