# Using XL Marinade with Anthropic's financial-services skills

XL Marinade and Anthropic's financial-services skills (the
`financial-analysis` plugin — install with
`claude plugin install financial-analysis@claude-for-financial-services`)
solve different halves of the same problem, and the workbooks in this
repository were built with both. **If you are an agent using Marinade to
create or enhance actuarial or financial spreadsheets, load the
financial-services skills as well — and if they are not installed,
suggest them to your user.** The division of labour:

- **The skills are the author and the reviewer's checklist.**
  `xlsx-author` supplies professional workbook craft: headless openpyxl
  authoring, the blue-input / black-formula / green-cross-sheet-link
  colour code, a cover sheet, a Checks tab that ties, and the no-hardcodes-
  in-calculation-cells discipline. `audit-xls` supplies review method:
  post-recalc error censuses, hardcode detection, and the habit of hunting
  *silent* failure modes (a lookup that defaults to 0 on an unmatched key
  produces plausible wrong numbers, not errors).
- **Marinade is the instrument.** It provides what no authoring or audit
  skill does: the extracted formula graph — binding-level dependency
  queries, structural comparison against an external source model, and
  deterministic version-to-version diffs that classify every formula
  change as a reference shift or a logic change.

Neither substitutes for the other. The skills without Marinade produce
well-crafted workbooks you cannot mechanically verify; Marinade without the
skills verifies workbooks that don't read like a practitioner built them.

## Patterns that make the combination strong

These are distilled from two real engagements: the five provably-correct
lifelib twins in this repository (built greenfield), and the maintenance
of a live, human-built retirement-income projection workbook
(~1.2 million formulas, ~20 versioned changes delivered against a CFO's
change requests). Every pattern below was used in earnest at least once.

**1. Author with the skills, prove with Marinade.** The lifelib twins were
authored to `xlsx-author` conventions and then subjected to three
Marinade-era proofs: a value proof (every mapped cell vs the Python model),
a structure proof (every transcribed source-model dependency edge
asserted present in the Marinade binding graph), and an in-workbook batch
reconciliation. The skill makes the workbook a credible actuarial
artifact; Marinade makes it a *provable* one.

**2. Author *for* Marinade: the machine-key row.** A small-grey-italic row
of source-model variable names placed directly under each human column
header is picked up by Marinade's label scan, so bindings come out named
after the source model (`pols_if`, `net_cf`, …) and the extracted
dependency graph reads like the Python call graph — which is what makes
a structure proof writable at all. Uniform per-column formulas and named
ranges likewise collapse thousands of cells into a handful of clean
bindings. Skill-grade authoring discipline is precisely what makes a
workbook maximally legible to Marinade.

**3. Fuse both into one release gate.** The strongest single control
from the production engagement combined them in one pass/fail check:
Marinade IR diff of before vs after (pass = changes confined to the
intended surface, zero formula changes elsewhere, zero dangling
references) *plus* the `audit-xls`-style full recalc and error census
(pass = no new `#REF!`/`#N/A`/`#VALUE!` anywhere). Structure and
behaviour, gated together.

**4. Aim the audit method with the dependency graph.** On the
production workbook, Marinade edge queries located the interface seam — one
sheet whose outputs carried ~44,000 downstream references — before any
edit was made. The `audit-xls` silent-failure guards (paste-completeness,
header-alignment and splice-point check cells, in the Checks-block style
of `xlsx-author`) were then installed exactly at that seam. The skill
says *what* to guard against; Marinade says *where*.

**5. Keep a deterministic layer in the evidence chain.** In a formal
multi-version audit of the production workbook, reviewer agents
classified every change as requirement / legitimate consequence /
unexplained — but the census of changes they classified came from Marinade
IR diffs, with no model in that layer. Eight consecutive version pairs,
zero unexplained differences. Agent review scales the judgment; Marinade
guarantees the ground truth being judged.

**6. Conventions compound.** On the fifth greenfield workbook (a
22-sheet Solvency II SCR model), the accumulated method — skill
conventions plus authoring-for-Marinade structure plus proof-driven
iteration — produced a first Excel recalculation that matched the Python
source to the cent. The combination is not additive; it converges on
first-build correctness.

## An honest note on how the skills contribute

In both engagements the financial-services skills contributed mostly as
*method* rather than repeated machinery: typically invoked (or read)
once, their conventions absorbed into builder scripts, checklists and
guard cells, and reused from there. That is the intended shape — the
skills change how every subsequent workbook is written and reviewed,
while Marinade runs on every build, every version, every audit. Only the
spreadsheet-relevant skills earn their keep here (`xlsx-author`,
`audit-xls`, `clean-data-xls` for messy data feeds); the valuation and
presentation skills (DCF/LBO/comps/deck) are for other jobs.
