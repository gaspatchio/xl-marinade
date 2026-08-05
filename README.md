# Provably Correct Actuarial Workbooks

Excel twins of open-source actuarial models ([lifelib](https://lifelib.io),
MIT), each verified cell-by-cell against the Python original — built to test
a hypothesis:

> Financial-services agent skills + [XL Marinade](https://github.com/opioinc/xl-marinade)
> + purpose-built skills let an actuary replicate even the most complicated
> actuarial models from open languages (Python/lifelib, gaspatchio) into
> Excel, provably correctly and highly automatedly — so Excel can serve as
> living documentation of production models.

The workbooks were built with Anthropic's financial-services skills
(`financial-analysis` plugin) and XL Marinade working together —
the skills as author and audit checklist, Marinade as the verification
instrument. Why that combination is strong, with the usage patterns it
produced, is documented in
[MARINADE_AND_FINANCIAL_SERVICES_SKILLS.md](MARINADE_AND_FINANCIAL_SERVICES_SKILLS.md).

## What "provably correct" means here

Every workbook ships with two machine-checked proofs:

1. **Value proof** — the workbook is recalculated in Excel and every mapped
   cell is compared against the Python model across a diverse set of model
   points (`verification/verify_*.py`; tolerance 1e-9, reports committed to
   `verification/reports/`).
2. **Structure proof** — the lifelib call graph is transcribed as
   (dependent, precedent) pairs and each pair is asserted to exist in the
   workbook's formula dependency graph as extracted by XL Marinade
   (`verification/structure_proof_*.py`). This is a presence proof at
   Marinade binding granularity: it certifies every transcribed lifelib
   dependency is wired in, not that no other dependencies exist.

Plus in-workbook consistency checks on a Checks sheet (all TRUE), mirroring
the source model's own validation cells, and an in-workbook VBA batch
reconciliation across every model point / policy / run against stored
lifelib results.

## Workbooks

| Workbook | Source model | Status |
|---|---|---|
| `workbooks/BasicTerm_S` | basiclife.BasicTerm_S — monthly term assurance, selectable model point + all-points batch | ALL PROOFS PASS: value proof (32,710 comparisons, 10 points, 1e-9); structure proof 66/66 lifelib call-graph edges (needs an extractor that emits argument-reference edges from INDEX/MATCH-style lookups — xl-marinade ≥ 0.1.0 does); in-workbook VBA reconciliation 10,000/10,000 points × 6 quantities (premium diff 0.00, max PV rel diff 5.2e-11, 89s) |
| `workbooks/BasicTerm_SE` | basiclife.BasicTerm_SE — in-force portfolio: existing durations (incl. negative = future new business), policy counts, tariff premiums, within-month decrement timings | ALL PROOFS PASS: value proof (24,767 comparisons, 10 in-force profiles, 1e-9); structure proof 88/88 lifelib call-graph edges; in-workbook VBA reconciliation 10,000/10,000 points × 6 quantities (premium diff 0.00, max PV rel diff 6.4e-11, 165s) |
| `workbooks/CashValue_SE` | savings.CashValue_SE — universal life: four-timing AV roll-forward, stochastic lognormal returns (10 scenarios), product-spec tariffs, surrender charges, whole-of-life terms | ALL PROOFS PASS: value proof (153,042 comparisons, 6 point×scenario profiles, 1e-9); structure proof 190/190 lifelib call-graph edges; in-workbook VBA reconciliation 40/40 runs × 7 PV components (max rel diff 4.2e-14, 7.3s) |
| `workbooks/SmithWilson` | smithwilson — Smith-Wilson risk-free curve extrapolation (CEIOPS/EIOPA QIS5 method; data = Switzerland EIOPA CHF term structure, 2019-05-31): 65×25 Wilson function grid, ζ calibrated with native MINVERSE/MMULT array formulas, P/R extrapolated to 65y | ALL PROOFS PASS: value proof (1,986 comparisons incl. the full Wilson grid and the array-solved ζ vector, 1e-9); structure proof 20/20 lifelib call-graph edges; in-workbook VBA reconciliation 155/155 values (max ζ abs diff 1.7e-9 = Excel MINVERSE vs numpy inv, max rel diff P 1.8e-12) |
| `workbooks/Solvency2` | solvency2 — life underwriting SCR (standard formula): commutation-function premium/reserve engine, modelx lookup cascade, 7 scenario sheets (base + 6 life stresses), 7×7 correlation aggregation, 300 policies | ALL PROOFS PASS: value proof (55,227 comparisons, 7 profiles across products/t0/scenarios, 1e-9); structure proof 328/328 lifelib call-graph edges (needs an extractor that tokenizes ref-lookalike defined names — extraction used to crash on the defined name `T0`, kept deliberately as a regression test; xl-marinade ≥ 0.1.0 handles it); in-workbook VBA reconciliation 300/300 policies × 9 quantities incl. lapse sub-shocks (max rel diff 4.7e-13, 10.4s) |
| ifrs17sim | nested projections, CSM roll-forward | stretch (Tier 4) |

## Repo layout

```
workbooks/<Model>/         build_workbook.py, BatchRunner.bas + bootstrap_vba.py,
                           vbaProject.bin, the generated .xlsx and .xlsm
models/                    vendored lifelib model sources (MIT)
verification/              reference generators, value proofs, structure proofs,
                           excel_lock.py (machine-wide Excel coordination)
verification/reports/      committed proof reports
marinade/                  IR databases (regenerated locally - see marinade/README.md)
marinade-notes.md          Marinade usage findings (P1-P10, A1-A3)
.claude/skills/xl-marinade/  the agent skill distilled from those findings
.claude/skills/actuarial-excel-modeling/  job-level playbook combining the
                           financial-services skills with Marinade per use case
```

## Reproducing

Versions the committed references and proofs were produced with: lifelib
0.13.0, modelx 0.31.1, openpyxl 3.1.5, xlwings 0.36.8 (pandas 3.0.3,
numpy 2.5.1); XL Marinade extractor 0.1.0 (`pip install xl-marinade`).
A different lifelib release may legitimately change model behaviour —
regenerate the references before blaming a workbook.

```bash
uv venv .venv && uv pip install --python .venv/bin/python \
    lifelib==0.13.0 modelx==0.31.1 openpyxl==3.1.5 xlwings==0.36.8 pandas numpy
# per workbook (BasicTerm_S shown; same pattern for basicterm_se,
# cashvalue_se, smithwilson, solvency2):
.venv/bin/python workbooks/BasicTerm_S/build_workbook.py       # build the xlsx/xlsm
.venv/bin/python verification/generate_reference_basicterm_s.py
.venv/bin/python verification/verify_basicterm_s.py            # needs Excel (xlwings)
marinade extract workbooks/BasicTerm_S/BasicTerm_S.xlsm \
    -o marinade/BasicTerm_S.ir.db                              # pip install xl-marinade
.venv/bin/python verification/structure_proof_basicterm_s.py   # needs Marinade IR db
```

The IR databases under `marinade/` are produced with
[XL Marinade](https://github.com/opioinc/xl-marinade) (`pip install
xl-marinade`, CLI `marinade`) — see marinade-notes.md. The structure
proofs need xl-marinade ≥ 0.1.0: earlier internal extractor builds missed
lookup-argument edges and crashed on the `T0` defined name (both cases
are kept in the workbooks as living regression tests).

Workbook conventions (all workbooks): blue = input, black = in-sheet formula,
green = cross-sheet link; run control (model point selection, batch settings)
on a Control sheet — separate from basis assumptions; cover sheet with version
control and sheet index; Checks sheet rolls up to a single ALL CHECKS PASS
cell; a batch reconciliation macro (RunAllModelPoints on the projection
workbooks, RunCurveReconciliation on SmithWilson, RunAllPolicies on Solvency2)
recomputes every model point / run / policy and reconciles all stored
quantities against lifelib values on the Lifelib_Reference sheet (verdict on
Batch_Results).

### VBA build note (macOS)

`build_workbook.py` produces the `.xlsm` by injecting
`workbooks/<Model>/vbaProject.bin` (zip surgery) — fully automated once the
bin exists. Producing the bin the first time (or after editing the `.bas`)
needs Excel's VBE, which macOS does not expose to scripting:
`bootstrap_vba.py` automates the import via UI scripting but requires
Accessibility permission for the terminal; alternatively import the `.bas`
by hand once (open the `.xlsx`, ⌥F11, File ▸ Import File…, save as `.xlsm`)
and run `bootstrap_vba.py --harvest <path.xlsm>` to extract the bin.

## License

MIT (see LICENSE). The repository redistributes two third-party MIT
codebases under `models/` — the lifelib model sources and Dejan Simic's
smith-wilson-py (which lifelib's smithwilson project builds on) — see
THIRD_PARTY_NOTICES.md for the required notices and exact provenance.
