# Marinade IR databases (regenerated locally, not committed)

The extracted IR databases (`*.ir.db`, ~35-80 MB each) are build artifacts
regenerable from the committed workbooks:

```bash
pip install xl-marinade
marinade extract workbooks/<Model>/<Model>.xlsm -o marinade/<Model>.ir.db
```

The proofs need xl-marinade ≥ 0.1.0: earlier internal extractor builds
under-counted lookup-argument edges (BasicTerm structure proofs) and
crashed on the defined name `T0` (Solvency2). The committed reports were
produced with xl-marinade 0.1.0.

Each database records the sha256 of the workbook it was extracted from in
its `ir_metadata` table; the committed structure-proof reports under
`verification/reports/` were produced against databases whose hashes match
the committed `.xlsm` files. To re-verify: regenerate the database, then
run `verification/structure_proof_<model>.py`.

`diff_S_vs_SE_summary.txt` is a sample `ir_diff` run comparing the
BasicTerm_S and BasicTerm_SE workbooks at binding granularity.
