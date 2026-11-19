# IronCalc 0.9.1 gaps: synthetic reproductions

One standalone script per gap. Each builds a small synthetic workbook with openpyxl in a
temporary directory, loads it with IronCalc 0.9.1, and prints Excel's documented result
next to IronCalc's. Run any of them with:

```bash
uv run --extra calc python ref/calc/upstream/<script>.py
```

Output recorded on 2026-11-19 with `ironcalc` 0.9.1. The last column says what
`marinade calc` does with the case; it does not catch all of them statically.

| Script | Excel (documented) | IronCalc 0.9.1 | `marinade calc` |
|---|---|---|---|
| `vlookup_text_vs_numbers.py`: approximate `VLOOKUP` of text against numbers | `#N/A` | `40.0` (a row is returned) | A `silently_wrong` trust failure if the stored value differs on the stored inputs; for a new policy, the lookup probe reads the value and the matched key, and a type difference refuses that policy's outputs (`lookup value type differs from keys`). |
| `implicit_intersection_name.py`: `=Col`, `Col` naming `S!A1:A5`, in row 3 | `300` | `#VALUE!` | Refused pattern `implicit intersection on multi-cell name`. |
| `sheet_qualified_range.py`: `SUM(S!$A$1:S!$A3)` | `6` | `#ERROR!` | No pattern. A trust failure (`engine_error`) when Excel stored a value, so the cell and its dependants are refused. |
| `date_1904.py`: `DATE(2020,1,1)` with the 1904 date system | `42369` | `43831.0`, no error | Whole-workbook refusal (exit code 6). |
| `precision_as_displayed.py`: `1/3` shown as `0.00`, times 3 | `0.99` | `1.0`, no error | Whole-workbook refusal (exit code 6). |
| `defined_name_constant_formula.py`: a name holding `0.05`, then one holding `Sheet!$A$1*2` | `0.05`, `20` | `#NAME?` for both | Refused pattern `name holds a formula` (it covers names holding a constant too). |
| `match_whole_column.py`: `MATCH(25,A:A,1)` over 10, 20, 30 | `2` | `1048576.0` | Refused pattern `approximate lookup on whole column`. |
| `cell_address_other_sheet.py`: `CELL("address",B!C3)` | an address | `#N/IMPL!` | No pattern. A trust failure (`engine_error`) when Excel stored a value, so the cell and its dependants are refused; for a new policy, `#N/IMPL!` in an output or probe is an `engine_error` too, even under `IFERROR`. |
