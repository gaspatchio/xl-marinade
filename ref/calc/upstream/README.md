# IronCalc 0.9.1 gaps: synthetic reproductions

One standalone script per gap. Each builds a small synthetic workbook with openpyxl in a
temporary directory, loads it with IronCalc 0.9.1, and prints Excel's documented result
next to IronCalc's. Run any of them with:

```bash
uv run --extra calc python ref/calc/upstream/<script>.py
```

`marinade calc` refuses or flags every case below. Output recorded on 2026-11-19 with
`ironcalc` 0.9.1.

| Script | Excel | IronCalc 0.9.1 |
|---|---|---|
| `vlookup_text_vs_numbers.py`: approximate `VLOOKUP` of text against numbers | `#N/A` | `40.0` (a row is returned) |
| `implicit_intersection_name.py`: `=Col`, `Col` naming `S!A1:A5`, entered in row 3 | `300` | `#VALUE!` |
| `sheet_qualified_range.py`: `SUM(S!$A$1:S!$A3)` | `6` | `#ERROR!` |
| `date_1904.py`: `DATE(2020,1,1)` with the 1904 date system | `42369` | `43831.0`, no error |
| `precision_as_displayed.py`: `1/3` shown as `0.00`, times 3 | `0.99` | `1.0`, no error |
| `defined_name_constant_formula.py`: a name holding `0.05`, then one holding `Sheet!$A$1*2` | `0.05`, `20` | `#NAME?` for both |
| `match_whole_column.py`: `MATCH(25,A:A,1)` over 10, 20, 30 | `2` | `1048576.0` |
| `cell_address_other_sheet.py`: `CELL("address",B!C3)` | an address | `#N/IMPL!` |
| `norm_dist_precision.py`: `NORM.DIST(x,0,1,TRUE)` against the exact CDF | within about 1e-15 | up to `2.26e-11` away on a grid of 401 points |

Notes.

- `norm_dist_precision.py`: at `x = -0.0707687815502868` the value matches the exact CDF
  to the last digit. A gap of about 4e-8 was reported for that point and does not
  reproduce here; the script scans a grid instead and shows the worst point found.
- 1904 dates and precision as displayed refuse the whole workbook (exit code 6).
- The defined-name rejection is why `marinade calc` refuses cells that read a name which
  holds a formula or constant.
