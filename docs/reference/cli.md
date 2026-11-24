# CLI reference

```bash
marinade --help
```

```
Usage: marinade [OPTIONS] COMMAND [ARGS]...

XL Marinade — deterministic Excel formula-graph extractor.

Commands:
  extract   Extract a workbook's formula graph to a SQLite database.
  document  Generate documentation for an extracted IR database.
  diff      Diff two IR databases; emit the changelist as JSON.
  calc      Experimental: recalculate a workbook per policy and report trust.
  values    Experimental: write the workbook's stored values as rows.
```

Five commands. `extract` and `diff` are always deterministic and network-free;
`document` is deterministic by default and only makes a network call if you pass
`--enrich` (see [LLM add-on](../llm.md)). `calc` and `values` are experimental; see
[Recalculating](../guide/calc.md).

## `marinade extract`

Extract a workbook's formula graph to a SQLite database.

```bash
marinade extract [OPTIONS] WORKBOOK
```

**Argument**

| Name | Description |
|---|---|
| `WORKBOOK` | Excel workbook (`.xlsx`/`.xlsm`). Required. |

**Options**

| Flag | Description | Default |
|---|---|---|
| `-o`, `--out PATH` | Output SQLite database path. | `ir.db` |
| `--enrich` | Opt-in LLM VBA enrichment (makes network calls; requires `xl-marinade[llm]`). | off |

**Example**

```bash
marinade extract book.xlsx -o ir.db
```

## `marinade document`

Generate documentation for an extracted IR database.

```bash
marinade document [OPTIONS] IR_DB
```

**Argument**

| Name | Description |
|---|---|
| `IR_DB` | IR SQLite database produced by `marinade extract`. Required. |

**Options**

| Flag | Description | Default |
|---|---|---|
| `-o`, `--out PATH` | Output directory (writes `documentation.md` + `model_spec.json`). | `docs_out` |
| `--enrich` | Opt-in LLM enrichment (requires `xl-marinade[llm]` + a key; degrades to deterministic output otherwise). | off |

**Example**

```bash
marinade document ir.db -o out/
```

## `marinade diff`

Diff two IR databases; emit the changelist as JSON.

```bash
marinade diff [OPTIONS] DB_A DB_B
```

**Arguments**

| Name | Description |
|---|---|
| `DB_A` | Version A (earlier) IR database. Required. |
| `DB_B` | Version B (later) IR database. Required. |

**Options**

| Flag | Description | Default |
|---|---|---|
| `-o`, `--out PATH` | Write the JSON changelist here. | stdout |

**Example**

```bash
marinade diff a.db b.db
```

## `marinade calc`

Experimental. Recalculate a workbook for each policy in a CSV and report how far to trust
each value. Needs `xl-marinade[calc]`. See [Recalculating](../guide/calc.md) for the file
formats, statuses and refused patterns.

```bash
marinade calc [OPTIONS] WORKBOOK
```

**Argument**

| Name | Description |
|---|---|
| `WORKBOOK` | Excel workbook (`.xlsx`/`.xlsm`). Required. |

**Options**

| Flag | Description | Default |
|---|---|---|
| `--policies FILE` | Policies CSV, one row per input set. Required. | |
| `--map FILE` | `map.json`: field to cell or name, plus `id_field`. Required. | |
| `--outputs TEXT` | Output cell, range or name (`'Sheet'!Name` for a sheet's own name); repeat for several. | |
| `--variables FILE` | `vars.json`: named variables for `rows.csv`; `period_start` required on every range. | |
| `--check-against FILE` | `check.json`: stored results for other input sets, to widen the trust run. | |
| `--ir FILE` | Reuse this IR database (must match the workbook). | extract afresh |
| `--as-of TEXT` | Valuation date `YYYY-MM-DD`; replaces `TODAY`/`NOW`. | unset (those cells are refused) |
| `--tolerance FLOAT` | Numeric comparison tolerance, relative or absolute. | `1e-09` |
| `--probe-cap INT` | Maximum number of probe cells. When probes do not fit, a warning on stderr gives the cap that holds them all. | `50000` |
| `-o`, `--out PATH` | Output directory. Required. | |

At least one of `--outputs` and `--variables` is required. Writes `results.csv`,
`rows.csv` (with `--variables`), `paths.csv`, `trust_report.json` and `run.json`, after
clearing those names from `--out`; an input file that is one of them is refused first.

**Exit codes**

| Code | Meaning |
|---|---|
| 0 | Success. |
| 1 | Error (bad input, unknown sheet, IR from another workbook, an input that is one of the output files, an `-o` that is a file, the engine crashing on a formula). |
| 5 | `xl-marinade[calc]` is not installed. |
| 6 | The workbook is refused outright (1904 dates, precision as displayed, Strict Open XML). |

**Example**

```bash
marinade calc model.xlsx --policies policies.csv --map map.json \
    --outputs 'Summary!B4' --as-of 2026-01-01 -o out/
```

## `marinade values`

Experimental. Write the workbook's stored values as variable/period rows, with status
`stored`, in the same shape `calc` writes. No recalculation; works on the base install.

```bash
marinade values [OPTIONS] WORKBOOK
```

**Options**

| Flag | Description | Default |
|---|---|---|
| `--variables FILE` | `vars.json`. Required. | |
| `--map FILE` | `map.json`; its id field's stored value becomes the policy id. | policy is `stored` |
| `--ir FILE` | Reuse this IR database. | extract afresh |
| `-o`, `--out PATH` | Output directory (`rows.csv` + `run.json`). Required. | |

**Exit codes**

| Code | Meaning |
|---|---|
| 0 | Success. |
| 1 | Error (bad input, unknown sheet, IR from another workbook, an input that is one of the output files, an `-o` that is a file). |

**Example**

```bash
marinade values model.xlsx --variables vars.json -o stored/
```

## See also

- [Recalculating (experimental)](../guide/calc.md) — `calc` and `values`.
- [Quickstart](../quickstart.md) — the extract → document → diff walkthrough.
- [API reference](api.md) — the equivalent Python API.
- [LLM add-on](../llm.md) — what `--enrich` does and how to configure it.
