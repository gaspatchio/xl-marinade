# CLI reference

```bash
marinade --help
```

```text
Usage: marinade [OPTIONS] COMMAND [ARGS]...

XL Marinade — deterministic Excel formula-graph extractor.

Commands:
  extract   Extract a workbook's formula graph to a SQLite database.
  document  Generate documentation for an extracted IR database.
  diff      Diff two IR databases; emit the changelist as JSON.
```

Three commands. `extract` and `diff` are always deterministic and network-free; `document` is deterministic by default and only makes a network call if you pass `--enrich` (see [LLM add-on](https://marinade.gaspatchio.dev/llm/index.md)).

## `marinade extract`

Extract a workbook's formula graph to a SQLite database.

```bash
marinade extract [OPTIONS] WORKBOOK
```

**Argument**

| Name       | Description                                 |
| ---------- | ------------------------------------------- |
| `WORKBOOK` | Excel workbook (`.xlsx`/`.xlsm`). Required. |

**Options**

| Flag               | Description                                                                   | Default |
| ------------------ | ----------------------------------------------------------------------------- | ------- |
| `-o`, `--out PATH` | Output SQLite database path.                                                  | `ir.db` |
| `--enrich`         | Opt-in LLM VBA enrichment (makes network calls; requires `xl-marinade[llm]`). | off     |

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

| Name    | Description                                                  |
| ------- | ------------------------------------------------------------ |
| `IR_DB` | IR SQLite database produced by `marinade extract`. Required. |

**Options**

| Flag               | Description                                                                                              | Default    |
| ------------------ | -------------------------------------------------------------------------------------------------------- | ---------- |
| `-o`, `--out PATH` | Output directory (writes `documentation.md` + `model_spec.json`).                                        | `docs_out` |
| `--enrich`         | Opt-in LLM enrichment (requires `xl-marinade[llm]` + a key; degrades to deterministic output otherwise). | off        |

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

| Name   | Description                                |
| ------ | ------------------------------------------ |
| `DB_A` | Version A (earlier) IR database. Required. |
| `DB_B` | Version B (later) IR database. Required.   |

**Options**

| Flag               | Description                     | Default |
| ------------------ | ------------------------------- | ------- |
| `-o`, `--out PATH` | Write the JSON changelist here. | stdout  |

**Example**

```bash
marinade diff a.db b.db
```

## See also

- [Quickstart](https://marinade.gaspatchio.dev/quickstart/index.md) — the extract → document → diff walkthrough.
- [API reference](https://marinade.gaspatchio.dev/reference/api/index.md) — the equivalent Python API.
- [LLM add-on](https://marinade.gaspatchio.dev/llm/index.md) — what `--enrich` does and how to configure it.
