# Quickstart

The core workflow is three commands: extract a workbook's formula graph, document it,
and diff two versions of it. All three read and write local files; none of them make a
network call unless you explicitly ask for LLM enrichment.

## 1. Extract

```bash
marinade extract book.xlsx -o ir.db
```

Parses `book.xlsx` and writes its entire formula graph — every formula cell, every
cross-sheet reference, every canonical formula family — into the provenance-stamped
SQLite database `ir.db`. `-o`/`--out` defaults to `ir.db` if omitted.

## 2. Document

```bash
marinade document ir.db -o out/
```

Reads the IR database and writes deterministic documentation into `out/`:
`documentation.md` (human-readable) and `model_spec.json` (machine-readable). No
network call is made.

To layer optional LLM-written narrative documentation on top (requires
`xl-marinade[llm]` and an API key — see [LLM add-on](llm.md)):

```bash
marinade document ir.db -o out/ --enrich
```

If the `[llm]` extra isn't installed, or no key is configured, `--enrich` degrades
gracefully to the same deterministic output as step 2 — it never blocks or errors.

## 3. Diff

```bash
marinade diff a.db b.db
```

Compares two extracted databases (e.g. before/after a workbook revision) and emits a
JSON changelist to stdout — every formula added, removed, or changed, by canonical
family where applicable. Use `-o`/`--out` to write the changelist to a file instead of
stdout.

## What's network-free vs. opt-in

| Command | Network calls |
|---|---|
| `marinade extract` | None, unless you pass `--enrich` (opt-in LLM VBA enrichment) |
| `marinade document` | None, unless you pass `--enrich` (opt-in LLM narrative documentation) |
| `marinade diff` | None — always deterministic |

`--enrich` is the only opt-in network call anywhere in the tool. Everything else —
extraction, diffing, and baseline documentation — is 100% deterministic and runs
entirely on your machine.

## Next

- Full flag-by-flag reference: [CLI reference](reference/cli.md)
- Same workflow from Python: [API reference](reference/api.md)
- Setting up `--enrich`: [LLM add-on](llm.md)
