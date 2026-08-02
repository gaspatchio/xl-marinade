# XL Marinade

**Deterministic Excel formula-graph extraction to SQLite — with optional,
bring-your-own-key LLM documentation.**

Excel workbooks that drive real business decisions (actuarial models, pricing tools,
finance workbooks) accumulate thousands of formulas across dozens of sheets, with no
audit trail beyond "open it and trace the references by hand." XL Marinade parses a
workbook once and writes its entire formula graph — every cell, every cross-sheet
reference, every VBA-driven mutation — into a provenance-stamped SQLite database you can
query, diff, and document like any other data asset.

## What it does

- **Formula graph extraction** — every formula cell becomes a node; every reference
  (same-sheet or cross-sheet) becomes an edge, so you can query "what feeds this cell"
  or "what does this cell feed" with SQL instead of `Ctrl+[`.
- **Cross-sheet dependency resolution** — references across sheets and named ranges are
  resolved into the same graph, not left as opaque strings.
- **Canonical formula families** — structurally identical formulas repeated down a
  column or across a block (the common "one row per asset/policy" pattern) are grouped
  into a single family with a deterministic ID, instead of being treated as thousands of
  unrelated formulas.
- **VBA edges** — macro-driven mutations (e.g. paste-special / `.Value = .Value`
  overwrites) are captured as edges in the same graph, so a value that "changes by
  magic" via a macro is still traceable.
- **Provenance-stamped SQLite output** — the output database carries a schema version
  and per-edge provenance, so downstream tooling can rely on it as a versioned contract
  rather than reverse-engineering table shapes.

## Deterministic core, LLM-optional

`extract`, `diff`, and the baseline `document` output are **100% deterministic and
network-free** — the same workbook always produces the same graph, and nothing ever
leaves your machine. An optional bring-your-own-key LLM tier layers narrative
documentation and VBA enrichment on top, and degrades gracefully back to the
deterministic output if no key is configured. See [LLM add-on](llm.md).

## Where to go next

- [Install](install.md) — `pip install xl-marinade`, plus the `[llm]`/`[vba]` extras.
- [Quickstart](quickstart.md) — the extract → document → diff walkthrough.
- [CLI reference](reference/cli.md) — every command, argument, and flag.
- [API reference](reference/api.md) — the public Python API.
- [LLM add-on](llm.md) — bring-your-own-key narrative documentation and VBA enrichment.
