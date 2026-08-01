# XL Marinade

[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.11%2B-blue.svg)](pyproject.toml)

**Deterministic Excel formula-graph extraction to SQLite — with optional, bring-your-own-key
LLM documentation.**

Excel workbooks that drive real business decisions (actuarial models, pricing tools,
finance workbooks) accumulate thousands of formulas across dozens of sheets, with no
audit trail beyond "open it and trace the references by hand." XL Marinade parses a
workbook once and writes its entire formula graph — every cell, every cross-sheet
reference, every VBA-driven mutation — into a provenance-stamped SQLite database you can
query, diff, and document like any other data asset.

---

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
- **Deterministic by default, LLM-optional** — extraction, diffing, and baseline
  documentation are 100% deterministic and make no network calls. An optional
  bring-your-own-key LLM tier adds narrative documentation and VBA enrichment on top,
  and degrades gracefully to the deterministic output if no key is configured.

## Install

```bash
pip install xl-marinade
```

For optional LLM-enriched documentation (bring your own API key):

```bash
pip install 'xl-marinade[llm]'
```

Requires Python 3.11+.

## CLI quickstart

```bash
# Extract a workbook's formula graph to a SQLite database
marinade extract book.xlsx -o ir.db

# Generate deterministic documentation (documentation.md + model_spec.json)
marinade document ir.db -o out/

# Same, with optional LLM enrichment (requires xl-marinade[llm] + an API key;
# degrades to deterministic documentation if no key is configured)
marinade document ir.db -o out/ --enrich

# Diff two extracted databases, emitting a JSON changelist
marinade diff a.db b.db
```

Extraction and diffing are always deterministic and network-free. `--enrich` on
`marinade document` is the only opt-in network call in the tool.

## Library API

```python
import xl_marinade

# Deterministic extraction and diff
xl_marinade.extract("book.xlsx", "ir.db")
changelist = xl_marinade.diff("a.db", "b.db")

# Deterministic documentation (no network)
from xl_marinade.docs import document
document("ir.db", "out/")

# Optional LLM-enriched documentation (requires the [llm] extra)
from xl_marinade.llm import document as document_enriched
document_enriched("ir.db", "out/")
```

All library functions raise a typed error hierarchy (`xl_marinade.errors.MarinadeError`
and subclasses) rather than leaking internal exceptions.

## LLM enrichment (optional, bring-your-own-key)

`extract`, `diff`, and `document` are fully deterministic and never touch the
network. The optional `[llm]` add-on layers LLM-generated narrative documentation
(and opt-in VBA enrichment) on top of that deterministic output, using a provider
and key you supply.

**1. Install the extra:**

```bash
pip install 'xl-marinade[llm]'
```

**2. Set your API key** — either variable works (`LLM_API_KEY` takes precedence):

```bash
export OPENAI_API_KEY="sk-..."
```

**3. Enrich** — add `--enrich` on the CLI, or call the `xl_marinade.llm` entry point:

```bash
marinade document ir.db -o out/ --enrich       # LLM-written narrative documentation
marinade extract book.xlsx -o ir.db --enrich   # opt-in LLM VBA enrichment
```

```python
from xl_marinade.llm import document
document("ir.db", "out/")   # uses the configured key
```

With the `[llm]` extra installed but **no key** configured, enrichment degrades to
deterministic documentation — it never raises or blocks. Enrichment is the only
network call in the tool, and your workbook data is sent only to the endpoint you
configure.

### Configuration

All configuration is via environment variables — the key is read at call time and
never stored:

| Variable | Purpose | Default |
|----------|---------|---------|
| `LLM_API_KEY` / `OPENAI_API_KEY` | API key (required to enrich) | — |
| `OPENAI_MODEL` | Model name | `gpt-5.2` |
| `LLM_BASE_URL` | OpenAI-compatible endpoint override | OpenAI's API |
| `LLM_PROVIDER` | Provider id, recorded in the audit log (`openai`, `azure`, `openai_compatible`) | `openai` |

### Azure, local, or proxied models

The add-on speaks to any **OpenAI-compatible** endpoint via `LLM_BASE_URL` — Azure
OpenAI, a local vLLM/Ollama server, or a LiteLLM proxy:

```bash
export LLM_API_KEY="..."
export LLM_BASE_URL="http://localhost:11434/v1"   # e.g. a local Ollama server
export OPENAI_MODEL="llama3.1"
marinade document ir.db -o out/ --enrich
```

## vs. alternatives

Tools like [pycel](https://github.com/dgorissen/pycel),
[xlcalculator](https://github.com/bradbase/xlcalculator),
[formulas](https://github.com/vinci1it2000/formulas), and
[koala](https://github.com/vallettea/koala) focus on *re-executing* Excel formulas in
Python — useful when you want to run a workbook's calculations outside Excel.

XL Marinade solves a different problem: understanding and auditing the formula graph
itself, without executing it. In particular it adds:

- a **cross-sheet dependency graph** as a first-class, queryable artifact — not an
  intermediate structure discarded after evaluation;
- **canonical formula families**, so a block of a thousand structurally identical
  formulas shows up as one family, not a thousand opaque nodes;
- **VBA edges**, capturing macro-driven mutations that pure-formula re-execution tools
  don't see at all;
- a **provenance-stamped SQLite output** designed to be a stable, versioned contract for
  downstream tooling, rather than an in-memory object graph.

## License

Apache-2.0 — see [LICENSE](LICENSE) and [NOTICE](NOTICE). Copyright 2026 Opio Inc. The
deterministic Excel formula-graph extractor at the core of this project was originally
authored by Klaas Stijnen — see [AUTHORS](AUTHORS).
