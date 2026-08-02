# LLM add-on

`extract`, `diff`, and `document` are fully deterministic and never touch the network.
The optional `[llm]` add-on layers LLM-generated narrative documentation (and an
opt-in VBA enrichment pass) on top of that deterministic output, using a provider and
key you supply. Enrichment is **advisory only**: if it's unavailable, misconfigured, or
errors out, it always falls back to the same deterministic output you'd get without it
— it never blocks or raises.

## 1. Install the extra

```bash
pip install 'xl-marinade[llm]'
```

## 2. Set your API key

Either variable works; `LLM_API_KEY` takes precedence if both are set:

```bash
export OPENAI_API_KEY="sk-..."
```

## 3. Enrich

Add `--enrich` on the CLI, or call the `xl_marinade.llm` entry point directly:

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

## Configuration

All configuration is via environment variables — the key is read at call time
and never stored. Both enrichment paths (`document --enrich` and the opt-in
`extract --enrich` VBA pass) go through the same client seam, so they honour the
same variables:

| Variable | Purpose | Default |
|---|---|---|
| `LLM_API_KEY` / `OPENAI_API_KEY` | API key (required to enrich) | — |
| `OPENAI_MODEL` | Model name | `gpt-5.2` |
| `LLM_BASE_URL` | OpenAI-compatible endpoint override | OpenAI's API |
| `LLM_PROVIDER` | Provider id, recorded in the audit log (`openai`, `azure`, `openai_compatible`) | `openai` |

!!! note "VBA enrichment uses a cheaper default model"
    `marinade extract --enrich` honours the same variables above — including
    `LLM_BASE_URL`, so it stays on your configured endpoint. The one difference:
    when `OPENAI_MODEL` is unset it defaults to a cheaper model than
    `document --enrich`, because it makes one call per VBA procedure. Set
    `OPENAI_MODEL` to pin a specific model for both.

## Azure, local, or proxied models

The add-on speaks to any **OpenAI-compatible** endpoint via `LLM_BASE_URL` — Azure
OpenAI, a local vLLM/Ollama server, or a LiteLLM proxy — for both `document --enrich`
and `extract --enrich`:

```bash
export LLM_API_KEY="..."
export LLM_BASE_URL="http://localhost:11434/v1"   # e.g. a local Ollama server
export OPENAI_MODEL="llama3.1"
marinade document ir.db -o out/ --enrich
```

## Failure mode is always "fall back," never "break"

Enrichment sits strictly on top of the deterministic core:

- No `[llm]` extra installed → `--enrich` degrades to deterministic documentation.
- Extra installed, no key configured → same graceful degradation.
- Key configured but the request errors (bad endpoint, rate limit, network failure) →
  same graceful degradation.

In every case you still get the deterministic `documentation.md` and `model_spec.json`
that `marinade document` produces without `--enrich` — the LLM tier can only add to
that output, never replace it with something worse or block the run.

## Next

- [Quickstart](quickstart.md) — where `--enrich` fits in the extract → document → diff
  workflow.
- [CLI reference](reference/cli.md) — the full `--enrich` flag reference for `extract`
  and `document`.
