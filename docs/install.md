# Install

```bash
pip install xl-marinade
```

Requires **Python 3.11+**.

## Optional extras

XL Marinade's core (`extract`, `diff`, deterministic `document`) has no optional
dependencies. Two add-ons are opt-in:

```bash
# Optional LLM-enriched documentation (bring your own API key)
pip install 'xl-marinade[llm]'

# Optional deterministic VBA macro parsing
pip install 'xl-marinade[vba]'

# Both
pip install 'xl-marinade[llm,vba]'
```

- **`[llm]`** — adds the bring-your-own-key LLM documentation tier (`--enrich`,
  `xl_marinade.llm`). See [LLM add-on](llm.md) for setup and configuration.
- **`[vba]`** — adds deterministic parsing of VBA macros, so macro-driven mutations
  show up as edges in the formula graph (see [What it does](index.md#what-it-does)).

Neither extra changes the deterministic behavior of `extract`, `diff`, or plain
`document` — they only add capability on top.

## From source with uv

```bash
git clone https://github.com/gaspatchio/xl-marinade
cd xl-marinade
uv sync
uv run marinade --help
```

## Next

Continue to the [Quickstart](quickstart.md).
