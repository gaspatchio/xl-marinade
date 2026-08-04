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

## Set up the plugin (recommended)

!!! tip "For best results, drive XL Marinade with an AI coding tool"
    XL Marinade is built to be operated by an AI agent: extract a workbook, query
    the `agent_*` SQLite views, diff two versions, and generate documentation. The
    plugin equips your editor with the exact CLI commands, the SQLite query contract,
    and the Excel gotchas — so an agent drives the tool correctly instead of guessing
    at the schema.

In Claude Code, register the marketplace and install the plugin and its skill:

```
/plugin marketplace add gaspatchio/xl-marinade
/plugin install xl-marinade@xl-marinade
```

Using VS Code, Cursor, or another agent? The two-minute setup for each is on the
[Plugins](ai/setup.md) page.

## From source with uv

```bash
git clone https://github.com/gaspatchio/xl-marinade
cd xl-marinade
uv sync
uv run marinade --help
```

## Next

Continue to the [Quickstart](quickstart.md).
