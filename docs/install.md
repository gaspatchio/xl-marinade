# Install

=== "uv"

    ```bash
    uv tool install xl-marinade
    ```

    !!! note "New to uv?"
        [uv](https://docs.astral.sh/uv/) is a fast Python package and project
        manager. Install it once and the command above will work:

        ```bash
        # macOS / Linux
        curl -LsSf https://astral.sh/uv/install.sh | sh
        ```

        ```powershell
        # Windows (PowerShell)
        powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
        ```

        Prefer a Python you already have? `pipx install uv` (or `pip install uv`)
        works too. Full options are in the
        [uv installation guide](https://docs.astral.sh/uv/getting-started/installation/).

=== "pip"

    ```bash
    pip install xl-marinade
    ```

Requires **Python 3.11+**. `uv tool install` puts the `marinade` CLI on your
PATH — run `marinade --help` from anywhere. To use XL Marinade as a library in
your own project instead, `uv add xl-marinade` and run it with `uv run marinade`.

## Optional extras

XL Marinade's core (`extract`, `diff`, deterministic `document`), including VBA
macro parsing, needs no optional dependencies. One add-on is opt-in:

=== "uv"

    ```bash
    # Optional LLM-enriched documentation (bring your own API key)
    uv tool install 'xl-marinade[llm]'
    ```

=== "pip"

    ```bash
    # Optional LLM-enriched documentation (bring your own API key)
    pip install 'xl-marinade[llm]'
    ```

- **`[llm]`** — adds the bring-your-own-key LLM documentation tier (`--enrich`,
  `xl_marinade.llm`). See [LLM add-on](llm.md) for setup and configuration.
- **`[vba]`** — no longer needed, and kept empty so existing install commands
  keep working. Since 0.7.0 every install parses VBA macros: macro-driven
  mutations show up as edges in the formula graph and `marinade diff` reports
  macro changes (see [VBA extraction](guide/vba.md)). Before 0.7.0 this extra
  installed `oletools`.

The extra doesn't change the deterministic behavior of `extract`, `diff`, or plain
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
