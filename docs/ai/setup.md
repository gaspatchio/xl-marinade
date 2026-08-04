# Plugins

Install the XL Marinade plugin in your editor. This takes 2 minutes and only needs to be done once.

---

## Prerequisites

Before you start, confirm you have:

- **XL Marinade installed** — `uv add xl-marinade` or `pip install xl-marinade`
- **One of these editors:** Claude Code, VS Code with GitHub Copilot, or Cursor

---

## Install the Plugin

=== "Claude Code"

    Run these two commands once in your terminal — the first registers the marketplace, the second installs the plugin and its skill:

    ```
    /plugin marketplace add gaspatchio/xl-marinade
    /plugin install xl-marinade@xl-marinade
    ```

=== "VS Code / Copilot"

    Agent Plugins is a preview feature — enable it with `"chat.plugins.enabled": true` first. Then add the XL Marinade marketplace to your `settings.json`:

    ```json
    {
      "chat.plugins.marketplaces": ["gaspatchio/xl-marinade"]
    }
    ```

    Open the Extensions view, filter with `@agentPlugins`, and install **xl-marinade**.

=== "Cursor"

    Cursor isn't on the plugin marketplace yet, so install the skill directly:

    ```bash
    npx skills add gaspatchio/xl-marinade
    ```

    On Windows, add `--copy` — symlinks need Developer Mode:

    ```bash
    npx skills add gaspatchio/xl-marinade --copy
    ```

    (Cursor only auto-detects `.cursor-plugin/` inside the xl-marinade repository itself; for your own project, use the command above.)

=== "Other agents"

    For any agent that supports the Agent Skills standard:

    ```bash
    npx skills add gaspatchio/xl-marinade
    ```

    On Windows, add `--copy` — symlinks need Developer Mode.

=== "Offline / Firewalled"

    Clone the repository locally. Your editor will auto-detect the plugin directories when you open the project.

---

## Verify It Works

Run a quick check from the terminal:

```bash
marinade --help
```

If you see the `extract`, `document`, and `diff` commands, XL Marinade is installed. Then open your editor and describe a task:

> "I have an Excel workbook. Help me map what drives its outputs with XL Marinade."

If the response reaches for `marinade extract` and queries the `agent_*` views, the plugin is working.

---

## What You Just Installed

| Component | Description |
|-----------|-------------|
| **1 skill** | `using-xl-marinade` — extract a workbook to SQLite, query the `agent_*` views, audit dependencies, and diff versions |
| **How it loads** | The agent invokes the skill when your request matches its description — you don't name it or run commands |

For what the skill covers and example prompts, see [Skills](skills.md).

---

## Next Steps

Hand this to your team. They don't need to run any commands — just open their editor and start working with a workbook.

See [Skills](skills.md) for what to describe.
