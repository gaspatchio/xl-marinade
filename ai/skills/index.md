# Skills

When you install the XL Marinade plugin, your editor gains one area of expertise — driving the extractor to map, audit, and diff Excel workbooks. It activates automatically — you just describe what you need.

No commands required

You don't need to remember the skill name or type special commands. Describe your task in plain English. The right expertise activates based on what you're working on.

## What the skill covers

The **`using-xl-marinade`** skill teaches an agent the full workflow: extract a workbook to a SQLite IR, query the stable `agent_*` views, trace dependencies, diff two versions, and verify a workbook against a source model — with the Excel operational gotchas that trip up headless edits.

| Task                              | Try describing...                                                                                      |
| --------------------------------- | ------------------------------------------------------------------------------------------------------ |
| **Map a workbook**                | *"Extract this workbook and show me the model — one row per labelled column with its formula."*        |
| **Trace what drives an output**   | *"What drives the `Total` cell on the Summary sheet?"*                                                 |
| **Audit an existing model**       | *"Audit this human-built workbook — find the interface seam before I change the lookup table."*        |
| **Attribute a numeric change**    | *"This number moved between v1 and v2. Show me which formula change caused it."*                       |
| **Diff two versions**             | *"Diff these two workbook versions and classify every change as a logic change or a reference shift."* |
| **Verify against a source model** | *"Check that this Excel workbook's dependency structure matches the Python model it was built from."*  |
| **Work with VBA**                 | *"This .xlsm has macros — include the VBA procedures in the dependency graph."*                        |

## How it works

The skill includes detailed instructions, the SQLite query contract (which views to use and their columns), a verification method, an adversarial-audit checklist, and the Excel operational hazards. Think of it as reference material embedded in the tooling — it loads based on what you describe.

## What the skill carries

When the skill is invoked you get:

- The CLI workflow — `marinade extract` / `document` / `diff`
- The SQLite query contract — the `agent_*` / `atlas_*` views and how to query them (never the base tables)
- A verification core (structure proof, then value proof) and an adversarial-audit checklist for what passing proofs still miss
- Excel operational hazards — recalc traps, `openpyxl` cached-value stripping, ArrayFormula corruption, single-instance discipline

That detail loads with the skill your description matches — the plugin installs the one skill, not an always-on knowledge base.

## Next Steps

See the [Quickstart](https://marinade.gaspatchio.dev/quickstart/index.md) for the `extract → document → diff` walkthrough you can run yourself, and the [CLI](https://marinade.gaspatchio.dev/reference/cli/index.md) and [SQLite schema](https://marinade.gaspatchio.dev/reference/schema/index.md) references for the exact contract the skill drives.
