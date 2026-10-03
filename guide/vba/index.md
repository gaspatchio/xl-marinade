# VBA extraction

Deterministic parsing of VBA macros is an opt-in extra:

```bash
pip install 'xl-marinade[vba]'
```

It's backed by [`oletools`](https://github.com/decalage2/oletools), a static-analysis library for Office file formats — nothing in the `[vba]` path executes the macro. Without the extra installed, `extract` still runs normally; it simply produces no VBA nodes or edges.

## What gets captured

For every VBA project embedded in the workbook, extraction captures:

- **Modules** — standard, class, form, and document modules, each with its full source text and a SHA-256 hash of that source.
- **Procedures** — every Sub, Function, and property accessor, with its signature, parameters, body, and a hash of its normalized body (so two procedures that are byte-different but logically identical can be recognized as such). Event handlers (e.g. `Worksheet_Change`) are flagged as such, and conditionally-compiled twins (the `#If Win64 ... #Else ...` pattern used for 32-bit/64-bit API calls) are both kept rather than one overwriting the other.
- **The call graph** — which procedures call which others, including calls made indirectly through the object model (reading or writing a property).
- **Cell references** — which ranges, named ranges, and tables each procedure reads from or writes to.

These become nodes and edges in the same graph that cells and bindings live in (see [Bindings & the graph](https://marinade.gaspatchio.dev/guide/bindings/index.md)), so a procedure that writes to a range shows up as a dependency of that range, right alongside the formulas that also feed it.

## Deterministic paste provenance

A common macro pattern in cashflow and pricing workbooks copies a "template" row of formulas into an output block, one row per iteration, as a **paste-as-values** operation (`PasteSpecial Paste:=xlPasteValues`, or the equivalent `.Value = .Value` assignment). The destination cells end up holding constants, not formulas — so the formula graph alone can't trace from the output block back to the template row it was copied from; that lineage only exists in the macro.

Extraction recovers it: static analysis of the VBA source recognizes these paste patterns and synthesizes a bridging edge from the output block back to the template row it was copied from, tagged with the name of the procedure responsible. The result is fully deterministic and requires no execution or network access — the same macro source always produces the same bridging edges — so a value that "changes by magic" via a macro stays traceable to the specific procedure that wrote it.

## Where LLM enrichment fits

Static analysis can't resolve every reference — a range built up from string concatenation or computed at runtime, for example. The optional `--enrich` flag layers an LLM-assisted pass on top to infer those dynamic references. It's strictly additive to the deterministic VBA extraction above, requires the separate `[llm]` extra and an API key, and never runs unless you explicitly ask for it. See [LLM add-on](https://marinade.gaspatchio.dev/llm/index.md).

## Next

- [Bindings & the graph](https://marinade.gaspatchio.dev/guide/bindings/index.md) — the graph VBA nodes and edges join.
- [Install](https://marinade.gaspatchio.dev/install/index.md) — the `[vba]` extra alongside `[llm]`.
- [LLM add-on](https://marinade.gaspatchio.dev/llm/index.md) — opt-in enrichment for dynamic VBA references.
