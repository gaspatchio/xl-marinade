# VBA extraction

VBA macros are parsed by every install of XL Marinade: a plain `pip install xl-marinade` reads them. Parsing is static -- nothing executes the macro. The VBA project (`xl/vbaProject.bin` in an `.xlsm`) is read with [`olefile`](https://github.com/decalage2/olefile), and its module source is decoded by XL Marinade itself, following Microsoft's published [MS-OVBA](https://learn.microsoft.com/openspecs/office_file_formats/ms-ovba/) format.

Before 0.7.0 this needed the `[vba]` extra, which installed `oletools`. That extra still exists, empty, so existing install commands keep working.

Every database records whether its VBA was read, in `ir_metadata`:

```sql
SELECT value FROM ir_metadata WHERE key = 'vba_extraction';
```

| Value                 | Meaning                                                                                                                                                     |
| --------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `extracted`           | The VBA tables hold the workbook's VBA.                                                                                                                     |
| `no_vba_project`      | The workbook has no VBA project, e.g. any `.xlsx`.                                                                                                          |
| `failed`              | The VBA project could not be read, or storing it raised (see the extraction log). The tables may be empty or partial.                                       |
| `skipped_no_oletools` | Only in databases extracted by 0.5.0–0.6.x without the `[vba]` extra: the workbook **has** a VBA project that was not read. Re-extract with 0.7.0 or later. |

Empty VBA tables mean "no macros" only when the value is `no_vba_project`. The extraction progress line says the same: `VBA: none (no_vba_project)` for a workbook without macros, `VBA: <n> procs, … (extracted)` when it was read. Databases from before this key existed have no row for it.

## What gets captured

For every VBA project embedded in the workbook, extraction captures:

- **Modules** — standard, class, form, and document modules, each with its full source text and a SHA-256 hash of that source.
- **Procedures** — every Sub, Function, and property accessor, with its signature, parameters, body, and a hash of its normalized body (so two procedures that are byte-different but logically identical can be recognized as such). Event handlers (e.g. `Worksheet_Change`) are flagged as such, and conditionally-compiled twins (the `#If Win64 ... #Else ...` pattern used for 32-bit/64-bit API calls) are both kept rather than one overwriting the other.
- **The call graph** — which procedures call which others, including calls made indirectly through the object model (reading or writing a property).
- **Cell references** — which ranges, named ranges, and tables each procedure reads from or writes to.
- **Security findings** — a keyword scan of the project, stored on each module in `vba_modules.security_findings_json` as `{"type", "keyword", "description"}` entries. `AutoExec` marks procedures Office runs on its own (`Workbook_Open`, `Auto_Open`, `Document_Open`, …). `Suspicious` marks calls that reach outside the workbook: running programs (`Shell`, `WScript.Shell`, PowerShell), network access (`XMLHTTP`, `URLDownloadToFile`), writing or deleting files, DLL `Declare`s, the registry, `CreateObject`/`GetObject`, environment variables, keystrokes sent to other programs (`SendKeys`), Excel 4 macros, calls by computed name (`CallByName`), and reading or editing the VBA project itself. A finding flags code to review; it is not a verdict. A model that only reads and writes its own cells has none.

These become nodes and edges in the same graph that cells and bindings live in (see [Bindings & the graph](https://marinade.gaspatchio.dev/guide/bindings/index.md)), so a procedure that writes to a range shows up as a dependency of that range, right alongside the formulas that also feed it.

## Deterministic paste provenance

A common macro pattern in cashflow and pricing workbooks copies a "template" row of formulas into an output block, one row per iteration, as a **paste-as-values** operation (`PasteSpecial Paste:=xlPasteValues`, or the equivalent `.Value = .Value` assignment). The destination cells end up holding constants, not formulas — so the formula graph alone can't trace from the output block back to the template row it was copied from; that lineage only exists in the macro.

Extraction recovers it: static analysis of the VBA source recognizes these paste patterns and synthesizes a bridging edge from the output block back to the template row it was copied from, tagged with the name of the procedure responsible. The result is fully deterministic and requires no execution or network access — the same macro source always produces the same bridging edges — so a value that "changes by magic" via a macro stays traceable to the specific procedure that wrote it.

## Comparing versions

`marinade diff` compares the VBA of two databases module by module and procedure by procedure, including code outside procedures. It reports added, removed, renamed and edited procedures, and separates comment-only edits from logic changes. See [VBA changes](https://marinade.gaspatchio.dev/guide/diffing/#vba-changes) for the change types. Check `vba_extraction` on both databases first: a side extracted by an older version without the `[vba]` extra (`skipped_no_oletools`) makes every module look added or removed.

## Where LLM enrichment fits

Static analysis can't resolve every reference — a range built up from string concatenation or computed at runtime, for example. The optional `--enrich` flag layers an LLM-assisted pass on top to infer those dynamic references. It's strictly additive to the deterministic VBA extraction above, requires the separate `[llm]` extra and an API key, and never runs unless you explicitly ask for it. See [LLM add-on](https://marinade.gaspatchio.dev/llm/index.md).

## Next

- [Bindings & the graph](https://marinade.gaspatchio.dev/guide/bindings/index.md) — the graph VBA nodes and edges join.
- [Install](https://marinade.gaspatchio.dev/install/index.md) — the `[llm]` extra.
- [LLM add-on](https://marinade.gaspatchio.dev/llm/index.md) — opt-in enrichment for dynamic VBA references.
