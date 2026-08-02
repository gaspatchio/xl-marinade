# Diffing workbooks

`marinade diff` compares two **extracted IR databases** — the SQLite output of
`marinade extract` — not two raw workbook files directly. Extract each version of
the workbook first, then diff the two databases:

```bash
marinade extract before.xlsx -o a.db
marinade extract after.xlsx  -o b.db
marinade diff a.db b.db
```

This is deterministic and network-free: the same pair of databases always
produces the same changelist.

## Why compare databases, not files

Comparing raw formula text between two workbook versions is noisy: inserting a
row shifts every absolute reference below it, so a plain text diff reports
hundreds of "changes" for a single row insertion. `marinade diff` first matches
sheets and rows/columns between the two versions and re-expresses version A's
addresses in version B's coordinate space, so a row insertion is reported once —
as a row insertion — instead of as a formula change on every cell it shifted.

## The changelist

The output is a JSON object with these top-level keys:

| Key | Contents |
|---|---|
| `version` | The changelist format version (currently `"1.0"`), not a workbook version. |
| `schema_version_a` / `schema_version_b` | The IR schema version each input database was extracted with. |
| `workbook_sha256_a` / `workbook_sha256_b` | The provenance hash of each source workbook, carried over from extraction. |
| `root_a` / `root_b` | The extraction root (sheet + range), if either database was a rooted extraction; `null` for a full-workbook extraction. |
| `summary` | Counts of changes by category — bindings added/removed/moved/resized/formula-changed, cell value/format/dtype changes, sheet and row/column edits, and more (see below). |
| `binding_map` | A flat list mapping every binding in A to its match in B (or `null` if added/removed), each with a `diff_state` of `unchanged`, `modified`, `moved`, `removed`, or `added`. |
| `changes` | The ordered list of individual changes that make up the diff. |

Each entry in `changes` is `{"seq": <position>, "type": <change type>, ...}`,
where the remaining fields depend on the change type (e.g. a formula change
carries `old_formula` and `new_formula`; a resize carries old and new shapes).

## Workbook edits vs. IR inference

Not every entry in `changes` represents something a person edited in the
workbook. Some of extraction's output is **inferred** — table-candidate
detection, binding-label evidence, and time-axis annotations are all derived by
the extractor's semantic layer, and can shift between two extractions even when
no cell in the workbook changed at all. Those entries carry `"layer":
"ir_inference"` and are rolled up separately in `summary.ir_inference_changes`,
so a consumer that only cares about actual workbook edits can filter the
inference layer out rather than mistaking it for a genuine change.

## Reference shift vs. logic change

A formula change (`BINDING_FORMULA_CHANGED`) carries a `modification_kind` of
either:

- **`reference_shift`** — only the absolute row/column targets moved (typically
  because rows or columns were inserted elsewhere in the sheet); the underlying
  logic is identical.
- **`logic_change`** — the formula's actual logic changed.

This distinction matters for review: a wall of `reference_shift` entries after a
row insertion is expected noise, while a `logic_change` on a binding you didn't
touch is worth a second look.

## Quoting-insensitive formula comparison

Excel sometimes stores the same cross-sheet reference with or without
single-quotes around the sheet name (`'Sheet1'!A1` vs `Sheet1!A1`) depending on
how the workbook was last saved, even though nothing about the reference changed.
The diff normalizes this away before comparing formulas, so re-saving a workbook
in a different Excel build doesn't manufacture spurious formula-changed entries.

## Next

- [How extraction works](how-extraction-works.md) — what gets diffed.
- [Bindings & the graph](bindings.md) — the unit `binding_map` and most `changes`
  entries are expressed in terms of.
- [Quickstart](../quickstart.md) — the extract → document → diff walkthrough.
- [CLI reference](../reference/cli.md) — the `marinade diff` flag reference.
