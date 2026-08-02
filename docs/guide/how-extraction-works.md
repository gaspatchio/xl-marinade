# How extraction works

`marinade extract` turns a workbook into a queryable database in four conceptual
stages. Nothing here requires opening Excel, and nothing leaves your machine.

```
workbook (.xlsx/.xlsm)
    │
    ▼
cells + formulas          — every cell, every formula, in both A1 and R1C1 form
    │
    ▼
the formula graph         — every reference becomes an edge (same-sheet, cross-sheet,
    │                        external, or range)
    ▼
bindings                  — cells collapsed into labelled model variables
    │
    ▼
SQLite (ir.db)            — provenance-stamped, versioned, queryable
```

## The formula IS the data

Extraction doesn't summarize or interpret the workbook — it records it. Every
formula cell keeps its own formula text (both the A1 form you'd see in Excel and a
position-invariant R1C1 form used to recognize repeated patterns), its value, its
format, and its position. Nothing is computed or approximated; the database is a
faithful, literal transcription of what the workbook contains. If you can see a
formula in Excel, you can find that same formula, verbatim, in the output database.

Repetition is recognized rather than duplicated: a formula that repeats,
structurally identical, down a column or across a block (the common "one row per
policy/asset" pattern) is grouped into a single canonical family with a
deterministic ID, instead of being stored as thousands of unrelated formulas.

## The formula graph

Every reference a formula makes to another cell becomes an edge: same-sheet
references, cross-sheet references, references to external workbooks, and range
references (`SUM(A1:A100)`) are all resolved into the same graph rather than left
as opaque strings. If `[vba]` is installed, macro-driven mutations (paste-special,
`.Value = .Value` overwrites) are captured as edges too, so a value that "changes
by magic" via a macro is still traceable — see [VBA extraction](vba.md).

## Bindings

Individual cell-level edges are precise but not how anyone thinks about a model —
a 10,000-row projection has as many cell nodes as it has rows. Extraction groups
cells into **bindings**: labelled rectangular ranges that correspond to a single
model variable (a whole column, a single input cell, a block of assumptions).
Reasoning about a model then means reasoning about its variables, not its
individual cells. See [Bindings & the graph](bindings.md) for what a binding is
and how bindings relate to each other.

## Determinism

The same workbook always produces the same graph — the same bindings, the same
binding IDs, the same edges. Sheets and cells are processed in a fixed sort order,
and every hash (cell structure, binding identity) is computed from sorted inputs,
so extraction has no run-to-run variation to reconcile away. The output database
also carries its own provenance: a schema version and a SHA-256 hash of the source
workbook are stamped into the database itself, so two databases can be compared
knowing exactly what schema and what workbook produced each one — this is what
makes `marinade diff` possible (see [Diffing workbooks](diffing.md)).

## Network-free

`extract` never makes a network call by default — parsing a workbook, building the
formula graph, and writing the database are all local, offline operations. The only
opt-in exception is `--enrich`, which layers an LLM-assisted pass on top for VBA
references that static analysis can't resolve; see [LLM add-on](../llm.md).

## Querying the result

The output database is meant to be queried, not just archived. A stable set of
`agent_*` views sits on top of the underlying tables and is the intended surface
for that querying — the raw tables underneath are an implementation detail that
can change between schema versions, while the views are the durable contract. The
full set of views and their columns is covered in the schema reference.

## Next

- [Bindings & the graph](bindings.md) — what a binding is and how bindings relate.
- [Diffing workbooks](diffing.md) — comparing two extracted databases.
- [VBA extraction](vba.md) — the optional `[vba]` extra.
- [Quickstart](../quickstart.md) — the extract → document → diff walkthrough.
