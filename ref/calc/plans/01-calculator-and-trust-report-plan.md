# `marinade calc` PR 1: Calculator and Trust Report — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship `marinade calc` (recalculate a workbook for new input sets with IronCalc, with a trust report and a status on every output value) and `marinade values` (stored values as variable/period rows, no IronCalc).

**Architecture:** A new package `xl_marinade.calc` that depends on `xl_marinade.core` and never the reverse. Only `calc/engine.py` imports `ironcalc`, lazily, so `import xl_marinade.calc` and `marinade values` work on the base install. The extraction IR supplies the graph, stored values, formula families, names and validations. IronCalc computes everything, including the probe cells that report branches, lookup positions, rounding inputs, comparisons and dynamic references.

**Tech Stack:** Python 3.11+, `ironcalc` 0.9.1 (optional extra `[calc]`), openpyxl (tests only), Typer CLI, SQLite IR, pytest.

**Spec:** `ref/calc/specs/01-calculator-and-trust-report-design.md` — read it before any task. Background: `ref/calc/README.md`.

## Global Constraints

- Tier boundary: `xl_marinade.core` and `xl_marinade.docs` must not import `xl_marinade.calc` or `ironcalc`. `import xl_marinade.calc` must succeed without `ironcalc`.
- `ironcalc` lives only in the `[calc]` extra: `calc = ["ironcalc==0.9.1"]`. Pin exactly; the engine is pre-1.0.
- The feature is experimental and outside SemVer. Every doc and `--help` text says "experimental".
- `calc` never writes to the extraction database. A temporary IR goes in a temporary directory that is deleted after the run.
- Offline, deterministic: two runs with the same inputs produce byte-identical files. No timestamps in any output.
- Default tolerance `1e-9`, relative or absolute. Near-flip tolerance `1e-12` relative (`max(1, |x|)` scale).
- Probe sheet name: `_marinade_probes`, hidden. Default probe cap: `50_000`.
- Whole-workbook refusal: `WorkbookRefused`, exit code 6, writes only `trust_report.json` and `run.json`.
- `ironcalc` missing: `CalcUnavailable`, exit code 5, message names `pip install 'xl-marinade[calc]'`.
- CSV: UTF-8, `\n` line endings, read with `utf-8-sig`. Floats written with `repr()`.
- Python standards from `AGENTS.md`: full type hints, `@dataclass(slots=True)`, no `print` outside CLI, `loguru` for logging, `pathlib`, typed errors from `errors.py`, early returns.
- Tests use synthetic workbooks only. Never a teaching or client workbook.
- Commits: signed, conventional, no AI signature or `Co-Authored-By` trailer.

## Review Focus

1. **A `policies.csv` saved by Excel starts with a UTF-8 BOM.** The id column must still be found. Test in Task 2.
2. **Sheet names that need quoting** (spaces, apostrophes, a trailing space like `'Life Table '`) in `map.json`, `--outputs`, `vars.json` and in lifted probe formulas. Test in Task 2 and Task 8.
3. **Stored values that are errors, booleans or text** (`#N/A` stored and computed must count as reproduced; `TRUE` is not `1`). Test in Task 6.
4. **Policy ids that look numeric** (`"007"`) stay text; duplicate ids are an input error. Test in Task 2.
5. **Reruns are byte-identical**, including ordering of rows and reasons. Test in Task 11.

---

## File structure

| File | Responsibility |
|---|---|
| `src/xl_marinade/errors.py` | add `CalcError`, `CalcInputError`, `CalcUnavailable`, `WorkbookRefused` |
| `src/xl_marinade/calc/__init__.py` | exports `calc`, `values`, `CalcRequest`, `ValuesRequest` |
| `src/xl_marinade/calc/models.py` | dataclasses and enums shared by every module |
| `src/xl_marinade/calc/config.py` | load and validate `map.json`, `policies.csv`, `vars.json`, `check.json`, `--outputs` |
| `src/xl_marinade/calc/settings.py` | read workbook settings from the xlsx zip |
| `src/xl_marinade/calc/graph.py` | `WorkbookGraph` built from the IR |
| `src/xl_marinade/calc/rows.py` | variable map → rows |
| `src/xl_marinade/calc/values.py` | `values()` |
| `src/xl_marinade/calc/prepare.py` | the engine copy: strip caches, remove data tables, `--as-of`, external refs |
| `src/xl_marinade/calc/engine.py` | the only `ironcalc` import |
| `src/xl_marinade/calc/patterns.py` | static refused-pattern scan |
| `src/xl_marinade/calc/trust.py` | compare engine vs stored |
| `src/xl_marinade/calc/probes.py` | plan, place and read probe cells |
| `src/xl_marinade/calc/status.py` | propagate reasons to outputs |
| `src/xl_marinade/calc/output.py` | write CSV/JSON files |
| `src/xl_marinade/calc/run.py` | `calc()` orchestration, policy loop, `--check-against` |
| `src/xl_marinade/cli/main.py` | `calc` and `values` commands |
| `tests/calc/conftest.py` | `make_book` fixture: synthetic xlsx with cached values; `ir_for` fixture |
| `tests/calc/test_*.py` | one file per module |

---

### Task 1: Package skeleton, errors, extra, test-book fixture

**Files:**
- Modify: `src/xl_marinade/errors.py`, `pyproject.toml`, `uv.lock`, `.github/workflows/licence-check.yml`
- Create: `src/xl_marinade/calc/__init__.py`, `src/xl_marinade/calc/models.py`, `src/xl_marinade/calc/engine.py` (stub `Engine.load` only)
- Test: `tests/calc/conftest.py`, `tests/calc/test_boundary.py`, `tests/calc/test_book_fixture.py`

**Interfaces:**
- Produces:
  - `errors.CalcError(MarinadeError)`, `CalcInputError(CalcError)`, `CalcUnavailable(CalcError)`, `WorkbookRefused(CalcError)` with attribute `reasons: list[str]`.
  - `models.CellRef(NamedTuple)`: `sheet: str, row: int, col: int`; method `a1() -> str` returns `Sheet!A1` with the sheet quoted via `core.ref_converter.quote_sheet_name`; classmethod `parse(text: str) -> CellRef` (raises `CalcInputError`).
  - `models.Status(StrEnum)`: `REFUSED="refused"`, `UNCERTAIN="uncertain"`, `UNVERIFIED="unverified_path"`, `OK="ok"`, `STORED="stored"`.
  - `models.ReasonKind(StrEnum)`: `silently_wrong, engine_error, macro_function, unsupported_function, no_stored_value, refused_pattern, check_against_mismatch, near_rounding_flip, near_comparison, unexercised_branch, unexercised_lookup_edge, not_checked_path, not_checked_value, external_link_held, input_outside_allowed_values`.
  - `models.Reason` (slots dataclass, frozen, ordered): `kind: ReasonKind, root: CellRef | None, detail: str`.
  - `models.SEVERITY: dict[ReasonKind, Status]`: refused ← silently_wrong, engine_error, macro_function, unsupported_function, no_stored_value, refused_pattern, check_against_mismatch, input_outside_allowed_values; uncertain ← near_rounding_flip, near_comparison, not_checked_value; unverified ← unexercised_branch, unexercised_lookup_edge, not_checked_path; ok ← external_link_held.
  - `engine.Engine.load(path: Path) -> Engine` raising `CalcUnavailable` when `ironcalc` is not importable.
  - Fixture `make_book(name: str, sheets: dict[str, dict[str, object]], *, names: dict[str, str] | None = None, date1904: bool = False, full_precision: bool = True, iterate: bool = False, manual: bool = False, validations: list[tuple[str, list[str]]] | None = None) -> Path`. A cell value is a constant, or `F(formula: str, cached: object)` (exported from `tests/calc/conftest.py`). The fixture saves with openpyxl, then patches each sheet XML in the zip to add the cached `<v>` (and `t="str"`, `t="b"`, `t="e"` as the type needs), so the file looks like one Excel saved.
  - Fixture `ir_for(path: Path) -> Path` runs `xl_marinade.extract` into `tmp_path`.

- [ ] **Step 1: Write the failing tests**

```python
# tests/calc/test_boundary.py
def test_calc_imports_without_ironcalc(blocked_ironcalc):
    import xl_marinade.calc  # noqa: F401

def test_engine_load_without_ironcalc_raises(blocked_ironcalc, tmp_path):
    from xl_marinade.calc.engine import Engine
    from xl_marinade.errors import CalcUnavailable
    with pytest.raises(CalcUnavailable, match=r"xl-marinade\[calc\]"):
        Engine.load(tmp_path / "x.xlsx")

def test_core_and_docs_do_not_import_calc():
    # run in a subprocess: python -c "import xl_marinade, xl_marinade.docs; import sys;
    # assert not any(m.startswith(('xl_marinade.calc','ironcalc')) for m in sys.modules)"
    ...

def test_cellref_quotes_sheet_with_trailing_space():
    assert CellRef("Life Table ", 9, 1).a1() == "'Life Table '!A9"
    assert CellRef.parse("'It''s'!$B$3") == CellRef("It's", 3, 2)
```

`blocked_ironcalc` follows the `blocked_llm` fixture in `tests/test_core_no_llm_dep.py`.

```python
# tests/calc/test_book_fixture.py
def test_make_book_writes_cached_values(make_book):
    p = make_book("b", {"S": {"A1": 2, "B1": F("=A1*3", 6), "C1": F('=IF(A1>1,"hi","lo")', "hi"),
                              "D1": F("=NA()", "#N/A"), "E1": F("=A1>1", True)}})
    wv = openpyxl.load_workbook(p, data_only=True)["S"]
    assert (wv["B1"].value, wv["C1"].value, wv["D1"].value, wv["E1"].value) == (6, "hi", "#N/A", True)
```

- [ ] **Step 2: Run, verify they fail** — `uv run pytest tests/calc -q` → import errors.
- [ ] **Step 3: Implement** the errors, `models.py`, the `Engine.load` stub, the fixtures. Add `calc = ["ironcalc==0.9.1"]` to `[project.optional-dependencies]`; `uv lock`. In `licence-check.yml` change `uv sync --frozen` to `uv sync --frozen --all-extras` and add `--all-extras` to the `uv pip compile` call, so the audit covers `[calc]`.
- [ ] **Step 4: Run** `uv run pytest tests/calc tests/test_import_cleanliness.py tests/test_core_no_llm_dep.py -q` → PASS. `uv sync --extra calc` installs ironcalc.
- [ ] **Step 5: Commit** — `feat(calc): package skeleton, errors and the [calc] extra`

---

### Task 2: Input files

**Files:**
- Create: `src/xl_marinade/calc/config.py`
- Test: `tests/calc/test_config.py`

**Interfaces:**
- Consumes: `CellRef`, `CalcInputError`.
- Produces (all in `models.py`, loaders in `config.py`):
  - `InputMap(fields: dict[str, str], id_field: str | None)` — values are the raw targets (`Sheet!A1` or a defined name); resolved later by `WorkbookGraph.resolve_target`. `load_input_map(path: Path) -> InputMap`. File shape: `{"id_field": "policy_id", "fields": {"age": "Inputs!B3", "sa": "SumAssured"}}`; `id_field` optional and must not be in `fields`.
  - `Policy(id: str, values: dict[str, str])`; `load_policies(path: Path, input_map: InputMap) -> list[Policy]`. The id column is `input_map.id_field` or `"id"`. Read with `utf-8-sig`. Errors: missing id column, duplicate id, a mapped field missing from the header, a blank value in a mapped field (message names policy id and field).
  - `VariableSpec(name: str, target: str, period_start: int | None)`; `load_variables(path: Path) -> list[VariableSpec]`. Range entries need `period_start` (int ≥ 0); cell entries must not have it. Order is file order.
  - `CheckAgainst(results_sheet: str, id_column: str, first_row: int, outputs: dict[str, str], table: InputTable | None, selector: str | None)`; `InputTable(sheet: str, id_column: str, first_row: int, fields: dict[str, str])`. Exactly one of `table`/`selector`. `load_check_against(path: Path) -> CheckAgainst`.
  - `parse_outputs(values: list[str]) -> list[str]` — keeps order, drops exact duplicates.

- [ ] **Step 1: Failing tests**

```python
def test_policies_bom_header(tmp_path):
    p = tmp_path / "p.csv"; p.write_bytes("﻿id,age\n007,40\n".encode())
    pols = load_policies(p, InputMap(fields={"age": "S!B1"}, id_field=None))
    assert pols == [Policy(id="007", values={"age": "40"})]

def test_duplicate_policy_id_is_error(tmp_path): ...   # CalcInputError, message contains "duplicate" and "007"
def test_blank_mapped_value_is_error(tmp_path): ...    # CalcInputError naming policy and field
def test_range_variable_needs_period_start(tmp_path):
    # {"premium": {"range": "Proj!C11:C20"}} -> CalcInputError match "period_start"
def test_cell_variable_rejects_period_start(tmp_path): ...
def test_check_against_needs_exactly_one_input_source(tmp_path): ...
def test_quoted_sheet_targets_survive(tmp_path):
    # map {"fields": {"age": "'Life Table '!B3"}} -> fields["age"] == "'Life Table '!B3"
```

- [ ] **Step 2: Run, verify fail** — `uv run pytest tests/calc/test_config.py -q`
- [ ] **Step 3: Implement** the loaders. JSON via `json.loads`; validate with early returns and `CalcInputError`.
- [ ] **Step 4: Run, verify pass**
- [ ] **Step 5: Commit** — `feat(calc): load and validate calc input files`

---

### Task 3: Workbook settings

**Files:**
- Create: `src/xl_marinade/calc/settings.py`
- Test: `tests/calc/test_settings.py`

**Interfaces:**
- Produces: `WorkbookSettings(date1904: bool, full_precision: bool, iterate: bool, manual_calc: bool)`; `read_settings(path: Path) -> WorkbookSettings` (reads `workbookPr@date1904` and `calcPr@fullPrecision|iterate|calcMode` from `xl/workbook.xml`); `refusal_reasons(s: WorkbookSettings) -> list[str]` returning `["1904 date system"]` and/or `["precision as displayed"]`, empty otherwise.

- [ ] **Step 1: Failing tests** — one per flag via `make_book(..., date1904=True)` etc.: `read_settings(p).date1904 is True`; `refusal_reasons(...) == ["1904 date system"]`; a default book gives all False and `[]`; `manual=True` gives `manual_calc=True` and `refusal_reasons == []`.
- [ ] **Step 2: Run, verify fail**
- [ ] **Step 3: Implement** with `zipfile` + `xml.etree.ElementTree`, namespace-aware. Absent attributes mean Excel defaults (`fullPrecision` default true).
- [ ] **Step 4: Run, verify pass**
- [ ] **Step 5: Commit** — `feat(calc): read workbook calculation settings`

---

### Task 4: Graph from the IR

**Files:**
- Create: `src/xl_marinade/calc/graph.py`
- Test: `tests/calc/test_graph.py`

**Interfaces:**
- Consumes: IR tables `cells`, `formulas`, `json_blobs`, `sheets`, `cell_edges_internal`, `range_edges`, `cell_edges_external`, `cell_udf_calls`, `defined_names`, `data_validations`, `ir_metadata`.
- Produces: `StoredValue = float | str | bool | None` (errors as their `#CODE` string); class `WorkbookGraph` with
  - `from_ir(path: Path) -> WorkbookGraph`
  - `workbook_sha256: str`
  - `formula(cell: CellRef) -> str | None` (A1 text without `=`)
  - `formula_cells() -> list[CellRef]` (sorted)
  - `stored(cell: CellRef) -> StoredValue` and `has_stored(cell: CellRef) -> bool` — convert the IR's raw text by `data_type` (`s`/`str` text, `b` bool, `e` error, `None` number via `float`)
  - `family(cell: CellRef) -> str` = `f"{sheet}|{formula_r1c1}"`
  - `precedents(cell: CellRef) -> frozenset[CellRef]` — direct cell edges plus every cell present in the IR inside each range edge
  - `downstream(roots: Iterable[CellRef]) -> dict[CellRef, CellRef]` — every cell reachable forward from the roots (roots included), mapped to the smallest root that reaches it (deterministic)
  - `resolve_target(target: str) -> list[CellRef]` — an A1 cell/range or a workbook-scoped defined name; raises `CalcInputError` if unknown
  - `upstream(cells: Iterable[CellRef]) -> set[CellRef]` — backward closure through `precedents` (cells included)
  - `udf_cells: frozenset[CellRef]`, `external_cells: dict[CellRef, list[str]]`
  - `allowed_values(cell: CellRef) -> list[str] | None` — list validations: literal values, or the stored values of a range source; `None` when the cell has no list validation
  - `is_sheet_scoped_name(name: str) -> bool`, `name_cells(name: str) -> list[CellRef]`

- [ ] **Step 1: Failing tests** on a book `S: A1=2, A2=3, B1=F("=SUM(A1:A2)",5), C1=F("=B1*2",10), D1=F("=C1+A1",12)`:
  - `g.precedents(CellRef("S",1,2)) == {CellRef("S",1,1), CellRef("S",2,1)}`
  - `g.downstream([CellRef("S",1,2)]) == {B1: B1, C1: B1, D1: B1}` (as CellRefs)
  - `g.stored(CellRef("S",1,3)) == 10.0`; a stored `#N/A` returns `"#N/A"`; a stored `TRUE` returns `True`
  - `g.family` equal for two cells holding the same copied formula on one sheet
  - `g.resolve_target("rate")` for `names={"rate": "S!$A$1"}` → `[CellRef("S",1,1)]`; unknown name raises `CalcInputError`
  - `g.allowed_values` for a list validation `["M","F"]` on `S!E1`
- [ ] **Step 2: Run, verify fail**
- [ ] **Step 3: Implement** with `sqlite3` read-only (`file:...?mode=ro`, see `core/db_uri.py`). Build a per-sheet interval index of range edges for `downstream`.
- [ ] **Step 4: Run, verify pass**
- [ ] **Step 5: Commit** — `feat(calc): load the formula graph and stored values from the IR`

---

### Task 5: Variable rows and `marinade values`

**Files:**
- Create: `src/xl_marinade/calc/rows.py`, `src/xl_marinade/calc/values.py`, `src/xl_marinade/calc/output.py` (row writer only)
- Modify: `src/xl_marinade/calc/__init__.py`, `src/xl_marinade/cli/main.py`
- Test: `tests/calc/test_values.py`, `tests/calc/test_cli_values.py`

**Interfaces:**
- Consumes: `VariableSpec`, `WorkbookGraph`, `InputMap`.
- Produces:
  - `VariableCells(name: str, cells: list[CellRef], period_start: int | None)`; `expand_variables(specs: list[VariableSpec], graph: WorkbookGraph) -> list[VariableCells]` — a range must be one row or one column (`CalcInputError` "2-D range" otherwise); runs top-to-bottom or left-to-right.
  - `Row(policy: str, variable: str, period: int | None, value: StoredValue, status: Status, reasons: list[Reason])`; `rows_for(policy: str, variables: list[VariableCells], value_of: Callable[[CellRef], tuple[StoredValue, Status, list[Reason]]]) -> list[Row]`.
  - `output.write_rows(path: Path, rows: list[Row]) -> None` — header `policy,variable,period,value,status,reasons`; empty period for single cells; reasons as `kind@root:detail` joined by `; `, sorted.
  - `ValuesRequest(workbook: Path, variables: Path, out_dir: Path, input_map: Path | None = None, ir: Path | None = None)`; `values(req: ValuesRequest) -> Path` (returns `rows.csv`). Stored policy id: the cell of `id_field` if mapped in `input_map`, else `"stored"`. Also writes `run.json` with, per variable, `range`, `period_start`, `period_0_cell`.
  - IR reuse helper in `graph.py`: `open_graph(workbook: Path, ir: Path | None, tmp: Path) -> WorkbookGraph` — with `ir`, check `sha256(workbook) == graph.workbook_sha256` else `CalcInputError`; without, extract into `tmp`.
  - CLI `marinade values WORKBOOK --variables PATH [--map PATH] [--ir PATH] -o DIR` (exit 1 on `MarinadeError`).

- [ ] **Step 1: Failing tests**

```python
def test_values_rows(make_book, tmp_path):
    p = make_book("v", {"Proj": {"C11": F("=1", 1), "C12": F("=2", 2), "C13": F("=3", 3)},
                        "Summary": {"B4": F("=6", 6)}})
    vars_ = write_json(tmp_path, {"premium": {"range": "Proj!C11:C13", "period_start": 1},
                                   "npv": {"cell": "Summary!B4"}})
    out = values(ValuesRequest(p, vars_, tmp_path / "o"))
    assert out.read_text().splitlines() == [
        "policy,variable,period,value,status,reasons",
        "stored,premium,1,1.0,stored,", "stored,premium,2,2.0,stored,", "stored,premium,3,3.0,stored,",
        "stored,npv,,6.0,stored,"]
    run = json.loads((tmp_path / "o" / "run.json").read_text())
    assert run["variables"]["premium"] == {"range": "Proj!C11:C13", "period_start": 1, "period_0_cell": "Proj!C10"}

def test_values_2d_range_rejected(...): ...      # CalcInputError "2-D range"
def test_values_ir_hash_mismatch(...): ...       # passing an IR of another book -> CalcInputError
def test_values_needs_no_ironcalc(blocked_ironcalc, ...): ...
def test_cli_values_exit_codes(...): ...         # 0 on success, 1 on bad vars file
```

`period_0_cell` is the cell one step before the first cell when `period_start` is 1 (it may sit outside the range; report it anyway).

- [ ] **Step 2: Run, verify fail**
- [ ] **Step 3: Implement**
- [ ] **Step 4: Run, verify pass**
- [ ] **Step 5: Commit** — `feat(calc): marinade values writes stored values as variable/period rows`

---

### Task 6: Engine copy, engine wrapper, trust run

**Files:**
- Create: `src/xl_marinade/calc/prepare.py`, `src/xl_marinade/calc/trust.py`
- Modify: `src/xl_marinade/calc/engine.py`
- Test: `tests/calc/test_prepare.py`, `tests/calc/test_engine.py`, `tests/calc/test_trust.py` (all `pytest.importorskip("ironcalc")`)

**Interfaces:**
- Produces:
  - `PreparedCopy(path: Path, stripped: int, data_table_cells: list[CellRef], held_externals: dict[CellRef, list[str]])`; `prepare_engine_copy(src: Path, dst: Path, *, as_of: date | None) -> PreparedCopy`. It deletes every formula cell's cached `<v>` (logic of `ref/calc/comparison/common.py:strip_cached`), removes `<f t="dataTable" ...>` cells' formulas (keeping no value), rewrites `TODAY()`/`NOW()` outside string literals to `DATE(y,m,d)` when `as_of` is set, and replaces each external reference `[n]Sheet!Ref` in formula text by the literal cached in `xl/externalLinks/externalLink{n}.xml` (`sheetDataSet`). An external reference with no cached value is left as is (IronCalc gives `#REF!`, which the trust run reports).
  - `EngineValue = float | str | bool | None`. `Engine`: `load(path) -> Engine`, `set_value(cell: CellRef, value: float | str | bool) -> None` (via `update_cell_with_number|text|bool`), `set_formula(cell: CellRef, formula: str) -> None`, `add_hidden_sheet(name: str) -> None`, `evaluate() -> None`, `value(cell: CellRef) -> EngineValue`, `version: str`.
  - `TrustFailure(cell: CellRef, reason: Reason)`; `TrustResult(checked: int, reproduced: int, failures: list[TrustFailure])` with `share` property; `same(stored: StoredValue, got: EngineValue, tol: float) -> bool`; `run_trust(engine: Engine, graph: WorkbookGraph, *, tolerance: float, skip: set[CellRef]) -> TrustResult`. Reason precedence for a mismatch: `no_stored_value` (no `<v>` in the original) → `macro_function` (cell in `udf_cells` and got `#NAME?`) → `unsupported_function` (got `#NAME?`) → `engine_error` (got `#...`, stored was not an error) → `silently_wrong`. `skip` holds refused-pattern cells (not compared).

- [ ] **Step 1: Failing tests**

```python
def test_same():
    assert same(1.0, 1.0 + 1e-12, 1e-9) and not same(1.0, 1.001, 1e-9)
    assert same("#N/A", "#N/A", 1e-9) and not same(True, 1.0, 1e-9) and same("", None, 1e-9)

def test_trust_reproduces_simple_book(make_book, ir_for, tmp_path):  # share == 1.0
def test_trust_flags_silently_wrong(make_book, ...):
    # S: A1=" " text, B1..B5 numbers, C1=F('=VLOOKUP(A1,B1:C5,2)', "#N/A") -> reason silently_wrong
def test_trust_flags_engine_error_sheet_qualified_range(...):  # F("=SUM(S!$A$1:S!$A3)", 6) -> engine_error
def test_prepare_as_of(...):          # F("=TODAY()", 46023) with as_of=date(2026,1,1) -> reproduced
def test_prepare_data_table_removed(...):  # a sheet XML with <f t="dataTable" ref="B2:B3" .../> loads; cells listed
def test_prepare_external_held(...):  # patched zip with externalLink1.xml cache -> value equals cache
def test_stored_missing_is_no_stored_value(...)
```

The data-table and external-link books are built by patching XML in the test (openpyxl cannot write them); keep the helper in `tests/calc/conftest.py`.

- [ ] **Step 2: Run, verify fail**
- [ ] **Step 3: Implement**
- [ ] **Step 4: Run, verify pass** — `uv run --extra calc pytest tests/calc -q`
- [ ] **Step 5: Commit** — `feat(calc): engine copy, IronCalc wrapper and trust run`

---

### Task 7: Refused patterns

**Files:**
- Create: `src/xl_marinade/calc/patterns.py`
- Test: `tests/calc/test_patterns.py`

**Interfaces:**
- Consumes: `WorkbookGraph`, `core.parser.FormulaParser`, `core.levels_cycles.detect_cycles`.
- Produces: `scan_patterns(graph: WorkbookGraph, *, as_of: date | None, data_table_cells: list[CellRef]) -> dict[CellRef, list[Reason]]`, all with `kind=refused_pattern`, `root=` the cell, `detail` one of:
  - `approximate lookup on unsorted column` / `approximate lookup on mixed text and numbers` / `approximate lookup on a non-range table` — `VLOOKUP`/`HLOOKUP` with 4th arg absent, `TRUE` or non-zero; `MATCH` with 3rd arg absent, `1` or `-1` (`-1` checks descending); `LOOKUP` always. Key column/row = first column (VLOOKUP), first row (HLOOKUP), the lookup array (MATCH), the lookup vector (LOOKUP). Sortedness uses the stored values, ignoring blanks.
  - `implicit intersection on multi-cell name` — a `RefNode` naming a defined name with more than one cell, used anywhere except directly as a function argument.
  - `RAND`, `RANDBETWEEN`; `TODAY`/`NOW` only when `as_of is None`.
  - `circular reference` — every cell in a cycle from `detect_cycles` (cell ids as strings).
  - `data table` — each of `data_table_cells`.
  - `unparseable formula` — the parser raised.

- [ ] **Step 1: Failing tests** — one per detail string, plus negatives: exact-match `VLOOKUP(...,FALSE)` on unsorted keys is not refused; an approximate lookup on a sorted numeric column is not refused; a multi-cell name inside `SUM(x)` is not refused.
- [ ] **Step 2: Run, verify fail**
- [ ] **Step 3: Implement**
- [ ] **Step 4: Run, verify pass**
- [ ] **Step 5: Commit** — `feat(calc): refuse formulas that only match the stored inputs by luck`

---

### Task 8: Probes

**Files:**
- Create: `src/xl_marinade/calc/probes.py`
- Test: `tests/calc/test_probes.py` (`importorskip("ironcalc")`)

**Interfaces:**
- Consumes: `Engine`, `WorkbookGraph`, `FormulaParser`.
- Produces:
  - `ProbeKind(StrEnum)`: `branch, lookup, rounding, comparison, dynamic_ref`.
  - `Probe(kind: ProbeKind, cell: CellRef, family: str, ordinal: int, formulas: tuple[str, ...], meta: dict[str, str])`; `ProbePlan(probes: list[Probe], unplaced: dict[CellRef, list[ProbeKind]])`.
  - `anchor(node: ASTNode, sheet: str) -> str` — formula text with every unqualified reference made sheet-qualified and absolute, so it computes the same on the probe sheet. Raises `ValueError` for a sheet-scoped name, an array constant, `LET`/`LAMBDA`.
  - `plan_probes(graph: WorkbookGraph, cells: Iterable[CellRef], *, cap: int) -> ProbePlan` for these nodes in each cell's formula:
    - `IF(c, ...)`: probe `c` → outcome `TRUE`/`FALSE`/error. `IFS`: index of first true condition. `SWITCH`: index of the matching case or `default`. `CHOOSE`: the index value. `IFERROR(x, ...)`: `ISERROR(x)`; `IFNA`: `ISNA(x)`.
    - Lookups (`VLOOKUP`, `HLOOKUP`, `MATCH`, `LOOKUP`, `INDEX` with a `MATCH` argument): `MATCH(v, keys, mode)` and `ROWS(keys)` (or `COLUMNS`) → edge class `first` / `last` / `interior` / `not_found`.
    - `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `MROUND`, `INT`, `TRUNC`, `CEILING`, `FLOOR`: probe the first argument and, where present, the second.
    - Comparisons `<`, `<=`, `>`, `>=`, `=`, `<>` inside a branch condition, and approximate lookups: probe both sides (lookup: the lookup value and the matched key).
    - `OFFSET`, `INDIRECT`: `CELL("address", ref)`, `ROWS(ref)`, `COLUMNS(ref)`.
    A node that `anchor` cannot lift, or any probe beyond `cap`, goes in `unplaced`.
  - `place_probes(engine: Engine, plan: ProbePlan) -> None` — writes probe formulas to `_marinade_probes`, one per row, in plan order.
  - `ProbeReadings(branch: dict[tuple[CellRef, int], str], lookup_edge: dict[tuple[CellRef, int], str], near_rounding: set[CellRef], near_comparison: set[CellRef], dynamic_refs: dict[CellRef, list[CellRef]])`; `read_probes(engine: Engine, plan: ProbePlan) -> ProbeReadings`.
  - `near_flip(fn: str, x: float, arg: float | None) -> bool` — `ROUND`: `x·10^d` within `1e-12·max(1,|x·10^d|)` of a half-integer; `ROUNDUP`/`ROUNDDOWN`/`TRUNC`: of an integer at `10^d`; `INT`: of an integer; `MROUND`: `x/m` of a half-integer; `CEILING`/`FLOOR`: `x/sig` of an integer. Comparisons: `|a−b| ≤ 1e-12·max(1,|a|,|b|)`.

- [ ] **Step 1: Failing tests**

```python
def test_near_flip():
    assert near_flip("ROUND", 1.005, 2) is True       # 100.49999999999999 is 1.4e-14 from the tie
    assert near_flip("ROUND", 0.125, 2) is True
    assert near_flip("INT", 3.0000000000001, None) is True
    assert near_flip("ROUND", 0.123, 2) is False

def test_anchor_quotes_sheet_and_absolutises():
    assert anchor(FormulaParser().parse("L45>='Proj Sheet'!A1"), "Life Table ") == "'Life Table '!$L$45>='Proj Sheet'!A1"

def test_branch_probe_reads_outcome(...):    # IF(A1>1,..) with A1=2 -> "TRUE"; set A1=0 -> "FALSE"
def test_lookup_edge_classes(...):           # value hitting first, last, interior, beyond end
def test_dynamic_ref_probe(...):             # SUM(OFFSET(A1,C1,0,2,1)), C1=4 -> [A5, A6]
def test_unliftable_goes_unplaced(...):      # LET(...) and cap=0 cases
```

- [ ] **Step 2: Run, verify fail**
- [ ] **Step 3: Implement**
- [ ] **Step 4: Run, verify pass**
- [ ] **Step 5: Commit** — `feat(calc): probe branches, lookup edges, rounding, comparisons and dynamic references`

---

### Task 9: Status assignment

**Files:**
- Create: `src/xl_marinade/calc/status.py`
- Test: `tests/calc/test_status.py` (pure Python; builds `WorkbookGraph` from a small IR)

**Interfaces:**
- Consumes: `WorkbookGraph.downstream`, `Reason`, `SEVERITY`, `ProbeReadings`, `ProbePlan`.
- Produces:
  - `Exercised(branches: set[tuple[str, int, str]], lookup_edges: set[tuple[str, int, str]])` keyed by `(family, ordinal, outcome)`; `exercised_from(plan: ProbePlan, readings: list[ProbeReadings]) -> Exercised`. `interior` counts per family; `first`, `last`, `not_found` each count only if hit.
  - `policy_roots(plan: ProbePlan, readings: ProbeReadings, exercised: Exercised) -> dict[CellRef, list[Reason]]` — unexercised branch/edge, near flips, near comparisons, `not_checked_path` for each unplaced branch or lookup probe and `not_checked_value` for each unplaced rounding or comparison probe.
  - `assign(graph: WorkbookGraph, outputs: list[CellRef], roots: dict[CellRef, list[Reason]], dynamic_refs: dict[CellRef, list[CellRef]]) -> dict[CellRef, tuple[Status, list[Reason]]]` — an output's reasons are those of every root whose `downstream` closure (extended through `dynamic_refs`: a cell depends on the cells its dynamic reference pointed to) contains it; status is the most severe; at most 5 reasons per kind, smallest roots first; `ok` with no reasons otherwise. An unplaced `dynamic_ref` probe is a `refused_pattern` root with detail `dynamic reference not checked`.

- [ ] **Step 1: Failing tests** — chain `A1 → B1 → C1`: a `silently_wrong` root at B1 refuses C1 with root B1; a `near_rounding_flip` at B1 gives C1 `uncertain`; both at once gives `refused` with both reasons; an output not downstream stays `ok`; a dynamic ref from C1 to a failing D9 refuses C1; a branch outcome seen in family member row 40 counts as exercised for row 50 (per-formula rule); a `last` edge not seen in the trust run gives `unverified_path`.
- [ ] **Step 2–4:** run fail, implement, run pass.
- [ ] **Step 5: Commit** — `feat(calc): assign a status to every output from its root causes`

---

### Task 10: The `calc()` run

**Files:**
- Create: `src/xl_marinade/calc/run.py`
- Modify: `src/xl_marinade/calc/output.py`, `src/xl_marinade/calc/__init__.py`
- Test: `tests/calc/test_run.py` (`importorskip("ironcalc")`)

**Interfaces:**
- Consumes: everything above.
- Produces:
  - `CalcRequest(workbook: Path, policies: Path, input_map: Path, out_dir: Path, outputs: list[str] = [], variables: Path | None = None, check_against: Path | None = None, ir: Path | None = None, as_of: date | None = None, tolerance: float = 1e-9, probe_cap: int = 50_000)` (use `field(default_factory=list)`).
  - `CalcResult(results_csv: Path, rows_csv: Path | None, trust_report: Path, paths_csv: Path, run_json: Path)`; `calc(req: CalcRequest) -> CalcResult`.
  - Writers in `output.py`: `write_results(path, values: list[OutputValue])`, `write_trust_report(path, report: dict)`, `write_paths(path, records: list[PathRecord])`, `write_run(path, record: dict)`. `OutputValue(policy_id: str, cell: CellRef, value: EngineValue, status: Status, reasons: list[Reason])`; `PathRecord(policy_id: str, cell: CellRef, family: str, kind: str, outcome: str, exercised: bool)`.
- Order of a run (spec "How a run works"): settings → refuse (write `trust_report.json` + `run.json`, raise `WorkbookRefused`) → graph → patterns → engine copy → load engine → plan and place probes for every formula cell upstream of the requested outputs (`graph` backward closure) → trust run on stored inputs → `--check-against` input sets (set inputs from the table or set the selector; compare each named output with the stored result; a mismatch is a `check_against_mismatch` root on that output cell) → each cell in `PreparedCopy.held_externals` becomes an `external_link_held` root → `Exercised` from all trust readings → for each policy: reset every mapped input to its stored value, check allowed values (outside → every output of that policy gets `input_outside_allowed_values`), coerce and set inputs (number cell: `float`; ISO `YYYY-MM-DD` on a number cell → 1900 serial; bool cell: `TRUE`/`FALSE`; else text), evaluate, read outputs and probes, `assign` → write files.
- Mapping onto a formula cell is a `CalcInputError`. `trust_report.json` keys: `formula_cells`, `reproduced`, `share`, `failures` (list of `{cell, kind, detail}` with `silently_wrong` first then by cell), `refused_patterns`, `settings`, `held_external_links`. `run.json` keys: `ironcalc_version`, `marinade_version`, `workbook_sha256`, `tolerance`, `as_of`, `probe_cap`, `probes_unplaced`, `variables` (as in Task 5).

- [ ] **Step 1: Failing tests** (synthetic book: `Inputs!B1` age, `Calc!B1=F("=IF(Inputs!B1>60,Inputs!B1*2,Inputs!B1)", 40)` with stored age 40, `Calc!B2=F("=ROUND(Calc!B1/8,2)", 5)`):
  - policy age 50 → `Calc!B1` 50.0 `ok`; age 70 → 140.0 `unverified_path` with reason `unexercised_branch`
  - age 41 → `Calc!B2` is `ROUND(41/8,2)` = `ROUND(5.125,2)` → `uncertain`, reason `near_rounding_flip`
  - an output behind a `RAND()` cell → `refused`
  - `date1904` book → `WorkbookRefused`, only `trust_report.json` and `run.json` exist
  - mapping a formula cell → `CalcInputError`
  - a dropdown input `"X"` not in `["M","F"]` → all that policy's outputs `refused`
  - `--check-against` with a results table where policy 2 took the other branch → that branch then counts as exercised for a new policy
  - `--variables` writes `rows.csv` with statuses
- [ ] **Step 2–4:** run fail, implement, run pass.
- [ ] **Step 5: Commit** — `feat(calc): recalculate input sets and write results, trust report and paths`

---

### Task 11: CLI `marinade calc`, determinism

**Files:**
- Modify: `src/xl_marinade/cli/main.py`
- Test: `tests/calc/test_cli_calc.py`

**Interfaces:**
- Consumes: `calc`, `CalcRequest`, `config` loaders.
- Produces: `marinade calc WORKBOOK --policies PATH --map PATH [--outputs REF]... [--variables PATH] [--check-against PATH] [--ir PATH] [--as-of YYYY-MM-DD] [--tolerance FLOAT] [--probe-cap INT] -o DIR`. Help text starts "Experimental:". Exit codes: 0 ok, 1 `MarinadeError`, 5 `CalcUnavailable` (with hint `pip install 'xl-marinade[calc]'`), 6 `WorkbookRefused`. Update the module docstring's exit-code line.

- [ ] **Step 1: Failing tests** — exit codes 0/1/5/6 via `typer.testing.CliRunner`; **two runs into two directories produce byte-identical `results.csv`, `trust_report.json`, `paths.csv`, `run.json`** (Review Focus 5); stdout stays clean (pattern of `tests/test_extract_stdout_clean.py`).
- [ ] **Step 2–4:** run fail, implement, run pass.
- [ ] **Step 5: Commit** — `feat(cli): marinade calc and marinade values commands`

---

### Task 12: Docs, skill, changelog, upstream reproductions

**Files:**
- Create: `docs/guide/calc.md`
- Modify: `mkdocs.yml` (nav), `skills/using-xl-marinade/SKILL.md` (section "Recalculating with `marinade calc` (experimental)"), generated manifests via `uv run python scripts/gen_skill_manifests.py`, `CHANGELOG.md` (`## [Unreleased]` → `### Added`), `README.md` (one line under commands), `docs/reference/cli.md` (both commands, every option, exit codes 5 and 6), `docs/ai/skills.md` (the workflow lines that list `extract` / `document` / `diff`), `ref/calc/README.md` (status line: PR 1 built, experimental)
- Create: `ref/calc/upstream/` with one openpyxl script per IronCalc gap (text-vs-number approximate `VLOOKUP`, implicit intersection, `'S'!A1:'S'!A3` range, 1904 dates, precision as displayed) — each prints Excel's documented result and IronCalc's

- [ ] **Step 1:** Write the guide: what `calc` is for, the four files in, the files out, the statuses and what each means, the fixed points, "experimental" up front, and that `period_start` has no default and periods count from it.
- [ ] **Step 2:** Run `uv run pytest tests/skills -q` and `uv run mkdocs build --strict` → PASS.
- [ ] **Step 3:** Run the full gate: `uv run ruff format . && uv run ruff check . && uv run mypy src/xl_marinade && uv run pyright src/xl_marinade && uv run --extra calc pytest -q && uv run pytest -q` (the last without the extra: calc engine tests skip, everything else passes).
- [ ] **Step 4: Commit** — `docs(calc): guide, skill section, changelog and upstream reproductions`
