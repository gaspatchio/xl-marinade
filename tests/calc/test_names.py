"""Defined names and targets resolved the way Excel resolves them.

Hand values follow Excel's documented rules:

- A formula on a sheet reads that sheet's own (sheet-scoped) definition of a name
  first, and the workbook's only when the sheet has none; ``Sheet!Name`` reads
  Sheet's own definition.
- Names and sheet names are case-insensitive.
"""

from __future__ import annotations

import csv
import json
from collections.abc import Callable
from pathlib import Path

import pytest

pytest.importorskip("ironcalc")

from tests.calc.conftest import F  # noqa: E402
from xl_marinade.calc import CalcRequest, CalcResult, calc  # noqa: E402
from xl_marinade.calc.graph import WorkbookGraph  # noqa: E402
from xl_marinade.calc.models import CellRef  # noqa: E402
from xl_marinade.calc.patterns import scan_patterns  # noqa: E402
from xl_marinade.calc.probes import anchor  # noqa: E402
from xl_marinade.core.parser import FormulaParser  # noqa: E402
from xl_marinade.errors import CalcInputError  # noqa: E402

Book = Callable[..., Path]
IrFor = Callable[[Path], Path]

IMPLICIT = "implicit intersection on multi-cell name"
UNSORTED = "approximate lookup on unsorted column"


def _write(tmp: Path, name: str, text: str) -> Path:
    path = tmp / name
    path.write_text(text, encoding="utf-8")
    return path


def _run(
    tmp: Path,
    book: Path,
    policies: str,
    fields: dict[str, str],
    outputs: list[str],
    **kw: object,
) -> CalcResult:
    return calc(
        CalcRequest(
            workbook=book,
            policies=_write(tmp, "policies.csv", policies),
            input_map=_write(tmp, "map.json", json.dumps({"fields": fields})),
            out_dir=tmp / "out",
            outputs=outputs,
            **kw,  # type: ignore[arg-type]
        )
    )


def _results(res: CalcResult) -> dict[tuple[str, str], dict[str, str]]:
    with res.results_csv.open(encoding="utf-8", newline="") as fh:
        return {(r["policy_id"], r["output"]): r for r in csv.DictReader(fh)}


def _paths(res: CalcResult) -> list[tuple[str, str, str, str]]:
    with res.paths_csv.open(encoding="utf-8", newline="") as fh:
        return [
            (r["policy_id"], r["cell"], r["outcome"], r["exercised"]) for r in csv.DictReader(fh)
        ]


# -- a sheet's own definition shadows the workbook's ---------------------------

# ``Rate`` is Other!A1 (0.04) for the workbook, and Calc!A5 for sheet Calc: a sheet
# copied in from another workbook brings its own definition along.
SHADOWED: dict[str, dict[str, object]] = {
    "Inputs": {"B1": 100},
    "Other": {"A1": 0.04},
    "Calc": {"A5": 0.03, "B1": F("=IF(Rate>0.05,Inputs!B1*2,Inputs!B1)", 100)},
}
SHADOWED_NAMES = {"Rate": "Other!$A$1"}
SHADOWED_LOCAL = {("Calc", "Rate"): "Calc!$A$5"}


@pytest.fixture
def shadowed(make_book: Book) -> Path:
    return make_book("shadow", SHADOWED, names=SHADOWED_NAMES, local_names=SHADOWED_LOCAL)


def test_branch_on_a_sheets_own_name_is_never_ok(shadowed: Path, tmp_path: Path) -> None:
    # Calc!B1 reads Calc's Rate (Calc!A5). Stored 0.03: FALSE. Policy P sets 0.06: in
    # Excel 0.06 > 0.05 is TRUE. The probe sheet would read the workbook's Rate, and a
    # name's stored reference may be relative, so the probe stays unplaced: never ok.
    res = _run(
        tmp_path,
        shadowed,
        "id,rate\nP,0.06\nS,0.03\n",
        {"rate": "Calc!A5"},
        ["Calc!B1"],
    )
    got = _results(res)
    for pid in ("P", "S"):
        row = got[(pid, "Calc!B1")]
        assert row["status"] == "uncertain"
        assert "not_checked_path@Calc!B1" in row["reasons"]
    assert got[("P", "Calc!B1")]["value"] == "200.0"


def test_anchor_leaves_a_sheets_own_name_unplaced(shadowed: Path, ir_for: IrFor) -> None:
    g = WorkbookGraph.from_ir(ir_for(shadowed))

    def lift(text: str, sheet: str) -> str:
        return anchor(FormulaParser().parse(text), sheet, names=g)  # type: ignore[arg-type]

    # The probe sheet sees only workbook names, and the engine reads no Sheet!Name.
    assert lift("Rate>0.05", "Other") == "Rate>0.05"
    for text, sheet in [("Rate>0.05", "Calc"), ("rate>0.05", "Calc"), ("Calc!Rate>0.05", "In")]:
        with pytest.raises(ValueError, match="sheet-scoped"):
            lift(text, sheet)


def test_offset_base_on_a_sheets_own_name(make_book: Book, tmp_path: Path) -> None:
    # On Calc, Base is Calc!A1; OFFSET(Base,1,0) lands on Calc!A2 (Excel 3). The base
    # cannot be lifted onto the probe sheet, so the reference is never followed: refused.
    book = make_book(
        "offset",
        {
            "Inputs": {"B1": 0},
            "Other": {"A1": 7, "A2": 9},
            "Calc": {"A1": 5, "A2": "=1+2", "C6": F("=OFFSET(Base,Inputs!B1,0)", 5)},
        },
        names={"Base": "Other!$A$1"},
        local_names={("Calc", "Base"): "Calc!$A$1"},
    )
    res = _run(tmp_path, book, "id,n\nP,1\nS,0\n", {"n": "Inputs!B1"}, ["Calc!C6"])
    got = _results(res)
    for pid in ("P", "S"):
        assert got[(pid, "Calc!C6")]["status"] == "refused"
        assert "dynamic reference not checked" in got[(pid, "Calc!C6")]["reasons"]


def test_scans_read_the_sheets_own_name(make_book: Book, ir_for: IrFor) -> None:
    # Col: one cell for the workbook, five on Calc. Tbl: sorted for the workbook,
    # unsorted on Calc. Each formula is scanned against the definition it reads.
    book = make_book(
        "scan",
        {
            "Other": {
                "A1": 1,
                "E1": 1,
                "E2": 2,
                "E3": 3,
                "C3": F("=Col*2", 2),
                "F1": F("=MATCH(2,Tbl,1)", 2),
            },
            "Calc": {
                **{f"A{r}": r for r in range(1, 6)},
                "E1": 3,
                "E2": 1,
                "E3": 2,
                "C3": F("=Col*2", 6),
                "F1": F("=MATCH(2,Tbl,1)", 3),
            },
        },
        names={"Col": "Other!$A$1", "Tbl": "Other!$E$1:$E$3"},
        local_names={("Calc", "Col"): "Calc!$A$1:$A$5", ("Calc", "Tbl"): "Calc!$E$1:$E$3"},
    )
    found = scan_patterns(WorkbookGraph.from_ir(ir_for(book)), as_of=None, data_table_cells=[])

    def details(a1: str) -> list[str]:
        return [r.detail for r in found.get(CellRef.parse(a1), [])]

    assert details("Calc!C3") == [IMPLICIT]
    assert details("Calc!F1") == [UNSORTED]
    assert details("Other!C3") == []
    assert details("Other!F1") == []


# -- targets -------------------------------------------------------------------


def test_targets_ignore_case(make_book: Book, ir_for: IrFor) -> None:
    book = make_book("case", {"In": {"A1": 1, "A2": 2}}, names={"RateCell": "In!$A$1"})
    g = WorkbookGraph.from_ir(ir_for(book))
    a1, a2 = CellRef("In", 1, 1), CellRef("In", 2, 1)
    assert g.resolve_target("ratecell") == [a1]
    assert g.resolve_target("RATECELL") == [a1]
    assert g.resolve_target("in!a1") == [a1]
    assert g.resolve_target("'IN'!A1:a2") == [a1, a2]


def test_sheet_scoped_names_as_targets(shadowed: Path, ir_for: IrFor) -> None:
    g = WorkbookGraph.from_ir(ir_for(shadowed))
    a5 = CellRef("Calc", 5, 1)
    # Ruling D7: 'Sheet'!Name names a sheet's own definition.
    assert g.resolve_target("'Calc'!Rate") == [a5]
    assert g.resolve_target("calc!RATE") == [a5]
    # Outside any sheet a bare Rate could mean either definition.
    with pytest.raises(CalcInputError, match=r"'Calc'!Rate"):
        g.resolve_target("Rate")
    # Other has no Rate of its own.
    with pytest.raises(CalcInputError, match="Other"):
        g.resolve_target("Other!Rate")


def test_a_name_only_a_sheet_defines(make_book: Book, ir_for: IrFor) -> None:
    book = make_book("local", {"In": {"A1": 1}}, local_names={("In", "LocalX"): "In!$A$1"})
    g = WorkbookGraph.from_ir(ir_for(book))
    assert g.resolve_target("In!LocalX") == [CellRef("In", 1, 1)]
    with pytest.raises(CalcInputError, match=r"'In'!LocalX"):
        g.resolve_target("LocalX")


def test_map_field_on_a_shadowed_bare_name_is_refused(shadowed: Path, tmp_path: Path) -> None:
    # A bare Rate in map.json would set Other!A1, which Calc!B1 never reads.
    with pytest.raises(CalcInputError, match="Rate"):
        _run(tmp_path, shadowed, "id,rate\nP,0.06\n", {"rate": "Rate"}, ["Calc!B1"])


def test_check_against_sheets_ignore_case(make_book: Book, tmp_path: Path) -> None:
    book = make_book(
        "chk",
        {
            "Inputs": {"B1": 1},
            "Calc": {"B1": F("=Inputs!B1*10", 10)},
            "Results": {"A2": "a", "B2": 20, "A3": "b", "B3": 30},
            "Table": {"A2": "a", "B2": 2, "A3": "b", "B3": 3},
        },
    )
    check = {
        "results": {
            "sheet": "results",
            "id_column": "A",
            "first_row": 2,
            "outputs": {"Calc!B1": "B"},
        },
        "inputs": {
            "table": {"sheet": "TABLE", "id_column": "A", "first_row": 2, "fields": {"x": "B"}}
        },
    }
    res = _run(
        tmp_path,
        book,
        "id,x\nP,3\n",
        {"x": "Inputs!B1"},
        ["Calc!B1"],
        check_against=_write(tmp_path, "check.json", json.dumps(check)),
    )
    report = json.loads(res.trust_report.read_text(encoding="utf-8"))
    assert report["check_against"]["input_sets"] == 2
    assert report["check_against"]["mismatches"] == []


def test_indirect_text_names_a_sheet_in_any_case(make_book: Book, tmp_path: Path) -> None:
    # INDIRECT("rates!B"&n) reads sheet Rates (Excel ignores case). n=2 lands on
    # Rates!B2, a formula with no stored value the trust run cannot check -> refused.
    # INDIRECT of a name's text is not followed, so it is never ok.
    book = make_book(
        "ind",
        {
            "Inputs": {"B1": 3},
            "Rates": {"B2": "=1+2", "B3": 7},
            "Calc": {
                "A1": F('=INDIRECT("rates!B"&Inputs!B1)', 7),
                "A2": F('=INDIRECT("Rt")', 7),
            },
        },
        names={"Rt": "Rates!$B$3"},
    )
    res = _run(tmp_path, book, "id,n\nP,2\nS,3\n", {"n": "Inputs!B1"}, ["Calc!A1", "Calc!A2"])
    got = _results(res)
    assert got[("P", "Calc!A1")]["status"] == "refused"
    assert "no_stored_value@Rates!B2" in got[("P", "Calc!A1")]["reasons"]
    assert (got[("S", "Calc!A1")]["value"], got[("S", "Calc!A1")]["status"]) == ("7.0", "ok")
    assert got[("P", "Calc!A2")]["status"] == "refused"
    assert got[("S", "Calc!A2")]["status"] == "refused"


# -- whole columns and rows, and names that hold no cells ----------------------

WHOLE_LINE = "approximate lookup on whole column"


def test_whole_line_names_are_multi_cell_in_the_scans(make_book: Book, ir_for: IrFor) -> None:
    # Excel: a single-cell use of a whole-column name intersects it with the formula's
    # row (Calc!D3 reads Data!A3); a whole-row name with the formula's column.
    book = make_book(
        "wl",
        {
            "Data": {**{f"A{r}": 0 for r in range(1, 7)}, "D1": 0},
            "Calc": {
                "D3": F("=IFERROR(WholeA*3,-1)", 0),
                "D4": F("=WholeA*3", 0),
                "D5": F("=IFERROR(WholeR*3,-1)", 0),
                "E1": F("=SUM(WholeA)", 0),
                "E2": F("=MATCH(0,WholeA,1)", 6),
            },
        },
        names={"WholeA": "Data!$A:$A", "WholeR": "Data!$1:$1"},
    )
    found = scan_patterns(WorkbookGraph.from_ir(ir_for(book)), as_of=None, data_table_cells=[])

    def details(a1: str) -> list[str]:
        return [r.detail for r in found.get(CellRef.parse(a1), [])]

    assert details("Calc!D3") == [IMPLICIT]
    assert details("Calc!D4") == [IMPLICIT]
    assert details("Calc!D5") == [IMPLICIT]
    assert details("Calc!E1") == []
    assert details("Calc!E2") == [WHOLE_LINE]


def test_whole_column_name_under_iferror_is_refused(make_book: Book, tmp_path: Path) -> None:
    # Policy P sets Data!A3 to 5: Excel gives 5*3 = 15 in Calc!D3. The engine reads the
    # column's first cell (0) and agrees with Excel only on the stored zeros.
    book = make_book(
        "k9",
        {
            "Data": {f"A{r}": 0 for r in range(1, 7)},
            "Calc": {"D3": F("=IFERROR(WholeA*3,-1)", 0)},
        },
        names={"WholeA": "Data!$A:$A"},
    )
    res = _run(tmp_path, book, "id,a3\nP,5\n", {"a3": "Data!A3"}, ["Calc!D3"])
    row = _results(res)[("P", "Calc!D3")]
    assert row["status"] == "refused"
    assert IMPLICIT in row["reasons"]


@pytest.fixture
def lines(make_book: Book) -> Path:
    # Calc's used extent: rows 1-4, columns A-C.
    return make_book(
        "lines",
        {
            "In": {"A1": 2},
            "Calc": {
                **{f"B{k}": F(f"=In!A1*{k}", 2 * k) for k in range(1, 5)},
                "C1": F("=SUM(Calc!B1:B4)", 20),
            },
            "Empty": {},
        },
        names={
            "WholeB": "Calc!$B:$B",
            "Row1": "Calc!$1:$1",
            "Dyn": "OFFSET(Calc!$B$1,0,0,In!$A$1,1)",
            "Net": "Calc!$C$1-Calc!$B$1",
            "Const": "0.05",
            "Broken": "#REF!",
            "EmptyCol": "Empty!$B:$B",
        },
    )


def test_whole_line_targets_resolve_to_the_used_extent(lines: Path, ir_for: IrFor) -> None:
    g = WorkbookGraph.from_ir(ir_for(lines))
    col_b = [CellRef("Calc", r, 2) for r in range(1, 5)]
    assert g.resolve_target("WholeB") == col_b
    assert g.resolve_target("Calc!B:B") == col_b
    assert g.resolve_target("calc!$B:$B") == col_b
    assert g.resolve_target("Row1") == [CellRef("Calc", 1, c) for c in range(1, 4)]
    assert g.resolve_target("Calc!1:1") == [CellRef("Calc", 1, c) for c in range(1, 4)]


@pytest.mark.parametrize(
    ("target", "message"),
    [
        ("Dyn", "'Dyn' holds a formula"),
        ("Net", "'Net' holds a formula"),
        ("Const", "'Const' holds a formula"),
        ("Broken", "'Broken' is broken"),
        ("EmptyCol", "'EmptyCol' holds no cells"),
        ("Empty!B:B", "holds no cells"),
    ],
)
def test_targets_that_hold_no_cells_are_refused(
    lines: Path, ir_for: IrFor, target: str, message: str
) -> None:
    g = WorkbookGraph.from_ir(ir_for(lines))
    with pytest.raises(CalcInputError, match=message):
        g.resolve_target(target)


def test_external_name_target_is_refused(make_book: Book, ir_for: IrFor, tmp_path: Path) -> None:
    from tests.calc.conftest import add_external_link

    base = make_book("xt", {"S": {"A1": 1}}, names={"Ext": "[1]Src!$A$1:$A$3"})
    book = add_external_link(base, tmp_path / "xt2.xlsx", "Src", {"A1": "1"})
    g = WorkbookGraph.from_ir(ir_for(book))
    with pytest.raises(CalcInputError, match="another workbook"):
        g.resolve_target("Ext")


def test_whole_column_name_as_output(lines: Path, tmp_path: Path) -> None:
    # x = 3: Calc!B1..B4 = 3, 6, 9, 12.
    res = _run(tmp_path, lines, "id,x\nP,3\n", {"x": "In!A1"}, ["WholeB"])
    got = _results(res)
    assert [got[("P", f"Calc!B{k}")]["value"] for k in range(1, 5)] == [
        "3.0",
        "6.0",
        "9.0",
        "12.0",
    ]


def test_formula_name_as_the_only_output_names_it(lines: Path, tmp_path: Path) -> None:
    with pytest.raises(CalcInputError, match="'Dyn'"):
        _run(tmp_path, lines, "id,x\nP,3\n", {"x": "In!A1"}, ["Dyn"])


def test_formula_name_as_a_variable_range_is_refused(lines: Path, tmp_path: Path) -> None:
    from xl_marinade.calc import ValuesRequest, values

    vars_json = _write(
        tmp_path, "vars.json", json.dumps({"dyn": {"range": "Dyn", "period_start": 1}})
    )
    with pytest.raises(CalcInputError, match="'Dyn'"):
        _run(tmp_path, lines, "id,x\nP,3\n", {"x": "In!A1"}, [], variables=vars_json)
    with pytest.raises(CalcInputError, match="'Dyn'"):
        values(ValuesRequest(lines, vars_json, tmp_path / "vout"))


# -- a variable range lies on one sheet ----------------------------------------


@pytest.fixture
def unions(make_book: Book) -> Path:
    # P1!A1:A2 = x*10, x*20 and P2!A1:A2 = x*30, x*40, with x = In!A1 = 2.
    return make_book(
        "unions",
        {
            "In": {"A1": 2},
            "P1": {"A1": F("=In!A1*10", 20), "A2": F("=In!A1*20", 40)},
            "P2": {"A1": F("=In!A1*30", 60), "A2": F("=In!A1*40", 80)},
            "S": {"A1": 1, "A2": 2, "A4": 4, "A5": 5},
        },
        names={"Both": "P1!$A$1:$A$2,P2!$A$1:$A$2", "Gap": "S!$A$1:$A$2,S!$A$4:$A$5"},
    )


def test_a_variable_range_across_two_sheets_is_refused(unions: Path, tmp_path: Path) -> None:
    from xl_marinade.calc import ValuesRequest, values

    # Sorted by row, the periods would alternate between P1 and P2 (20, 60, 40, 80).
    vars_json = _write(
        tmp_path, "vars.json", json.dumps({"both": {"range": "Both", "period_start": 1}})
    )
    with pytest.raises(CalcInputError, match="more than one sheet"):
        _run(tmp_path, unions, "id,x\nP,3\n", {"x": "In!A1"}, [], variables=vars_json)
    with pytest.raises(CalcInputError, match="more than one sheet"):
        values(ValuesRequest(unions, vars_json, tmp_path / "vout"))


def test_one_sheet_unions_and_multi_sheet_outputs_still_resolve(
    unions: Path, tmp_path: Path
) -> None:
    from xl_marinade.calc import ValuesRequest, values

    # A one-sheet union in one column runs down the rows; --outputs is a set of cells,
    # so a union across sheets is fine there.
    vars_json = _write(
        tmp_path, "vars.json", json.dumps({"gap": {"range": "Gap", "period_start": 1}})
    )
    rows = values(ValuesRequest(unions, vars_json, tmp_path / "vout"))
    with rows.open(encoding="utf-8", newline="") as fh:
        got = [(r["period"], r["value"]) for r in csv.DictReader(fh)]
    assert got == [("1", "1.0"), ("2", "2.0"), ("3", "4.0"), ("4", "5.0")]
    res = _run(tmp_path, unions, "id,x\nP,3\n", {"x": "In!A1"}, ["Both"])
    values_by_cell = {out: r["value"] for (_p, out), r in _results(res).items()}
    assert values_by_cell == {"P1!A1": "30.0", "P1!A2": "60.0", "P2!A1": "90.0", "P2!A2": "120.0"}


# -- names whose reference is relative ------------------------------------------

RELATIVE = "name with a relative reference: Left"


@pytest.mark.parametrize("scope", ["sheet", "workbook"])
def test_a_relative_name_in_a_formula_is_refused(
    make_book: Book, tmp_path: Path, scope: str
) -> None:
    # Left is stored relative to A1 as Calc!XFD1: one column left of the cell using it.
    # Excel: at Calc!C5, Left is Calc!B5. P sets x = 10, so B5 = 10 > 5 and C5 = 100;
    # IronCalc reads Calc!XFD1 (empty) and gives 200. Refused for every policy.
    sheets: dict[str, dict[str, object]] = {
        "In": {"B1": 1},
        "Calc": {"B5": F("=In!B1", 1), "C5": F("=IF(Left>5,100,200)", 200)},
    }
    where = {"names": {"Left": "Calc!XFD1"}}
    if scope == "sheet":
        where = {"local_names": {("Calc", "Left"): "Calc!XFD1"}}
    book = make_book(f"rel_{scope}", sheets, **where)
    res = _run(tmp_path, book, "id,x\nP,10\nS,1\n", {"x": "In!B1"}, ["Calc!C5"])
    got = _results(res)
    for pid in ("P", "S"):
        assert got[(pid, "Calc!C5")]["status"] == "refused"
        assert f"refused_pattern@Calc!C5:{RELATIVE}" in got[(pid, "Calc!C5")]["reasons"]


def test_a_relative_name_as_a_target_is_refused(make_book: Book, tmp_path: Path) -> None:
    book = make_book(
        "rel_target",
        {"In": {"B1": 1}, "Calc": {"B5": F("=In!B1*2", 2)}},
        names={"Here": "Calc!B5", "Fixed": "Calc!$B$5"},
    )
    res = _run(tmp_path, book, "id,x\nP,3\n", {"x": "In!B1"}, ["Fixed"])
    assert _results(res)[("P", "Calc!B5")]["value"] == "6.0"
    with pytest.raises(CalcInputError, match="'Here' has a relative reference"):
        _run(tmp_path, book, "id,x\nP,3\n", {"x": "In!B1"}, ["Here"])


def test_a_relative_name_as_a_list_source_is_refused(make_book: Book, tmp_path: Path) -> None:
    # The list Opts is Lists!A1:A2 written without $, so its cells depend on the cell
    # validated: Excel's list for In!B1 is not knowable from the stored reference.
    book = make_book(
        "rel_dv",
        {
            "In": {"B1": "a"},
            "Lists": {"A1": "a", "A2": "b"},
            "Calc": {"B1": F('=In!B1&"!"', "a!")},
        },
        names={"Opts": "Lists!A1:A2"},
        list_sources=[("In!B1", "Opts")],
    )
    res = _run(tmp_path, book, "id,x\nP,b\n", {"x": "In!B1"}, ["Calc!B1"])
    row = _results(res)[("P", "Calc!B1")]
    assert row["status"] == "refused"
    assert "refused_pattern@In!B1:name with a relative reference: Opts" in row["reasons"]
