"""Every cell of an array formula's range is a formula cell.

Excel stores an array formula once, on the range's top-left cell (the anchor); the other
cells hold only Excel's results. They are computed by the anchor's formula, so the
graph links each to its anchor, the trust run compares each, and none takes an input.

Hand values come from Excel's documented semantics: an array formula over ``A1:A3``
times ``x`` fills its three cells with ``A1*x``, ``A2*x``, ``A3*x``; a legacy
(Ctrl+Shift+Enter) array formula entered in more cells than its result has fills the
extra cells with ``#N/A``; ``SEQUENCE(n)`` spills ``1..n`` down from its cell.
"""

from __future__ import annotations

import csv
import json
import re
from collections.abc import Callable
from pathlib import Path

import pytest

pytest.importorskip("ironcalc")

from tests.calc.conftest import F, patch_book  # noqa: E402
from xl_marinade.calc import CalcRequest, CalcResult, calc  # noqa: E402
from xl_marinade.errors import CalcInputError  # noqa: E402

Book = Callable[..., Path]


def _arrays(
    src: Path, sheet_part: str, arrays: list[tuple[str, str, str, dict[str, object], bool]]
) -> Path:
    """Turn plain formulas into array formulas as Excel saves them.

    Each entry is (anchor, range, formula XML, the other cells' stored values, dynamic):
    the anchor's ``<f>`` becomes ``<f t="array" ref=range>`` (with ``cm="1"`` on the
    cell for a dynamic array) and every other cell holds a bare stored ``<v>``.
    """

    def edit(xml: str) -> str:
        for anchor, rng, formula, others, dynamic in arrays:
            pat = re.compile(rf'(<c r="{anchor}")([^>]*>)<f>{re.escape(formula)}</f>')
            assert pat.search(xml), (anchor, formula)
            cm = ' cm="1"' if dynamic else ""
            f_xml = f'<f t="array" ref="{rng}">{formula}</f>'
            xml = pat.sub(lambda m, cm=cm, f_xml=f_xml: f"{m.group(1)}{cm}{m.group(2)}{f_xml}", xml)
            for ref, v in others.items():
                cell = re.compile(rf'<c r="{ref}"[^>]*>.*?</c>|<c r="{ref}"[^>]*/>', re.S)
                assert cell.search(xml), ref
                if isinstance(v, str) and v.startswith("#"):
                    new = f'<c r="{ref}" t="e"><v>{v}</v></c>'
                else:
                    new = f'<c r="{ref}"><v>{v!r}</v></c>'
                xml = cell.sub(new, xml)
        return xml

    return patch_book(src, src.with_name(f"{src.stem}_arr.xlsx"), edit={sheet_part: edit})


def _run(
    tmp: Path,
    book: Path,
    policies: str,
    fields: dict[str, str],
    outputs: list[str],
    *,
    variables: dict[str, object] | None = None,
    check: dict[str, object] | None = None,
) -> CalcResult:
    (tmp / "p.csv").write_text(policies, encoding="utf-8")
    (tmp / "m.json").write_text(json.dumps({"id_field": "id", "fields": fields}), encoding="utf-8")
    vars_path = check_path = None
    if variables is not None:
        vars_path = tmp / "v.json"
        vars_path.write_text(json.dumps(variables), encoding="utf-8")
    if check is not None:
        check_path = tmp / "c.json"
        check_path.write_text(json.dumps(check), encoding="utf-8")
    return calc(
        CalcRequest(
            workbook=book,
            policies=tmp / "p.csv",
            input_map=tmp / "m.json",
            out_dir=tmp / "out",
            outputs=outputs,
            variables=vars_path,
            check_against=check_path,
        )
    )


def _rows(res: CalcResult) -> dict[tuple[str, str], dict[str, str]]:
    with res.results_csv.open(encoding="utf-8", newline="") as fh:
        return {(r["policy_id"], r["output"]): r for r in csv.DictReader(fh)}


def _report(res: CalcResult) -> dict[str, object]:
    return json.loads(res.trust_report.read_text(encoding="utf-8"))


_TODAY = "refused_pattern@Data!A2:TODAY"


def _refused_anchor_book(make_book: Book, *, dynamic: bool) -> Path:
    # Data!A2 = TODAY()*0+2 is always 2, but TODAY without --as-of is a refused pattern.
    # Calc!B1:B3 = {Data!A1:A3*Inputs!B1}: x, 2x, 3x. G1 = B2+1 = 2x+1.
    x = 2
    book = make_book(
        "arr_refused",
        {
            "Inputs": {"B1": x},
            "Data": {"A1": 1, "A2": F("=TODAY()*0+2", 2), "A3": 3},
            "Calc": {
                "B1": F("=Data!A1:A3*Inputs!B1", x),
                "B2": 0,
                "B3": 0,
                "G1": F("=Calc!B2+1", 2 * x + 1),
            },
        },
    )
    spec = ("B1", "B1:B3", "Data!A1:A3*Inputs!B1", {"B2": 2.0 * x, "B3": 3.0 * x}, dynamic)
    return _arrays(book, "xl/worksheets/sheet3.xml", [spec])


@pytest.mark.parametrize("dynamic", [False, True], ids=["legacy", "dynamic"])
def test_array_cells_and_their_readers_carry_the_anchors_refusal(
    make_book: Book, tmp_path: Path, dynamic: bool
) -> None:
    book = _refused_anchor_book(make_book, dynamic=dynamic)
    res = _run(tmp_path, book, "id,x\nP,10\n", {"x": "Inputs!B1"}, ["Calc!B1:B3", "Calc!G1"])
    rows = _rows(res)
    # Excel for x = 10: B1:B3 = 10, 20, 30; G1 = 21.
    want = {"Calc!B1": "10.0", "Calc!B2": "20.0", "Calc!B3": "30.0", "Calc!G1": "21.0"}
    for out, value in want.items():
        row = rows[("P", out)]
        assert row["value"] == value, (out, row)
        assert row["status"] == "refused", (out, row)
        assert _TODAY in row["reasons"], (out, row)


def test_array_variables_carry_the_anchors_refusal(make_book: Book, tmp_path: Path) -> None:
    book = _refused_anchor_book(make_book, dynamic=True)
    variables = {"v": {"range": "Calc!B1:B3", "period_start": 0}}
    res = _run(tmp_path, book, "id,x\nP,10\n", {"x": "Inputs!B1"}, [], variables=variables)
    assert res.rows_csv is not None
    with res.rows_csv.open(encoding="utf-8", newline="") as fh:
        rows = list(csv.DictReader(fh))
    assert [(r["period"], r["value"], r["status"]) for r in rows] == [
        ("0", "10.0", "refused"),
        ("1", "20.0", "refused"),
        ("2", "30.0", "refused"),
    ]


def test_a_branch_inside_an_array_formula_never_gives_ok(make_book: Book, tmp_path: Path) -> None:
    # K1:K3 = {IF(Inputs!B1>5, Data!C1:C3*2, Data!C1:C3)} over C = 1, 2, 3; L1 = K2*10.
    # Stored x = 1: 1, 2, 3 and L1 = 20. x = 9 takes the TRUE arm: 2, 4, 6 and L1 = 40.
    # A probe inside an array formula cannot be placed, so no cell it computes is ok.
    book = make_book(
        "arr_branch",
        {
            "Inputs": {"B1": 1},
            "Data": {"C1": 1, "C2": 2, "C3": 3},
            "Calc": {
                "K1": F("=IF(Inputs!B1>5,Data!C1:C3*2,Data!C1:C3)", 1),
                "K2": 0,
                "K3": 0,
                "L1": F("=Calc!K2*10", 20),
            },
        },
    )
    formula = "IF(Inputs!B1&gt;5,Data!C1:C3*2,Data!C1:C3)"
    spec = ("K1", "K1:K3", formula, {"K2": 2.0, "K3": 3.0}, False)
    book = _arrays(book, "xl/worksheets/sheet3.xml", [spec])
    res = _run(tmp_path, book, "id,x\nP,9\n", {"x": "Inputs!B1"}, ["Calc!K2", "Calc!L1"])
    rows = _rows(res)
    assert rows[("P", "Calc!K2")]["value"] == "4.0"
    assert rows[("P", "Calc!L1")]["value"] == "40.0"
    for out in ("Calc!K2", "Calc!L1"):
        row = rows[("P", out)]
        assert row["status"] != "ok", (out, row)
        assert "not_checked_path@Calc!K1:branch not checked" in row["reasons"], (out, row)


def test_every_array_cell_is_compared_in_the_trust_run(make_book: Book, tmp_path: Path) -> None:
    # E1:E3 = {S!A1:A2*2} entered in three cells over a two-cell result: Excel gives 6, 0
    # and #N/A in the cell the result does not reach (A1 = 3, A2 empty). The anchor
    # reproduces; IronCalc 0.9.1 gives #VALUE! in E3. H1 = E3 reads it.
    book = make_book(
        "arr_trust",
        {
            "S": {
                "A1": 3,
                "B1": "id",
                "E1": F("=S!A1:A2*2", 6),
                "E2": 0,
                "E3": 0,
                "H1": F("=IFERROR(S!E3,-1)", -1),
            }
        },
    )
    spec = ("E1", "E1:E3", "S!A1:A2*2", {"E2": 0.0, "E3": "#N/A"}, False)
    book = _arrays(book, "xl/worksheets/sheet1.xml", [spec])
    # The id goes into S!B1, which nothing reads (a map field nothing reads is refused).
    res = _run(tmp_path, book, "id\nP\n", {"id": "S!B1"}, ["S!E1:E3", "S!H1"])
    report = _report(res)
    assert report["formula_cells"] == 4  # E1, its two other cells, H1
    assert report["checked"] == 4
    failing = {f["cell"] for f in report["failures"]}  # type: ignore[union-attr]
    assert "S!E3" in failing
    rows = _rows(res)
    assert rows[("P", "S!E1")]["status"] == "ok"
    assert rows[("P", "S!E3")]["status"] == "refused"
    assert rows[("P", "S!H1")]["status"] == "refused"
    assert "@S!E3:stored '#N/A'" in rows[("P", "S!H1")]["reasons"]


def _transpose_book(make_book: Book) -> Path:
    # Inputs!B1:B3 = {TRANSPOSE(Data!A1:C1)} = 40, 10, 5; Calc!C1 = Inputs!B2*2 = 20.
    book = make_book(
        "arr_input",
        {
            "Data": {"A1": 40, "B1": 10, "C1": 5},
            "Inputs": {"B1": F("=TRANSPOSE(Data!A1:C1)", 40), "B2": 0, "B3": 0, "D1": "P"},
            "Calc": {"C1": F("=Inputs!B2*2", 20)},
            "Results": {"A2": "P", "B2": 20},
        },
    )
    spec = ("B1", "B1:B3", "TRANSPOSE(Data!A1:C1)", {"B2": 10.0, "B3": 5.0}, False)
    return _arrays(book, "xl/worksheets/sheet2.xml", [spec])


@pytest.mark.parametrize("target", ["Inputs!B2", "Inputs!B3"])
def test_an_input_inside_an_array_formula_is_an_error(
    make_book: Book, tmp_path: Path, target: str
) -> None:
    # Excel cannot hold a constant in part of an array formula's range.
    book = _transpose_book(make_book)
    with pytest.raises(CalcInputError, match="maps onto formula cell"):
        _run(tmp_path, book, "id,v\nP,99\n", {"v": target}, ["Calc!C1"])


def test_a_selector_inside_an_array_formula_is_an_error(make_book: Book, tmp_path: Path) -> None:
    book = _transpose_book(make_book)
    check = {
        "results": {
            "sheet": "Results",
            "id_column": "A",
            "first_row": 2,
            "outputs": {"Calc!C1": "B"},
        },
        "inputs": {"selector": "Inputs!B3"},
    }
    with pytest.raises(CalcInputError, match="maps onto formula cell"):
        _run(tmp_path, book, "id\nP\n", {"id": "Inputs!D1"}, ["Calc!C1"], check=check)


def test_a_spill_that_grows_carries_its_anchors_status(make_book: Book, tmp_path: Path) -> None:
    # Calc!B1 = SEQUENCE(Inputs!B1) + Data!A1*0 spills 1..n down column B; Data!A1 =
    # TODAY()*0 is a refused pattern without --as-of. The workbook stored n = 3 (B1:B3).
    # For n = 5 the spill reaches B5: G1 = B5*2 = 10 and G2 = SUM(B4:B6) = 4+5 = 9, both
    # computed from cells the stored spill never held.
    book = make_book(
        "spill",
        {
            "Inputs": {"B1": 3},
            "Data": {"A1": F("=TODAY()*0", 0)},
            "Calc": {
                "B1": F("=_xlfn.SEQUENCE(Inputs!B1)+Data!A1*0", 1),
                "B2": 0,
                "B3": 0,
                "G1": F("=Calc!B5*2", 0),
                "G2": F("=SUM(Calc!B4:B6)", 0),
            },
        },
    )
    formula = "_xlfn.SEQUENCE(Inputs!B1)+Data!A1*0"
    book = _arrays(
        book, "xl/worksheets/sheet3.xml", [("B1", "B1:B3", formula, {"B2": 2.0, "B3": 3.0}, True)]
    )
    res = _run(tmp_path, book, "id,n\nP,5\n", {"n": "Inputs!B1"}, ["Calc!G1", "Calc!G2"])
    rows = _rows(res)
    assert rows[("P", "Calc!G1")]["value"] == "10.0"
    assert rows[("P", "Calc!G2")]["value"] == "9.0"
    for out in ("Calc!G1", "Calc!G2"):
        row = rows[("P", out)]
        assert row["status"] == "refused", (out, row)
        assert "refused_pattern@Data!A1:TODAY" in row["reasons"], (out, row)


def test_a_spill_grown_from_an_unprobed_formula_is_not_ok(make_book: Book, tmp_path: Path) -> None:
    # Calc!B1 = SEQUENCE(Inputs!B1)*IF(Inputs!B2>0,1,2): stored n = 3, b = 1 (B1:B3 =
    # 1, 2, 3). G1 = B5*2 reads a cell only a larger spill reaches, so nothing upstream
    # of G1 leads to the formula and it is not probed. n = 5: B5 = 5 and G1 = 10, from
    # a formula whose IF no probe checked.
    book = make_book(
        "spill_branch",
        {
            "Inputs": {"B1": 3, "B2": 1},
            "Calc": {
                "B1": F("=_xlfn.SEQUENCE(Inputs!B1)*IF(Inputs!B2>0,1,2)", 1),
                "B2": 0,
                "B3": 0,
                "G1": F("=Calc!B5*2", 0),
            },
        },
    )
    formula = "_xlfn.SEQUENCE(Inputs!B1)*IF(Inputs!B2&gt;0,1,2)"
    book = _arrays(
        book, "xl/worksheets/sheet2.xml", [("B1", "B1:B3", formula, {"B2": 2.0, "B3": 3.0}, True)]
    )
    res = _run(tmp_path, book, "id,n\nP,5\n", {"n": "Inputs!B1"}, ["Calc!G1"])
    row = _rows(res)[("P", "Calc!G1")]
    assert row["value"] == "10.0"
    assert row["status"] != "ok", row
    assert "not_checked_path@Calc!B1:" in row["reasons"], row


def test_an_input_beside_a_spill_is_not_spill_growth(make_book: Book, tmp_path: Path) -> None:
    # Inputs!B1:B3 = SEQUENCE(3) + Data!A1 (Data!A1 = TODAY()*0, refused without
    # --as-of). A map field sets the empty cell below the spill, Inputs!B4 = 7; the spill
    # stays three cells, so Calc!C1 = B4*2 = 14 owes nothing to the refused formula.
    book = make_book(
        "beside",
        {
            "Data": {"A1": F("=TODAY()*0", 0)},
            "Inputs": {"B1": F("=_xlfn.SEQUENCE(3)+Data!A1", 1), "B2": 0, "B3": 0},
            "Calc": {"C1": F("=Inputs!B4*2", 0), "C2": F("=Inputs!B5*2", 0)},
        },
    )
    formula = "_xlfn.SEQUENCE(3)+Data!A1"
    book = _arrays(
        book, "xl/worksheets/sheet2.xml", [("B1", "B1:B3", formula, {"B2": 2.0, "B3": 3.0}, True)]
    )
    res = _run(tmp_path, book, "id,v\nP,7\n", {"v": "Inputs!B4"}, ["Calc!C1"])
    row = _rows(res)[("P", "Calc!C1")]
    assert row["value"] == "14.0"
    assert row["status"] == "ok", row


def test_a_failing_anchor_refuses_its_other_cells(make_book: Book, tmp_path: Path) -> None:
    # D1:D3 = {MATCH(Inputs!A1:A3, Keys!A1:A3, 1)} over keys 10, 20, 30 with inputs 15,
    # "x", 25: Excel {1; #N/A; 2}. IronCalc 0.9.1 gives #N/A in D1, so the anchor fails
    # the trust run, and D2 (whose own value matches) carries that failure.
    book = make_book(
        "anchorfail",
        {
            "Inputs": {"A1": 15, "A2": "x", "A3": 25},
            "Keys": {"A1": 10, "A2": 20, "A3": 30},
            "Calc": {"D1": F("=MATCH(Inputs!A1:A3,Keys!A1:A3,1)", 1), "D2": 0, "D3": 0, "Z1": "id"},
        },
    )
    spec = ("D1", "D1:D3", "MATCH(Inputs!A1:A3,Keys!A1:A3,1)", {"D2": "#N/A", "D3": 2.0}, False)
    book = _arrays(book, "xl/worksheets/sheet3.xml", [spec])
    res = _run(tmp_path, book, "id\nP\n", {"id": "Calc!Z1"}, ["Calc!D2"])
    row = _rows(res)[("P", "Calc!D2")]
    assert row["status"] == "refused", row
    assert "@Calc!D1:stored 1.0" in row["reasons"], row


def test_an_indirect_landing_inside_an_array_is_not_ok(make_book: Book, tmp_path: Path) -> None:
    # K1:K3 = {IF(x>5, C1:C3*2, C1:C3)} over C = 1, 2, 3; E1 = INDIRECT("K"&k). Stored
    # k = 1, x = 1: 1. Policy k = 2, x = 9: K2 = 4, from an IF no probe can check.
    book = make_book(
        "land_in_array",
        {
            "Inputs": {"B1": 1, "B2": 1},
            "Calc": {
                "C1": 1,
                "C2": 2,
                "C3": 3,
                "K1": F("=IF(Inputs!$B$2>5,Calc!C1:C3*2,Calc!C1:C3)", 1),
                "K2": 0,
                "K3": 0,
                "E1": F('=INDIRECT("K"&Inputs!$B$1)', 1),
            },
        },
    )
    formula = "IF(Inputs!$B$2&gt;5,Calc!C1:C3*2,Calc!C1:C3)"
    book = _arrays(
        book, "xl/worksheets/sheet2.xml", [("K1", "K1:K3", formula, {"K2": 2.0, "K3": 3.0}, False)]
    )
    res = _run(tmp_path, book, "id,k,x\nP,2,9\n", {"k": "Inputs!B1", "x": "Inputs!B2"}, ["Calc!E1"])
    row = _rows(res)[("P", "Calc!E1")]
    assert row["value"] == "4.0"
    assert row["status"] != "ok", row


def test_a_join_over_cells_a_spill_can_grow_into_is_read_per_policy(
    make_book: Book, tmp_path: Path
) -> None:
    # Arr!A1 = 1/SEQUENCE(Inputs!B1) spills down column A; the workbook stored n = 2
    # (A1:A2 = 1, 0.5), so A3:A5 are empty. Calc!B1 = TEXTJOIN(",",TRUE,Arr!A1:A5).
    # n = 2: "1,0.5" in Excel and the engine alike. n = 3: the spill reaches A3 = 1/3,
    # which Excel writes "0.333333333333333" (15 digits) and IronCalc 0.9.1 with 16,
    # so the joined text differs from Excel's and must not come back ok.
    book = make_book(
        "spill_join",
        {
            "Inputs": {"B1": 2},
            "Arr": {"A1": F("=1/_xlfn.SEQUENCE(Inputs!B1)", 1), "A2": 0},
            "Calc": {"B1": F('=_xlfn.TEXTJOIN(",",TRUE,Arr!A1:A5)', "1,0.5")},
        },
    )
    formula = "1/_xlfn.SEQUENCE(Inputs!B1)"
    book = _arrays(book, "xl/worksheets/sheet2.xml", [("A1", "A1:A2", formula, {"A2": 0.5}, True)])
    res = _run(tmp_path, book, "id,n\nP2,2\nP3,3\n", {"n": "Inputs!B1"}, ["Calc!B1"])
    rows = _rows(res)
    assert rows[("P2", "Calc!B1")]["value"] == "1,0.5"
    assert rows[("P2", "Calc!B1")]["status"] == "ok", rows[("P2", "Calc!B1")]["reasons"]
    p3 = rows[("P3", "Calc!B1")]
    assert p3["value"].startswith("1,0.5,0.333333333333333"), p3
    assert p3["status"] == "refused", p3
