"""Lifted probe text computes what its cell computes.

Three ways it did not: ``TODAY()``/``NOW()`` read the machine clock under ``--as-of``;
``&`` was parsed looser than comparisons (Excel binds it tighter: ``A&B="x"`` is
``(A&B)="x"``); and a sheet whose name starts with a digit, or reads as a cell, an R1C1
reference or a boolean, was not quoted. Hand values follow Excel's documented operator
precedence (``:`` / ``-`` / ``%`` / ``^`` / ``*`` ``/`` / ``+`` ``-`` / ``&`` / comparisons)
and its sheet-name quoting.
"""

from __future__ import annotations

import csv
import json
from collections.abc import Callable
from datetime import date
from pathlib import Path

import pytest

pytest.importorskip("ironcalc")

from tests.calc.conftest import F  # noqa: E402
from xl_marinade.calc import CalcRequest, CalcResult, calc  # noqa: E402
from xl_marinade.calc.formula_ast import parse_formula, quote_sheet  # noqa: E402
from xl_marinade.calc.probes import anchor  # noqa: E402

Book = Callable[..., Path]


def _run(
    tmp: Path,
    book: Path,
    policies: str,
    fields: dict[str, str],
    outputs: list[str],
    *,
    as_of: date | None = None,
) -> CalcResult:
    (tmp / "p.csv").write_text(policies, encoding="utf-8")
    (tmp / "m.json").write_text(json.dumps({"id_field": "id", "fields": fields}), encoding="utf-8")
    return calc(
        CalcRequest(
            workbook=book,
            policies=tmp / "p.csv",
            input_map=tmp / "m.json",
            out_dir=tmp / "out",
            outputs=outputs,
            as_of=as_of,
        )
    )


def _rows(res: CalcResult) -> dict[tuple[str, str], dict[str, str]]:
    with res.results_csv.open(encoding="utf-8", newline="") as fh:
        return {(r["policy_id"], r["output"]): r for r in csv.DictReader(fh)}


def _paths(res: CalcResult) -> dict[tuple[str, str], str]:
    with res.paths_csv.open(encoding="utf-8", newline="") as fh:
        return {(r["policy_id"], r["cell"]): r["outcome"] for r in csv.DictReader(fh)}


def _lift(text: str, sheet: str = "S") -> str:
    root = parse_formula(text)
    assert root is not None
    return anchor(root, sheet)


# -- T15: TODAY()/NOW() under --as-of ----------------------------------------------

# 2000-01-01 is serial 36526; any machine clock in this century is later.
_AS_OF = date(2000, 1, 1)


def test_clock_in_a_branch_reads_the_as_of_date(make_book: Book, tmp_path: Path) -> None:
    # Maturity stored 36600 (2000-03-15): as of 2000-01-01, TODAY()>=36600 is FALSE, so
    # Excel stored 1000 (B1) and 1000 (B2, NOW() the same). Policy P (maturity 36000) is
    # matured as of then: TRUE, 0 -- an arm the trust run never took. Policy Q (37000) is
    # not: FALSE, 1000, the stored path.
    book = make_book(
        "clock",
        {
            "Inputs": {"B1": 36600},
            "Calc": {
                "B1": F("=IF(TODAY()>=Inputs!B1,0,1000)", 1000.0),
                "B2": F("=IF(NOW()>=Inputs!B1,0,1000)", 1000.0),
            },
        },
    )
    res = _run(
        tmp_path,
        book,
        "id,m\nP,36000\nQ,37000\n",
        {"m": "Inputs!B1"},
        ["Calc!B1", "Calc!B2"],
        as_of=_AS_OF,
    )
    rows, paths = _rows(res), _paths(res)
    for out in ("Calc!B1", "Calc!B2"):
        assert rows[("P", out)]["value"] == "0.0"
        assert rows[("P", out)]["status"] == "unverified_path"
        assert "took 'TRUE', not exercised" in rows[("P", out)]["reasons"]
        assert rows[("Q", out)]["value"] == "1000.0"
        assert rows[("Q", out)]["status"] == "ok"
        assert paths[("P", out)] == "TRUE"
        assert paths[("Q", out)] == "FALSE"


def test_clock_in_a_dynamic_reference_lands_on_the_as_of_row(
    make_book: Book, tmp_path: Path
) -> None:
    # OFFSET(D1, YEAR(TODAY())-2000, 0) as of 2000-01-01 is D1 (a constant, 5). Rows
    # D2:D60 hold RAND() (refused): a probe reading the machine clock lands there.
    calc_cells: dict[str, object] = {
        "D1": 5,
        "B1": F("=OFFSET(D1,YEAR(TODAY())-2000,0)+Inputs!B1", 6.0),
    }
    calc_cells.update({f"D{r}": F("=RAND()", 0.5) for r in range(2, 61)})
    book = make_book("clockdyn", {"Inputs": {"B1": 1}, "Calc": calc_cells})
    res = _run(tmp_path, book, "id,x\nP,2\n", {"x": "Inputs!B1"}, ["Calc!B1"], as_of=_AS_OF)
    row = _rows(res)[("P", "Calc!B1")]
    assert row["value"] == "7.0"
    assert row["status"] == "ok", row


def test_clock_in_every_probe_kind_reads_the_as_of_date(make_book: Book, tmp_path: Path) -> None:
    # As of 2000-01-01 (36526). Maturity stored 36600, z stored 1, r stored 0.5.
    #   B3 IF(TODAY()>=m, IF(z>0,1,2), 3): stored 3; P (m 36000) takes TRUE and then the
    #      inner IF's TRUE -- an inner arm the trust run never evaluated.
    #   B4 ROUND(m*TODAY()/36526/2, 0): W (m 37001) gives 18500.5, a tie.
    #   B5 IF(TODAY()/36526*r>0.3, 1, 0): R (r 0.30000000000000004) is within 1e-12 of 0.3.
    #   B6 IFS / B7 CHOOSE on TODAY()>=m: P takes the arm the stored inputs did not.
    #   B9 INDEX/MATCH on YEAR(TODAY())+z-1 over the years 2000..2059 in E: stored 2000,
    #      the first row; Z2 (z 2) lands on 2001, an interior row never exercised.
    calc_cells: dict[str, object] = {
        "B3": F("=IF(TODAY()>=Inputs!B1,IF(Inputs!B2>0,1,2),3)", 3.0),
        "B4": F("=ROUND(Inputs!B1*TODAY()/36526/2,0)", 18300.0),
        "B5": F("=IF(TODAY()/36526*Inputs!B3>0.3,1,0)", 1.0),
        "B6": F("=_xlfn.IFS(TODAY()>=Inputs!B1,0,TRUE,1000)", 1000.0),
        "B7": F("=CHOOSE(1+(TODAY()>=Inputs!B1),1000,0)", 1000.0),
        "B9": F("=INDEX(E1:E60,MATCH(YEAR(TODAY())+Inputs!B2-1,E1:E60,0))", 2000.0),
    }
    calc_cells.update({f"E{r}": 1999 + r for r in range(1, 61)})
    book = make_book(
        "clockkinds", {"Inputs": {"B1": 36600, "B2": 1, "B3": 0.5}, "Calc": calc_cells}
    )
    res = _run(
        tmp_path,
        book,
        "id,m,z,r\nP,36000,1,0.5\nW,37001,1,0.5\nR,37000,1,0.30000000000000004\nZ2,37000,2,0.5\n",
        {"m": "Inputs!B1", "z": "Inputs!B2", "r": "Inputs!B3"},
        ["Calc!B3", "Calc!B4", "Calc!B5", "Calc!B6", "Calc!B7", "Calc!B9"],
        as_of=_AS_OF,
    )
    rows = _rows(res)
    assert rows[("P", "Calc!B3")]["value"] == "1.0"
    assert "IF #1 took 'TRUE'" in rows[("P", "Calc!B3")]["reasons"]
    assert rows[("W", "Calc!B4")]["status"] == "uncertain"
    assert "near_rounding_flip@Calc!B4" in rows[("W", "Calc!B4")]["reasons"]
    assert rows[("R", "Calc!B5")]["status"] == "uncertain"
    assert "near_comparison@Calc!B5" in rows[("R", "Calc!B5")]["reasons"]
    for out in ("Calc!B6", "Calc!B7"):
        assert rows[("P", out)]["value"] == "0.0"
        assert rows[("P", out)]["status"] == "unverified_path", out
    assert rows[("Z2", "Calc!B9")]["value"] == "2001.0"
    assert "landed interior" in rows[("Z2", "Calc!B9")]["reasons"]
    for out in ("Calc!B4", "Calc!B5", "Calc!B9"):
        assert rows[("P", out)]["status"] == "ok", (out, rows[("P", out)])


def test_lifted_clock_text_is_the_as_of_date() -> None:
    root = parse_formula('IF(TODAY()-NOW()>A1,"TODAY()",now())')
    assert root is not None
    lifted = anchor(root, "S", as_of=_AS_OF)
    assert lifted == 'IF((DATE(2000,1,1)-DATE(2000,1,1))>S!$A$1,"TODAY()",DATE(2000,1,1))'
    assert anchor(root, "S") == 'IF((TODAY()-NOW())>S!$A$1,"TODAY()",now())'


# -- T14: & binds tighter than comparisons -------------------------------------------


@pytest.mark.parametrize(
    ("text", "lifted"),
    [
        ('A1&B1="MN"', '(S!$A$1&S!$B$1)="MN"'),
        ('A1&B1<>"MN"', '(S!$A$1&S!$B$1)<>"MN"'),
        ("A1=B1&C1", "S!$A$1=(S!$B$1&S!$C$1)"),
        ('(A1&B1)="MN"', '(S!$A$1&S!$B$1)="MN"'),
        ('"a"&1+2', '"a"&(1+2)'),
        ("A1+B1=C1", "(S!$A$1+S!$B$1)=S!$C$1"),
        ("A1*B1>C1", "(S!$A$1*S!$B$1)>S!$C$1"),
        ("-2^2", "-2^2"),
        ("2^3^2", "(2^3)^2"),
        ("1-2-3", "(1-2)-3"),
        ("A1%>B1", "S!$A$1%>S!$B$1"),
    ],
)
def test_lifted_text_keeps_excel_precedence(text: str, lifted: str) -> None:
    assert _lift(text) == lifted


def test_concatenated_condition_reads_its_real_outcome(make_book: Book, tmp_path: Path) -> None:
    # Sex and smoker stored "M","N", so "MN"="MN" is TRUE. Sum assured 100000.
    #   B2 IF(sex&smk="MN", SA*0.001, SA*0.002)     stored 100
    #   B3 IF(sex&smk<>"MN", SA*0.002, SA*0.001)    stored 100 (FALSE arm)
    #   B4 IF(sex=smk&"", 1, 2)  -> "M"=("N"&"") FALSE  stored 2
    #   B5 IF((sex&smk)="MN", SA*0.001, SA*0.002)    the same test, parenthesised, in the
    #      same book (one probe text for both spellings)                 stored 100
    #   B6 IFS(sex&smk="MN", 1, TRUE, 2)            stored 1
    # Policy F,N: "FN"="MN" FALSE -> B2 200, B3 200, B4 2 ("F"="N"), B5 200, B6 2: B2, B3,
    # B5, B6 take an arm the trust run never took; B4 takes the stored one.
    # Policy N,N: B4 "N"=("N"&"") TRUE -> 1, a new arm.
    sa = "Inputs!B4"
    cond = 'Inputs!B2&Inputs!B3="MN"'
    book = make_book(
        "prec",
        {
            "Inputs": {"B2": "M", "B3": "N", "B4": 100000},
            "Calc": {
                "B2": F(f"=IF({cond},{sa}*0.001,{sa}*0.002)", 100.0),
                "B3": F(f'=IF(Inputs!B2&Inputs!B3<>"MN",{sa}*0.002,{sa}*0.001)', 100.0),
                "B4": F('=IF(Inputs!B2=Inputs!B3&"",1,2)', 2.0),
                "B5": F(f'=IF((Inputs!B2&Inputs!B3)="MN",{sa}*0.001,{sa}*0.002)', 100.0),
                "B6": F(f"=_xlfn.IFS({cond},1,TRUE,2)", 1.0),
                # CHOOSE(1+TRUE) is the second choice: stored 100; F,N takes the first.
                "B7": F(f"=CHOOSE(1+({cond}),{sa}*0.002,{sa}*0.001)", 100.0),
            },
        },
    )
    outputs = [f"Calc!B{r}" for r in range(2, 8)]
    res = _run(
        tmp_path,
        book,
        "id,sex,smk\nMN,M,N\nFN,F,N\nNN,N,N\n",
        {"sex": "Inputs!B2", "smk": "Inputs!B3"},
        outputs,
    )
    rows, paths = _rows(res), _paths(res)
    for out in outputs:
        assert rows[("MN", out)]["status"] == "ok", out
    assert paths[("MN", "Calc!B2")] == "TRUE"
    assert paths[("MN", "Calc!B3")] == "FALSE"
    assert paths[("MN", "Calc!B4")] == "FALSE"
    assert paths[("MN", "Calc!B5")] == "TRUE"
    assert paths[("FN", "Calc!B2")] == "FALSE"
    assert paths[("FN", "Calc!B5")] == "FALSE"
    for out, value in (("Calc!B2", "200.0"), ("Calc!B3", "200.0"), ("Calc!B5", "200.0")):
        assert rows[("FN", out)]["value"] == value
        assert rows[("FN", out)]["status"] == "unverified_path", out
    assert rows[("FN", "Calc!B6")]["value"] == "2.0"
    assert rows[("FN", "Calc!B6")]["status"] == "unverified_path"
    assert paths[("MN", "Calc!B7")] == "2"
    assert rows[("FN", "Calc!B7")]["value"] == "200.0"
    assert rows[("FN", "Calc!B7")]["status"] == "unverified_path"
    assert "CHOOSE #0 took '1'" in rows[("FN", "Calc!B7")]["reasons"]
    assert rows[("FN", "Calc!B4")]["status"] == "ok"
    assert rows[("NN", "Calc!B4")]["value"] == "1.0"
    assert rows[("NN", "Calc!B4")]["status"] == "unverified_path"


# -- T19: sheet names that must be quoted ---------------------------------------------


@pytest.mark.parametrize(
    ("sheet", "written"),
    [
        ("Plain", "Plain"),
        ("Q1_x", "Q1_x"),
        ("2024", "'2024'"),
        ("1Q", "'1Q'"),
        ("3Yr", "'3Yr'"),
        ("A1", "'A1'"),
        ("xfd9", "'xfd9'"),
        ("R1C1", "'R1C1'"),
        ("R", "'R'"),
        ("c", "'c'"),
        ("RC", "'RC'"),
        ("R2C", "'R2C'"),
        ("RC3", "'RC3'"),
        ("TRUE", "'TRUE'"),
        ("false", "'false'"),
        ("My Sheet", "'My Sheet'"),
        ("It's", "'It''s'"),
    ],
)
def test_quote_sheet(sheet: str, written: str) -> None:
    assert quote_sheet(sheet) == written


@pytest.mark.parametrize("sheet", ["2024", "1Q", "3Yr", "A1", "R1C1", "TRUE"])
def test_probes_on_sheets_that_need_quotes_compute(
    make_book: Book, tmp_path: Path, sheet: str
) -> None:
    # Stored B1 = 1: C2 = IF(B1>5, B1*2, 0) took FALSE (0). D2 = INDIRECT("A"&B1) read
    # A1 (10). Policy B1 = 3: C2 FALSE again (0, the stored path); D2 reads A3 (30).
    # Policy B1 = 7: C2 TRUE (14), an arm the trust run never took.
    q = quote_sheet(sheet)
    book = make_book(
        f"s{sheet}",
        {
            sheet: {
                "A1": 10,
                "A2": 20,
                "A3": 30,
                "B1": 1,
                "C2": F("=IF(B1>5,B1*2,0)", 0.0),
                "D2": F('=INDIRECT("A"&B1)', 10.0),
            },
            # Another sheet reading C2: 1 stored, 1 for P3 (C2 = 0 again).
            "Other": {"B1": F(f"={q}!C2+1", 1.0)},
        },
    )
    res = _run(
        tmp_path,
        book,
        "id,x\nP3,3\nP7,7\n",
        {"x": f"{q}!B1"},
        [f"{q}!C2", f"{q}!D2", "Other!B1"],
    )
    rows = {(pid, out.rpartition("!")[2]): r for (pid, out), r in _rows(res).items()}
    assert rows[("P3", "C2")]["status"] == "ok", rows[("P3", "C2")]
    assert rows[("P3", "D2")]["value"] == "30.0"
    assert rows[("P3", "D2")]["status"] == "ok", rows[("P3", "D2")]
    assert rows[("P7", "C2")]["value"] == "14.0"
    assert rows[("P7", "C2")]["status"] == "unverified_path"
    assert "engine_error" not in rows[("P7", "C2")]["reasons"]
    assert rows[("P3", "B1")]["status"] == "ok", rows[("P3", "B1")]
