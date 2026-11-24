"""Engine gaps the trust run cannot see: functions IronCalc lacks, and INDIRECT text
that addresses a cell off the grid.

Hand values come from Excel's documented semantics: ``AGGREGATE(14,6,r,k)`` is
``LARGE(r,k)`` ignoring errors (``#NUM!`` when ``k`` exceeds the count), and an
``INDIRECT`` address on row 0, or past row 1,048,576 or column XFD, is ``#REF!``.
"""

from __future__ import annotations

import csv
import json
from collections.abc import Callable
from pathlib import Path

import pytest

ironcalc = pytest.importorskip("ironcalc")

from openpyxl.utils.formulas import FORMULAE  # noqa: E402

from tests.calc.conftest import F  # noqa: E402
from xl_marinade.calc import CalcRequest, CalcResult, calc  # noqa: E402
from xl_marinade.calc.formula_ast import ENGINE_LACKS  # noqa: E402
from xl_marinade.calc.models import CellRef  # noqa: E402
from xl_marinade.calc.probes import _text_cells  # noqa: E402

Book = Callable[..., Path]

# Excel's worksheet functions added after the 2007 list openpyxl carries (the ``_xlfn.``
# functions, dynamic arrays and later), plus the add-in and legacy ones.
_EXCEL_SINCE_2007 = """
ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE
BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE
CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT
CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL
DETECTLANGUAGE DROP ECMA.CEILING ENCODEURL ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST
F.DIST.RT F.INV F.INV.RT F.TEST FIELDVALUE FILTER FILTERXML FLOOR.MATH FLOOR.PRECISE
FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR
FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS GROUPBY HSTACK HYPGEOM.DIST IFNA
IFS IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISO.CEILING ISOMITTED
ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL
MUNIT NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE
PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTOF PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA
PHI PIVOTBY POISSON.DIST QUARTILE.EXC QUARTILE.INC RANDARRAY RANK.AVG RANK.EQ REDUCE REGEXEXTRACT
REGEXREPLACE REGEXTEST RRI SCAN SEC SECH SEQUENCE SHEET SHEETS SKEW.P SORT SORTBY STDEV.P STDEV.S
STOCKHISTORY SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE
TEXTJOIN TEXTSPLIT TOCOL TOROW TRANSLATE TRIMRANGE UNICHAR UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S
VSTACK WEBSERVICE WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST
BAHTTEXT CALL DBCS EUROCONVERT GETPIVOTDATA HYPERLINK INFO JIS PHONETIC REGISTER.ID RTD
SQL.REQUEST
""".split()
EXCEL_FUNCTIONS = frozenset(FORMULAE) | frozenset(_EXCEL_SINCE_2007)


def test_engine_lacks_is_what_the_installed_ironcalc_lacks() -> None:
    # A function the engine knows answers a bare call with an argument error
    # (#ERROR!); one it does not know answers #NAME?.
    names = sorted(EXCEL_FUNCTIONS)
    model = ironcalc.create("survey")
    for row, name in enumerate(names, 1):
        model.update_cell_with_formula(0, row, 1, f"={name}()")
    model.evaluate()
    lacking = {
        name for row, name in enumerate(names, 1) if model.get_cell_value(0, row, 1) == "#NAME?"
    }
    assert lacking == ENGINE_LACKS
    assert {"AGGREGATE", "BAHTTEXT"} <= ENGINE_LACKS


def _run(
    tmp: Path, book: Path, policies: str, fields: dict[str, str], outputs: list[str]
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
        )
    )


def _rows(res: CalcResult) -> dict[tuple[str, str], dict[str, str]]:
    with res.results_csv.open(encoding="utf-8", newline="") as fh:
        return {(r["policy_id"], r["output"]): r for r in csv.DictReader(fh)}


def _patterns(res: CalcResult) -> dict[str, list[str]]:
    report = json.loads(res.trust_report.read_text(encoding="utf-8"))
    out: dict[str, list[str]] = {}
    for entry in report["refused_patterns"]:
        out.setdefault(entry["cell"], []).append(entry["detail"])
    return out


_LACKS_AGGREGATE = "refused_pattern@{cell}:function IronCalc does not implement: AGGREGATE"


def test_masked_aggregate_is_refused_by_name(make_book: Book, tmp_path: Path) -> None:
    # Claims 300, 900, 500. Stored k = 4: Excel's AGGREGATE is #NUM! (k > count), so each
    # error handler takes its error arm, which IronCalc's #NAME? takes too. For k = 1
    # Excel gives LARGE = 900: B2 = 900, B3 = 1, B4 = 1, B5 = 1 (ISERR is TRUE only for
    # errors other than #N/A), B6 = "x" & BAHTTEXT(1) (a Thai text) -- IronCalc gives 0,
    # 0, 0, 0 and "none". B7 counts the numbers among B2 and AGGREGATE's own cell C1.
    book = make_book(
        "agg",
        {
            "Inputs": {"B2": 300, "B3": 900, "B4": 500, "B5": 4},
            "Calc": {
                "B2": F("=IFERROR(_xlfn.AGGREGATE(14,6,Inputs!B2:B4,Inputs!B5),0)", 0.0),
                "B3": F("=IF(ISERROR(_xlfn.AGGREGATE(14,6,Inputs!B2:B4,Inputs!B5)),0,1)", 0.0),
                "B4": F("=IF(ISNUMBER(_xlfn.AGGREGATE(14,6,Inputs!B2:B4,Inputs!B5)),1,0)", 0.0),
                "B5": F("=IF(ISERR(_xlfn.AGGREGATE(14,6,Inputs!B2:B4,Inputs!B5)),0,1)", 0.0),
                "B6": F('=IF(Inputs!B5>3,"none",BAHTTEXT(Inputs!B5))', "none"),
                "C1": F("=_xlfn.AGGREGATE(14,6,Inputs!B2:B4,Inputs!B5)", "#NUM!"),
                "B7": F("=COUNT(B2,C1)", 1.0),
            },
        },
    )
    outputs = [f"Calc!B{r}" for r in range(2, 8)]
    res = _run(tmp_path, book, "id,k\nP1,1\n", {"k": "Inputs!B5"}, outputs)
    rows = _rows(res)
    for out in outputs:
        row = rows[("P1", out)]
        assert row["status"] == "refused", (out, row)
    for out in ("Calc!B2", "Calc!B3", "Calc!B4", "Calc!B5"):
        assert _LACKS_AGGREGATE.format(cell=out) in rows[("P1", out)]["reasons"]
    for cell in ("Calc!B2", "Calc!B3", "Calc!B4", "Calc!B5", "Calc!C1"):
        assert _patterns(res)[cell] == ["function IronCalc does not implement: AGGREGATE"]
    assert _patterns(res)["Calc!B6"] == ["function IronCalc does not implement: BAHTTEXT"]
    assert (
        "refused_pattern@Calc!C1:function IronCalc does not implement"
        in (rows[("P1", "Calc!B7")]["reasons"])
    )


def test_xlookup_not_found_arm_holding_aggregate_is_refused(
    make_book: Book, tmp_path: Path
) -> None:
    # Ages 20..22 with rates; age 99 is not in the table, so Excel evaluates the
    # if_not_found arm: AGGREGATE(4,6,rates) = MAX ignoring errors = 0.03. IronCalc: #NAME?.
    book = make_book(
        "xl",
        {
            "Inputs": {"B1": 21},
            "Rates": {"A1": 20, "A2": 21, "A3": 22, "B1": 0.01, "B2": 0.02, "B3": 0.03},
            "Calc": {
                "B1": F(
                    "=_xlfn.XLOOKUP(Inputs!B1,Rates!A1:A3,Rates!B1:B3,"
                    "_xlfn.AGGREGATE(4,6,Rates!B1:B3))",
                    0.02,
                )
            },
        },
    )
    res = _run(tmp_path, book, "id,age\nP99,99\n", {"age": "Inputs!B1"}, ["Calc!B1"])
    row = _rows(res)[("P99", "Calc!B1")]
    assert row["status"] == "refused"
    assert _LACKS_AGGREGATE.format(cell="Calc!B1") in row["reasons"]


def test_name_error_off_the_stored_path_is_an_engine_error(make_book: Book, tmp_path: Path) -> None:
    # A function the engine does not know (here one no list could hold) answers #NAME?
    # only on the arm the stored inputs did not take; the stored value 2 reproduces.
    # For x = 7 the cell evaluates the unknown function: whatever Excel makes of it,
    # the engine's #NAME? is not evidence, so the value must not come back ok.
    book = make_book(
        "nf",
        {
            "Inputs": {"B1": 1},
            "Calc": {"B1": F("=IF(Inputs!B1>5,NOTAFUNCTIONYET(Inputs!B1),2)", 2.0)},
        },
    )
    res = _run(tmp_path, book, "id,x\nP1,1\nP7,7\n", {"x": "Inputs!B1"}, ["Calc!B1"])
    rows = _rows(res)
    assert rows[("P1", "Calc!B1")]["status"] in {"ok", "unverified_path"}
    row = rows[("P7", "Calc!B1")]
    assert row["value"] == "#NAME?"
    assert row["status"] == "refused"
    assert "engine_error@Calc!B1:" in row["reasons"]


# -- T07: INDIRECT text off the grid ---------------------------------------------


@pytest.mark.parametrize(
    "text",
    ["Rates!B0", "Rates!B0:B2", "B0", "Rates!XFE1", "Rates!B1048577", "Rates!A1:B1048577"],
)
def test_text_address_off_the_grid_is_unread(text: str) -> None:
    assert _text_cells(text, "Calc", {"rates": "Rates", "calc": "Calc"}) is None


def test_text_address_on_the_grid_edges_is_read() -> None:
    sheets = {"rates": "Rates"}
    assert _text_cells("Rates!XFD1048576", "Calc", sheets) == [CellRef("Rates", 1_048_576, 16_384)]
    assert _text_cells("rates!A1", "Calc", sheets) == [CellRef("Rates", 1, 1)]


def test_indirect_on_row_zero_is_refused(make_book: Book, tmp_path: Path) -> None:
    # Rates for ages 20..22 in Rates!B1:B3; the row is age - 19. Age 19 addresses row 0:
    # Excel's INDIRECT is #REF! there, so B2 and B3 are #REF!; B4 SUM over B0:B1 is #REF!;
    # ROWS over it is #REF!, ISREF FALSE. Ages 20..22 read rows 1..3.
    book = make_book(
        "ind",
        {
            "Inputs": {"B1": 21, "B2": 250000},
            "Rates": {"B1": 0.001, "B2": 0.002, "B3": 0.003},
            "Calc": {
                "B2": F('=INDIRECT("Rates!B"&(Inputs!B1-19))', 0.002),
                "B3": F('=Inputs!B2*INDIRECT("Rates!B"&(Inputs!B1-19))', 500.0),
                "B4": F('=SUM(INDIRECT("Rates!B"&(Inputs!B1-20)&":B"&(Inputs!B1-19)))', 0.003),
                "B5": F('=ROWS(INDIRECT("Rates!B"&(Inputs!B1-19)))', 1.0),
                "B6": F('=ISREF(INDIRECT("Rates!B"&(Inputs!B1-19)))', True),
            },
        },
    )
    outputs = [f"Calc!B{r}" for r in range(2, 7)]
    res = _run(tmp_path, book, "id,age\nA19,19\nA20,20\nA21,21\n", {"age": "Inputs!B1"}, outputs)
    rows = _rows(res)
    for out in outputs:
        row = rows[("A19", out)]
        assert row["status"] == "refused", (out, row)
        assert f"refused_pattern@{out}:dynamic reference not checked" in row["reasons"]
    # Age 20 reads row 1 (B2..B3, B5, B6) but B4's range starts on row 0.
    assert rows[("A20", "Calc!B2")]["value"] == "0.001"
    assert rows[("A20", "Calc!B2")]["status"] == "ok"
    assert rows[("A20", "Calc!B4")]["status"] == "refused"
    assert rows[("A21", "Calc!B4")]["status"] == "ok"
