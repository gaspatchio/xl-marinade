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

from tests.calc.conftest import F  # noqa: E402
from xl_marinade.calc import CalcRequest, CalcResult, calc  # noqa: E402
from xl_marinade.calc.engine import functions_engine_lacks  # noqa: E402
from xl_marinade.calc.formula_ast import ENGINE_LACKS  # noqa: E402
from xl_marinade.calc.models import CellRef  # noqa: E402
from xl_marinade.calc.probes import _text_cells  # noqa: E402

Book = Callable[..., Path]

# Excel's worksheet functions, written out so the survey does not move with the
# openpyxl version: the 2007 set, the ``_xlfn.`` functions, dynamic arrays and later,
# and the add-in and legacy ones (524 names).
_EXCEL_FUNCTIONS = """
ABS ACCRINT ACCRINTM ACOS ACOSH ACOT ACOTH ADDRESS AGGREGATE AMORDEGRC AMORLINC AND ARABIC
AREAS ARRAYTOTEXT ASC ASIN ASINH ATAN ATAN2 ATANH AVEDEV AVERAGE AVERAGEA AVERAGEIF
AVERAGEIFS BAHTTEXT BASE BESSELI BESSELJ BESSELK BESSELY BETA.DIST BETA.INV BETADIST BETAINV
BIN2DEC BIN2HEX BIN2OCT BINOM.DIST BINOM.DIST.RANGE BINOM.INV BINOMDIST BITAND BITLSHIFT
BITOR BITRSHIFT BITXOR BYCOL BYROW CALL CEILING CEILING.MATH CEILING.PRECISE CELL CHAR
CHIDIST CHIINV CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHITEST CHOOSE
CHOOSECOLS CHOOSEROWS CLEAN CODE COLUMN COLUMNS COMBIN COMBINA COMPLEX CONCAT CONCATENATE
CONFIDENCE CONFIDENCE.NORM CONFIDENCE.T CONVERT CORREL COS COSH COT COTH COUNT COUNTA
COUNTBLANK COUNTIF COUNTIFS COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD COVAR
COVARIANCE.P COVARIANCE.S CRITBINOM CSC CSCH CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY
CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE CUMIPMT CUMPRINC DATE DATEDIF DATEVALUE
DAVERAGE DAY DAYS DAYS360 DB DBCS DCOUNT DCOUNTA DDB DEC2BIN DEC2HEX DEC2OCT DECIMAL DEGREES
DELTA DETECTLANGUAGE DEVSQ DGET DISC DMAX DMIN DOLLAR DOLLARDE DOLLARFR DPRODUCT DROP DSTDEV
DSTDEVP DSUM DURATION DVAR DVARP ECMA.CEILING EDATE EFFECT ENCODEURL EOMONTH ERF ERF.PRECISE
ERFC ERFC.PRECISE ERROR.TYPE EUROCONVERT EVEN EXACT EXP EXPAND EXPON.DIST EXPONDIST F.DIST
F.DIST.RT F.INV F.INV.RT F.TEST FACT FACTDOUBLE FALSE FDIST FIELDVALUE FILTER FILTERXML FIND
FINDB FINV FISHER FISHERINV FIXED FLOOR FLOOR.MATH FLOOR.PRECISE FORECAST FORECAST.ETS
FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT
FREQUENCY FTEST FV FVSCHEDULE GAMMA GAMMA.DIST GAMMA.INV GAMMADIST GAMMAINV GAMMALN
GAMMALN.PRECISE GAUSS GCD GEOMEAN GESTEP GETPIVOTDATA GROUPBY GROWTH HARMEAN HEX2BIN HEX2DEC
HEX2OCT HLOOKUP HOUR HSTACK HYPERLINK HYPGEOM.DIST HYPGEOMDIST IF IFERROR IFNA IFS IMABS
IMAGE IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMCOSH IMCOT IMCSC IMCSCH IMDIV IMEXP IMLN
IMLOG10 IMLOG2 IMPOWER IMPRODUCT IMREAL IMSEC IMSECH IMSIN IMSINH IMSQRT IMSUB IMSUM IMTAN
INDEX INDIRECT INFO INT INTERCEPT INTRATE IPMT IRR ISBLANK ISERR ISERROR ISEVEN ISFORMULA
ISLOGICAL ISNA ISNONTEXT ISNUMBER ISO.CEILING ISODD ISOMITTED ISOWEEKNUM ISPMT ISREF ISTEXT
JIS KURT LAMBDA LARGE LCM LEFT LEFTB LEN LENB LET LINEST LN LOG LOG10 LOGEST LOGINV
LOGNORM.DIST LOGNORM.INV LOGNORMDIST LOOKUP LOWER MAKEARRAY MAP MATCH MAX MAXA MAXIFS
MDETERM MDURATION MEDIAN MID MIDB MIN MINA MINIFS MINUTE MINVERSE MIRR MMULT MOD MODE
MODE.MULT MODE.SNGL MONTH MROUND MULTINOMIAL MUNIT N NA NEGBINOM.DIST NEGBINOMDIST
NETWORKDAYS NETWORKDAYS.INTL NOMINAL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NORMDIST
NORMINV NORMSDIST NORMSINV NOT NOW NPER NPV NUMBERVALUE OCT2BIN OCT2DEC OCT2HEX ODD
ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD OFFSET OR PDURATION PEARSON PERCENTILE
PERCENTILE.EXC PERCENTILE.INC PERCENTOF PERCENTRANK PERCENTRANK.EXC PERCENTRANK.INC PERMUT
PERMUTATIONA PHI PHONETIC PI PIVOTBY PMT POISSON POISSON.DIST POWER PPMT PRICE PRICEDISC
PRICEMAT PROB PRODUCT PROPER PV QUARTILE QUARTILE.EXC QUARTILE.INC QUOTIENT RADIANS RAND
RANDARRAY RANDBETWEEN RANK RANK.AVG RANK.EQ RATE RECEIVED REDUCE REGEXEXTRACT REGEXREPLACE
REGEXTEST REGISTER.ID REPLACE REPLACEB REPT RIGHT RIGHTB ROMAN ROUND ROUNDDOWN ROUNDUP ROW
ROWS RRI RSQ RTD SCAN SEARCH SEARCHB SEC SECH SECOND SEQUENCE SERIESSUM SHEET SHEETS SIGN
SIN SINH SKEW SKEW.P SLN SLOPE SMALL SORT SORTBY SQL.REQUEST SQRT SQRTPI STANDARDIZE STDEV
STDEV.P STDEV.S STDEVA STDEVP STDEVPA STEYX STOCKHISTORY SUBSTITUTE SUBTOTAL SUM SUMIF
SUMIFS SUMPRODUCT SUMSQ SUMX2MY2 SUMX2PY2 SUMXMY2 SWITCH SYD T T.DIST T.DIST.2T T.DIST.RT
T.INV T.INV.2T T.TEST TAKE TAN TANH TBILLEQ TBILLPRICE TBILLYIELD TDIST TEXT TEXTAFTER
TEXTBEFORE TEXTJOIN TEXTSPLIT TIME TIMEVALUE TINV TOCOL TODAY TOROW TRANSLATE TRANSPOSE
TREND TRIM TRIMMEAN TRIMRANGE TRUE TRUNC TTEST TYPE UNICHAR UNICODE UNIQUE UPPER VALUE
VALUETOTEXT VAR VAR.P VAR.S VARA VARP VARPA VDB VLOOKUP VSTACK WEBSERVICE WEEKDAY WEEKNUM
WEIBULL WEIBULL.DIST WORKDAY WORKDAY.INTL WRAPCOLS WRAPROWS XIRR XLOOKUP XMATCH XNPV XOR
YEAR YEARFRAC YIELD YIELDDISC YIELDMAT Z.TEST ZTEST
"""
EXCEL_FUNCTIONS = frozenset(_EXCEL_FUNCTIONS.split())


def test_engine_lacks_is_what_the_installed_ironcalc_lacks() -> None:
    # A function the engine knows answers a bare call with an argument error
    # (#ERROR!); one it does not know answers #NAME?.
    names = sorted(EXCEL_FUNCTIONS)
    assert len(names) == 524
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


def test_unknown_function_is_refused_by_name_on_every_path(make_book: Book, tmp_path: Path) -> None:
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
    # The installed engine is asked about every function the workbook calls, so the
    # unknown one is refused by name on every path, the stored one included.
    for pid in ("P1", "P7"):
        row = rows[(pid, "Calc!B1")]
        assert row["status"] == "refused", (pid, row)
        assert (
            "refused_pattern@Calc!B1:function IronCalc does not implement: NOTAFUNCTIONYET"
            in row["reasons"]
        )
    assert rows[("P7", "Calc!B1")]["value"] == "#NAME?"


# -- functions the installed engine lacks, asked at run time ----------------------


def test_engine_reports_the_functions_it_lacks() -> None:
    # A bare call answers an argument error for a function IronCalc implements,
    # #NAME? for one it does not know, and #N/IMPL! for one it knows but has not
    # implemented. Prefixes are dropped; LET/LAMBDA parameters (_xlpm.) are not asked.
    asked = {
        "SUM",
        "_xlfn.XLOOKUP",
        "MYUDF",
        "_xll.ADDIN",
        "_xlfn.FUTUREFN",
        "_xlfn.FORECAST.ETS",
        "_xlfn.AGGREGATE",
        "SHEETS",
        "_xlpm.f",
    }
    assert functions_engine_lacks(asked) == frozenset(
        {"MYUDF", "ADDIN", "FUTUREFN", "FORECAST.ETS", "AGGREGATE"}
    )
    assert functions_engine_lacks(set()) == frozenset()


@pytest.mark.parametrize(
    ("call", "name"),
    [
        ("MYUDF(Inputs!B1)", "MYUDF"),  # a VBA function
        ("_xll.ADDIN(Inputs!B1)", "ADDIN"),  # an add-in function
        ("_xlfn.FUTUREFN(Inputs!B1)", "FUTUREFN"),  # newer than the engine
        ("_xlfn.FORECAST.ETS(4,Inputs!B2:B4,Inputs!C2:C4)", "FORECAST.ETS"),  # #N/IMPL!
    ],
)
def test_a_function_the_engine_lacks_is_refused_even_behind_an_error_handler(
    make_book: Book, tmp_path: Path, call: str, name: str
) -> None:
    # IronCalc answers #NAME? or #N/IMPL!; IFERROR and IF(ISERROR()) turn either into
    # the fallback 0 on every input, a plausible value no trust run can tell apart
    # from Excel's when the stored inputs take the other arm. x = 1 keeps the stored
    # arm (2); x = 7 takes the handler.
    book = make_book(
        "lack",
        {
            "Inputs": {"B1": 1, "B2": 1, "B3": 2, "B4": 3, "C2": 10, "C3": 20, "C4": 30},
            "Calc": {
                "B1": F(f"=IF(Inputs!B1>5,IFERROR({call},0),2)", 2.0),
                "B2": F(f"=IF(Inputs!B1>5,IF(ISERROR({call}),0,1),2)", 2.0),
            },
        },
    )
    res = _run(tmp_path, book, "id,x\nP1,1\nP7,7\n", {"x": "Inputs!B1"}, ["Calc!B1", "Calc!B2"])
    rows = _rows(res)
    for pid in ("P1", "P7"):
        for out in ("Calc!B1", "Calc!B2"):
            row = rows[(pid, out)]
            assert row["status"] == "refused", (pid, out, row)
            assert (
                f"refused_pattern@{out}:function IronCalc does not implement: {name}"
                in row["reasons"]
            )


def test_sheets_of_a_reference_is_refused(make_book: Book, tmp_path: Path) -> None:
    # Excel's SHEETS(Inputs!B1) is 1, so B1 = 1; IronCalc answers #N/IMPL! for SHEETS
    # with an argument, which ISERROR turns into 0. SHEETS() with no argument works.
    book = make_book(
        "sh",
        {
            "Inputs": {"B1": 1},
            "Calc": {
                "B1": F("=IF(ISERROR(_xlfn.SHEETS(Inputs!B1)),0,1)", 1.0),
                "B2": F("=_xlfn.SHEETS()+Inputs!B1", 3.0),
            },
        },
    )
    res = _run(tmp_path, book, "id,x\nP1,1\n", {"x": "Inputs!B1"}, ["Calc!B1", "Calc!B2"])
    rows = _rows(res)
    assert rows[("P1", "Calc!B1")]["status"] == "refused"
    assert (
        "refused_pattern@Calc!B1:function IronCalc does not implement: SHEETS with a reference"
        in rows[("P1", "Calc!B1")]["reasons"]
    )
    assert "SHEETS" not in rows[("P1", "Calc!B2")]["reasons"]


def test_a_let_bound_lambda_is_not_mistaken_for_an_unknown_function(
    make_book: Book, tmp_path: Path
) -> None:
    # LET(f, LAMBDA(x, x+1), f(B1)): Excel stores f as _xlpm.f, which a bare call
    # outside the LET does not know. For x = 1 the result is 2, for x = 4 it is 5.
    book = make_book(
        "let",
        {
            "Inputs": {"B1": 1},
            "Calc": {
                "B1": F(
                    "=_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlpm.f(Inputs!B1))",
                    2.0,
                )
            },
        },
    )
    res = _run(tmp_path, book, "id,x\nP4,4\n", {"x": "Inputs!B1"}, ["Calc!B1"])
    row = _rows(res)[("P4", "Calc!B1")]
    assert "does not implement" not in row["reasons"], row
    assert row["value"] == "5.0"


# -- T07: INDIRECT text off the grid ---------------------------------------------


@pytest.mark.parametrize(
    "text",
    [
        "Rates!B0",
        "Rates!B0:B2",
        "B0",
        "Rates!B-1",
        "Rates!XFE1",
        "Rates!B1048577",
        "Rates!A1:B1048577",
    ],
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
