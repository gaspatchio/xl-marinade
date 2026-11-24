"""Number-to-text conversion is refused per policy where the engine's text can differ
from Excel's (ruling D6).

Hand values from Excel's documented behaviour: ``TEXT``, ``FIXED`` and ``DOLLAR`` round
half away from zero on the 15-digit decimal value (``TEXT(1234.5,"#,##0")`` is
``"1,235"``, ``FIXED(1.005,2)`` is ``"1.01"``, ``TEXT(0.0125,"0.0%")`` is ``"1.3%"``, a
trailing comma divides by 1,000); ``&``, ``CONCAT``, ``CONCATENATE``, ``TEXTJOIN`` and
the text functions write a number with at most 15 significant digits
(``"t="&(7/12)`` is ``"t=0.583333333333333"``, ``LEN(7/12)`` is 17). IronCalc 0.9.1
rounds a binary tie to even, cannot see a 15-digit tie and writes up to 17 digits.
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
from xl_marinade.calc.probes import excel_general_text, number_format_rounding  # noqa: E402

Book = Callable[..., Path]

TIE = "number-to-text at a decimal tie"
DIGITS = "number-to-text where Excel's 15-digit text differs"
CRITERION = "criterion built with & from a non-integer number"
NOT_CHECKED = "number-to-text not checked"


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


def _refused(row: dict[str, str], cell: str, detail: str) -> None:
    assert row["status"] == "refused", row
    assert f"refused_pattern@{cell}:{detail}" in row["reasons"], row


# Inputs!B1 is x, stored 1234.4 (no tie anywhere below).
_TEXT_CELLS: dict[str, object] = {
    "B1": F('=TEXT(Inputs!B1,"#,##0")', "1,234"),
    "B2": F('="Annual premium "&TEXT(Inputs!B1,"#,##0")', "Annual premium 1,234"),
    "B3": F('=TEXT(Inputs!B1/100000,"0.0%")', "1.2%"),
    "B4": F('=VALUE(TEXT(Inputs!B1/1000,"0.00"))', 1.23),
    "B5": F("=FIXED(Inputs!B1/1000,2)", "1.23"),
    "B6": F("=DOLLAR(Inputs!B1/1000,2)", "$1.23"),
    "B7": F('=TEXT(Inputs!B1*1000,"#,##0,")', "1,234"),
    "B8": F('=TEXT(Inputs!B1-2000,"0;(0)")', "(766)"),
    "B9": F('=TEXT(Inputs!B1,"yyyy-mm-dd")', "1903-05-18"),
    "B10": F('=TEXT(Inputs!B1,"0.0E+0")', "1.2E+3"),
    "B11": F("=TEXT(Inputs!B1,Inputs!C1)", "1234"),
    "B12": F('=IF(Inputs!B1>100,TEXT(Inputs!B1,"0"),"small")', "1234"),
    "B13": F("=FIXED(Inputs!B1,0)", "1,234"),
}
_TEXT_OUTPUTS = [f"Calc!B{r}" for r in range(1, 14)]


def test_text_fixed_dollar_at_a_decimal_tie_are_refused(make_book: Book, tmp_path: Path) -> None:
    book = make_book(
        "tie",
        {"Inputs": {"B1": 1234.4, "C1": "0"}, "Calc": dict(_TEXT_CELLS)},
    )
    # T1 1234.5: B1, B2 "1,235" (tie at 0 decimals); B3 1.2345% at one decimal, no tie;
    #   B4 1.2345 at two decimals, no tie; B5, B6 no tie; B7 1234500/1000 = 1234.5, a tie;
    #   B8 -765.5 a tie in the negative section; B12 "1235" a tie; B13 FIXED(1234.5,0) a
    #   tie (an exact binary one, which IronCalc happens to round right: over-refused).
    # T2 1250: B3 1.25% at one decimal is a tie ("1.3%"); B4 1.25 at two decimals, no.
    # T3 1005: B4 TEXT(1.005,"0.00") is a 15-digit tie ("1.01"); B5 FIXED(1.005,2), B6
    #   DOLLAR(1.005,2) too.
    # T4 2.5: B12 takes "small" -- its TEXT is never evaluated, so no tie there.
    # I 1234: no tie anywhere.
    res = _run(
        tmp_path,
        book,
        "id,x\nT1,1234.5\nT2,1250\nT3,1005\nT4,2.5\nI,1234\n",
        {"x": "Inputs!B1"},
        _TEXT_OUTPUTS,
    )
    rows = _rows(res)
    for out in ("Calc!B1", "Calc!B2", "Calc!B7", "Calc!B8", "Calc!B12", "Calc!B13"):
        _refused(rows[("T1", out)], out, TIE)
    for out in ("Calc!B3", "Calc!B4", "Calc!B5", "Calc!B6"):
        assert TIE not in rows[("T1", out)]["reasons"], out
    _refused(rows[("T2", "Calc!B3")], "Calc!B3", TIE)
    assert TIE not in rows[("T2", "Calc!B4")]["reasons"]
    for out in ("Calc!B4", "Calc!B5", "Calc!B6"):
        _refused(rows[("T3", out)], out, TIE)
    assert rows[("T4", "Calc!B12")]["value"] == "small"
    assert rows[("T4", "Calc!B12")]["status"] == "unverified_path"
    # Date formats need no check; integers and non-ties stay ok.
    assert rows[("T1", "Calc!B9")]["status"] == "ok"
    for out in ("Calc!B1", "Calc!B2", "Calc!B3", "Calc!B4", "Calc!B5", "Calc!B6", "Calc!B13"):
        assert rows[("I", out)]["status"] == "ok", (out, rows[("I", out)])
    # A format the scan cannot read, or one held in a cell, is not checkable.
    for pid in ("T1", "I"):
        _refused(rows[(pid, "Calc!B10")], "Calc!B10", NOT_CHECKED)
        _refused(rows[(pid, "Calc!B11")], "Calc!B11", NOT_CHECKED)


def test_general_conversion_beyond_15_digits_is_refused(make_book: Book, tmp_path: Path) -> None:
    # Inputs!B1 = months, stored 6, so B1/12 = 0.5 prints the same everywhere.
    # Policy M7: 7/12 = 0.58333333333333337 -- Excel writes 0.583333333333333, IronCalc
    # 0.5833333333333334. Policy M3: 0.25. Inputs!B2 holds an integer count (stored 3).
    x = "(Inputs!B1/12)"
    book = make_book(
        "gen",
        {
            "Inputs": {"B1": 6, "B2": 3, "A1": 1, "A2": 2},
            "Calc": {
                "B1": F(f'="t="&{x}', "t=0.5"),
                "B2": F(f"=LEN({x})", 3.0),
                "B3": F(f'=CONCATENATE("x",{x})', "x0.5"),
                "B4": F(f'=_xlfn.CONCAT("x",{x})', "x0.5"),
                "B5": F(f'=_xlfn.TEXTJOIN(",",TRUE,"a",{x})', "a,0.5"),
                "B6": F(f'=EXACT({x}&"","0.5")', True),
                "B7": F('="Year "&Inputs!B2', "Year 3"),
                "B8": F('=_xlfn.TEXTJOIN(",",TRUE,Inputs!A1:A2)', "1,2"),
                "B9": F(f"=UPPER({x})", "0.5"),
            },
        },
    )
    outputs = [f"Calc!B{r}" for r in range(1, 10)]
    res = _run(
        tmp_path, book, "id,m,n\nM7,7,3\nM3,3,4\n", {"m": "Inputs!B1", "n": "Inputs!B2"}, outputs
    )
    rows = _rows(res)
    for out in ("Calc!B1", "Calc!B2", "Calc!B3", "Calc!B4", "Calc!B5", "Calc!B6", "Calc!B9"):
        _refused(rows[("M7", out)], out, DIGITS)
        assert rows[("M3", out)]["status"] == "ok", (out, rows[("M3", out)])
    assert rows[("M7", "Calc!B7")]["value"] == "Year 3"
    assert rows[("M3", "Calc!B7")]["value"] == "Year 4"
    assert rows[("M3", "Calc!B7")]["status"] == "ok"
    # A range handed to TEXTJOIN cannot be read cell by cell from one probe.
    _refused(rows[("M3", "Calc!B8")], "Calc!B8", NOT_CHECKED)


def test_criteria_built_from_a_non_integer_are_refused(make_book: Book, tmp_path: Path) -> None:
    # Paid-to-date months in Data!A1:A4 = 1..4, amounts B1:B4 = 10..40. The criterion
    # "<="&(Inputs!B1/12) is built from the policy's term in months; stored 24 -> "<=2".
    # Policy T7 (7 months, 0.583...): Excel re-reads its 15-digit text, which the engine
    # does not write -- refused. Policy T36 -> "<=3": an integer, ok.
    crit = '"<="&(Inputs!B1/12)'
    book = make_book(
        "crit",
        {
            "Inputs": {"B1": 24},
            "Data": {"A1": 1, "A2": 2, "A3": 3, "A4": 4, "B1": 10, "B2": 20, "B3": 30, "B4": 40},
            "Calc": {
                "B1": F(f"=COUNTIF(Data!A1:A4,{crit})", 2.0),
                "B2": F(f"=SUMIF(Data!A1:A4,{crit},Data!B1:B4)", 30.0),
                "B3": F(f"=SUMIFS(Data!B1:B4,Data!A1:A4,{crit})", 30.0),
                "B4": F(f"=COUNTIFS(Data!A1:A4,{crit})", 2.0),
                "B5": F(f"=AVERAGEIF(Data!A1:A4,{crit},Data!B1:B4)", 15.0),
                "B6": F(f"=_xlfn.MAXIFS(Data!B1:B4,Data!A1:A4,{crit})", 20.0),
            },
        },
    )
    outputs = [f"Calc!B{r}" for r in range(1, 7)]
    res = _run(tmp_path, book, "id,t\nT7,7\nT36,36\n", {"t": "Inputs!B1"}, outputs)
    rows = _rows(res)
    for out in outputs:
        _refused(rows[("T7", out)], out, CRITERION)
        assert rows[("T36", out)]["status"] == "ok", (out, rows[("T36", out)])


@pytest.mark.parametrize(
    ("fmt", "rounding"),
    [
        ("0", (0, 1.0)),
        ("#,##0", (0, 1.0)),
        ("0.00", (2, 1.0)),
        ("#,##0.0##", (3, 1.0)),
        ("0.0%", (1, 100.0)),
        ("#,##0,", (0, 0.001)),
        ("0.0,,", (1, 1e-6)),
        ("0,.00", (2, 0.001)),
        ("$#,##0.00_);[Red]($#,##0.00)", (2, 1.0)),
        ('0.00 "years"', (2, 1.0)),
        ("[$€-407] #,##0.00", (2, 1.0)),
        ("0;-0;;@", (0, 1.0)),
        ("yyyy-mm-dd", None),
        ("[h]:mm", None),
        ('"fixed text"', None),
    ],
)
def test_number_format_rounding(fmt: str, rounding: tuple[int, float] | None) -> None:
    got = number_format_rounding(fmt)
    if rounding is None:
        assert got is None
    else:
        assert got is not None
        assert got[0] == rounding[0]
        assert got[1] == pytest.approx(rounding[1])


@pytest.mark.parametrize(
    "fmt",
    ["General", "@", "0.0E+0", "# ?/?", "[>100]0;0.0", "0;0.00", "hh:mm:ss.00", "0 abc", '0"'],
)
def test_number_format_rounding_refuses_what_it_cannot_read(fmt: str) -> None:
    with pytest.raises(ValueError):
        number_format_rounding(fmt)


@pytest.mark.parametrize(
    ("x", "text"),
    [
        (0.0, "0"),
        (1234.0, "1234"),
        (0.5, "0.5"),
        (-0.25, "-0.25"),
        (7 / 12, "0.583333333333333"),
        (0.1 + 0.2, "0.3"),
        (123456789012345.0, "123456789012345"),
        (0.0001234, "0.0001234"),
        (1e15, None),
        (1e-5, None),
    ],
)
def test_excel_general_text(x: float, text: str | None) -> None:
    # Excel's text for a number: 15 significant digits, positional in this range.
    assert excel_general_text(x) == text
