"""Number-to-text sites the first pass missed or over-refused (review of PR-P part 1).

Hand values from Excel's documented behaviour:
- ``TEXT`` with a one-section format writes a negative number with its minus sign,
  also when it rounds to zero: ``TEXT(-0.01,"0%")`` is ``"-1%"``,
  ``TEXT(-1,"0")`` is ``"-1"``, ``TEXT(-0.01,"0")`` is ``"-0"``. A format with a
  negative section writes that section instead (``TEXT(-1,"0;(0)")`` is ``"(1)"``).
- ``VALUETOTEXT``, ``ARRAYTOTEXT``, the ``REGEX`` functions and ``TEXTSPLIT`` read a
  number as the text Excel writes for it, at most 15 significant digits:
  ``VALUETOTEXT(7/12)`` is ``"0.583333333333333"``, so ``REGEXTEST(7/12,"4$")`` is
  FALSE and ``REGEXREPLACE(7/12,"3","")`` is ``"0.58"``.
- ``TEXTJOIN``, ``CONCAT`` and ``CONCATENATE`` over a range join each cell's text:
  text as it is, integers as their digits.
- A number format's third section is used only for an exact zero.
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
from xl_marinade.calc.probes import (  # noqa: E402
    TEXT_ARGUMENTS,
    number_format_rounding,
    text_drops_sign,
)

Book = Callable[..., Path]

NEGATIVE = "number-to-text of a negative value that rounds to zero or one unit"
DIGITS = "number-to-text where Excel's 15-digit text differs"
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


def _ok(row: dict[str, str]) -> None:
    assert row["status"] == "ok", row


# -- I1: a negative value TEXT writes without its sign ---------------------------------


_SIGN_FORMATS = [("0", 0, 1.0), ("0.0", 1, 1.0), ("0.00", 2, 1.0), ("0%", 0, 100.0)]
_SIGN_FORMATS += [("0.0%", 1, 100.0), ("#,##0,", 0, 0.001), ("0.000", 3, 1.0)]


def test_the_sign_rule_covers_every_sign_ironcalc_drops() -> None:
    # Excel keeps the minus sign of every negative value under a one-section format.
    # Survey IronCalc 0.9.1 on a grid of 0.05 to 3 units of the format's last place.
    model = ironcalc.create("signs", "en", "UTC")
    cases: list[tuple[float, int, float, int]] = []
    row = 1
    for fmt, decimals, scale in _SIGN_FORMATS:
        for step in range(1, 61):
            x = -(step / 20) / 10**decimals / scale
            model.update_cell_with_formula(0, row, 1, f'=TEXT({x!r},"{fmt}")')
            cases.append((x, decimals, scale, row))
            row += 1
    model.evaluate()
    dropped = [
        (x, d, s) for x, d, s, r in cases if not str(model.get_cell_value(0, r, 1)).startswith("-")
    ]
    assert dropped, "IronCalc writes every sign now: drop the sign rule"
    for x, d, s in dropped:
        assert text_drops_sign(x, d, s), (x, d, s)
    # The rule asks no more than "rounds to zero or one unit".
    assert text_drops_sign(-1.4, 0, 1.0) and not text_drops_sign(-1.6, 0, 1.0)
    assert text_drops_sign(-0.014, 2, 1.0) and not text_drops_sign(-0.016, 2, 1.0)
    assert not text_drops_sign(0.01, 2, 1.0) and not text_drops_sign(-0.0, 0, 1.0)


def test_negative_value_written_without_its_sign_is_refused(
    make_book: Book, tmp_path: Path
) -> None:
    # Inputs!B1 is a change, stored 0.05. Policy N: -0.01. Excel:
    #   B1 TEXT(x,"0%")            "-1%"          B2 TEXT(x*100,"0")      "-1"
    #   B3 "Change: "&TEXT(x,"0.00") "Change: -0.01" B4 TEXT(x,"0")        "-0"
    #   B5 TEXT(x*100,"0;(0)")     "(1)"  -- a negative section: IronCalc writes it too
    #   B6 FIXED(x*100,0)          "-1"   -- FIXED keeps the sign in IronCalc
    #   B7 TEXT(x*1000,"0")        "-10"  B8 TEXT(x,"0.000") "-0.010": ten units
    # Policy M: -0.016 -> B2 "-2", B3 "Change: -0.02": two units, the sign is kept.
    book = make_book(
        "neg",
        {
            "Inputs": {"B1": 0.05},
            "Calc": {
                "B1": F('=TEXT(Inputs!B1,"0%")', "5%"),
                "B2": F('=TEXT(Inputs!B1*100,"0")', "5"),
                "B3": F('="Change: "&TEXT(Inputs!B1,"0.00")', "Change: 0.05"),
                "B4": F('=TEXT(Inputs!B1,"0")', "0"),
                "B5": F('=TEXT(Inputs!B1*100,"0;(0)")', "5"),
                "B6": F("=FIXED(Inputs!B1*100,0)", "5"),
                "B7": F('=TEXT(Inputs!B1*1000,"0")', "50"),
                "B8": F('=TEXT(Inputs!B1,"0.000")', "0.050"),
            },
        },
    )
    outputs = [f"Calc!B{r}" for r in range(1, 9)]
    res = _run(tmp_path, book, "id,x\nN,-0.01\nM,-0.016\n", {"x": "Inputs!B1"}, outputs)
    rows = _rows(res)
    for out in ("Calc!B1", "Calc!B2", "Calc!B3", "Calc!B4"):
        _refused(rows[("N", out)], out, NEGATIVE)
    for out in ("Calc!B5", "Calc!B6", "Calc!B7", "Calc!B8"):
        _ok(rows[("N", out)])
    assert rows[("N", "Calc!B5")]["value"] == "(1)"
    for out in ("Calc!B2", "Calc!B3"):
        _ok(rows[("M", out)])
    assert rows[("M", "Calc!B3")]["value"] == "Change: -0.02"


# -- I2: text functions that read a number as text -------------------------------------

# One call per listed function, with the number 7/12 at the listed argument and an
# outcome that shows how many digits the engine wrote. Excel writes 0.583333333333333.
_SAMPLES: dict[str, tuple[str, object]] = {
    "LEN": ("LEN(7/12)", 18.0),
    "LENB": ("LENB(7/12)", 18.0),
    "LEFT": ("LEFT(7/12,99)", "0.5833333333333334"),
    "LEFTB": ("LEFTB(7/12,99)", "0.5833333333333334"),
    "RIGHT": ("RIGHT(7/12,1)", "4"),
    "RIGHTB": ("RIGHTB(7/12,1)", "4"),
    "MID": ("MID(7/12,18,1)", "4"),
    "MIDB": ("MIDB(7/12,18,1)", "4"),
    "UPPER": ("UPPER(7/12)", "0.5833333333333334"),
    "LOWER": ("LOWER(7/12)", "0.5833333333333334"),
    "PROPER": ("PROPER(7/12)", "0.5833333333333334"),
    "TRIM": ("TRIM(7/12)", "0.5833333333333334"),
    "CLEAN": ("CLEAN(7/12)", "0.5833333333333334"),
    "CODE": ("CODE(7/12)", 48.0),
    "UNICODE": ("UNICODE(7/12)", 48.0),
    "REPT": ("REPT(7/12,1)", "0.5833333333333334"),
    "SUBSTITUTE": ('SUBSTITUTE(7/12,"3","")', "0.584"),
    "REPLACE": ('REPLACE(7/12,1,0,"")', "0.5833333333333334"),
    "EXACT": ('EXACT(7/12,"0.583333333333333")', False),
    "FIND": ('FIND("4",7/12)', 18.0),
    "FINDB": ('FINDB("4",7/12)', 18.0),
    "SEARCH": ('SEARCH("4",7/12)', 18.0),
    "SEARCHB": ('SEARCHB("4",7/12)', 18.0),
    "TEXTBEFORE": ('TEXTBEFORE(7/12,"4")', "0.583333333333333"),
    "TEXTAFTER": ('TEXTAFTER(7/12,"0.")', "5833333333333334"),
    "VALUETOTEXT": ("VALUETOTEXT(7/12)", "0.5833333333333334"),
    "ARRAYTOTEXT": ("ARRAYTOTEXT(7/12)", "0.5833333333333334"),
    "REGEXTEST": ('REGEXTEST(7/12,"4$")', True),
    "REGEXEXTRACT": ('REGEXEXTRACT(7/12,"[0-9]+$")', "5833333333333334"),
    "REGEXREPLACE": ('REGEXREPLACE(7/12,"3","")', "0.584"),
    "TEXTSPLIT": ('LEN(TEXTSPLIT(7/12,"x"))', 18.0),
}


def test_text_arguments_is_one_pinned_list() -> None:
    # Every listed function writes a number given to it with the engine's 17 digits,
    # where Excel writes 15: each is a site, and the list holds no other.
    assert set(TEXT_ARGUMENTS) == set(_SAMPLES)
    model = ironcalc.create("sites", "en", "UTC")
    names = sorted(_SAMPLES)
    for row, name in enumerate(names, 1):
        model.update_cell_with_formula(0, row, 1, f"={_SAMPLES[name][0]}")
    model.evaluate()
    got = {name: model.get_cell_value(0, row, 1) for row, name in enumerate(names, 1)}
    assert got == {name: _SAMPLES[name][1] for name in names}


def test_newer_text_functions_are_sites(make_book: Book, tmp_path: Path) -> None:
    # x = months, stored 6 (x/12 = 0.5 prints the same everywhere). Policy M7: 7/12.
    x = "(Inputs!B1/12)"
    book = make_book(
        "newer",
        {
            "Inputs": {"B1": 6},
            "Calc": {
                "B1": F(f"=_xlfn.VALUETOTEXT({x})", "0.5"),
                "B2": F(f"=_xlfn.ARRAYTOTEXT({x})", "0.5"),
                "B3": F(f'=_xlfn.REGEXTEST({x},"4$")', False),
                "B4": F(f'=_xlfn.REGEXEXTRACT({x},"[0-9]+$")', "5"),
                "B5": F(f'=_xlfn.REGEXREPLACE({x},"3","")', "0.5"),
                "B6": F(f'=LEN(_xlfn.TEXTSPLIT({x},"x"))', 3.0),
            },
        },
    )
    outputs = [f"Calc!B{r}" for r in range(1, 7)]
    res = _run(tmp_path, book, "id,m\nM7,7\nM3,3\n", {"m": "Inputs!B1"}, outputs)
    rows = _rows(res)
    for out in outputs:
        _refused(rows[("M7", out)], out, DIGITS)
        _ok(rows[("M3", out)])


# -- I3: ranges handed to a join -------------------------------------------------------


def _join_book(make_book: Book) -> Path:
    return make_book(
        "joins",
        {
            "Inputs": {"B1": 6, "B2": 1, "B3": 0.25},
            "Codes": {"A1": "AB", "A2": "CD", "A3": True, "B1": 1, "B2": 2, "B3": 0.25},
            "Bad": {"A1": "x", "A2": 1 / 3},
            "Mixed": {"A1": "m", "A2": F("=Inputs!B1/12", 0.5)},
            "Calc": {
                # Constants only: text, a boolean, a blank (A4), integers, a short decimal.
                "B1": F('=_xlfn.TEXTJOIN(",",TRUE,Codes!A1:A4)', "AB,CD,TRUE"),
                "B2": F("=_xlfn.CONCAT(Codes!A1:A3)", "ABCDTRUE"),
                # A legacy CONCATENATE intersects its range with the formula's row 3.
                "B3": F("=CONCATENATE(Codes!B1:B3)", "0.25"),
                "B4": F('=_xlfn.TEXTJOIN(",",TRUE,Codes!B1:B3)', "1,2,0.25"),
                # A constant Excel writes as 0.333333333333333; IronCalc 0.3333333333333333.
                "B5": F('=_xlfn.TEXTJOIN(",",TRUE,Bad!A1:A2)', "x,0.333333333333333"),
                # The same join on an arm only Inputs!B2 <= 0 takes.
                "B6": F('=IF(Inputs!B2>0,"none",_xlfn.TEXTJOIN(",",TRUE,Bad!A1:A2))', "none"),
                # A formula cell: 0.5 stored; 7/12 for policy M7.
                "B7": F('=_xlfn.TEXTJOIN(",",TRUE,Mixed!A1:A2)', "m,0.5"),
                # A mapped input: 0.25 stored; 0.30000000000000004 for policy P3.
                "B8": F('=_xlfn.TEXTJOIN(",",TRUE,Inputs!B2:B3)', "1,0.25"),
            },
        },
    )


def test_ranges_joined_as_text(make_book: Book, tmp_path: Path) -> None:
    # Policies: OK (stored path), M7 (months 7), P3 (input 0.1+0.2), NEG (switch <= 0).
    res = _run(
        tmp_path,
        _join_book(make_book),
        "id,m,s,v\nOK,6,1,0.25\nM7,7,1,0.25\nP3,6,1,0.30000000000000004\nNEG,6,-1,0.25\n",
        {"m": "Inputs!B1", "s": "Inputs!B2", "v": "Inputs!B3"},
        [f"Calc!B{r}" for r in range(1, 9)],
    )
    rows = _rows(res)
    for pid in ("OK", "M7", "P3", "NEG"):
        for out in ("Calc!B1", "Calc!B2", "Calc!B3", "Calc!B4"):
            _ok(rows[(pid, out)])
        _refused(rows[(pid, "Calc!B5")], "Calc!B5", DIGITS)
    assert rows[("OK", "Calc!B1")]["value"] == "AB,CD,TRUE"
    # B6's join is on an arm the policy does not take: nothing is converted.
    _ok(rows[("OK", "Calc!B6")])
    _refused(rows[("NEG", "Calc!B6")], "Calc!B6", DIGITS)
    # Formula cells and mapped inputs are read per policy.
    _ok(rows[("OK", "Calc!B7")])
    _refused(rows[("M7", "Calc!B7")], "Calc!B7", DIGITS)
    _ok(rows[("OK", "Calc!B8")])
    _refused(rows[("P3", "Calc!B8")], "Calc!B8", DIGITS)


def test_a_join_too_wide_to_probe_is_refused_per_policy(
    make_book: Book, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # With room for one probe per range, Mixed!A1:A2 (two cells to read) is refused
    # instead -- for the policies whose arm evaluates it, not for every policy.
    monkeypatch.setattr("xl_marinade.calc.probes.JOIN_RANGE_PROBES", 1)
    book = make_book(
        "wide",
        {
            "Inputs": {"B1": 6, "B2": 1},
            "Mixed": {"A1": F("=Inputs!B1/12", 0.5), "A2": F("=Inputs!B1/6", 1.0)},
            "Calc": {
                "B1": F('=IF(Inputs!B2>0,"none",_xlfn.TEXTJOIN(",",TRUE,Mixed!A1:A2))', "none"),
            },
        },
    )
    res = _run(
        tmp_path,
        book,
        "id,m,s\nOK,6,1\nNEG,6,-1\n",
        {"m": "Inputs!B1", "s": "Inputs!B2"},
        ["Calc!B1"],
    )
    rows = _rows(res)
    _ok(rows[("OK", "Calc!B1")])
    _refused(rows[("NEG", "Calc!B1")], "Calc!B1", NOT_CHECKED)


# -- M1: the zero section ----------------------------------------------------------------


@pytest.mark.parametrize(
    ("fmt", "rounding"),
    [
        ("0.00;-0.00;0", (2, 1.0)),
        ('_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', (2, 1.0)),
        ("#,##0;(#,##0);-", (0, 1.0)),
    ],
)
def test_the_zero_section_does_not_decide_the_rounding(
    fmt: str, rounding: tuple[int, float]
) -> None:
    # Excel shows the third section only for an exact zero, which sits on no tie.
    assert number_format_rounding(fmt) == rounding


def test_a_zero_section_with_other_decimals_stays_ok(make_book: Book, tmp_path: Path) -> None:
    # TEXT(1234,"0.00;-0.00;0") is "1234.00"; TEXT(-1234.4,...) is "-1234.40".
    book = make_book(
        "zero",
        {
            "Inputs": {"B1": 1234},
            "Calc": {"B1": F('=TEXT(Inputs!B1,"0.00;-0.00;0")', "1234.00")},
        },
    )
    res = _run(tmp_path, book, "id,x\nA,1234\nB,-1234.4\n", {"x": "Inputs!B1"}, ["Calc!B1"])
    rows = _rows(res)
    _ok(rows[("A", "Calc!B1")])
    assert rows[("B", "Calc!B1")]["value"] == "-1234.40"
    _ok(rows[("B", "Calc!B1")])
