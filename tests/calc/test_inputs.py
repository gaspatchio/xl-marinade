"""Input, config and output-file handling of ``calc()`` and ``values()``.

A value an input file holds is never dropped or converted by rules Excel does not use,
and a run never touches a file or sheet it does not own.
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
from xl_marinade.errors import CalcInputError  # noqa: E402

Book = Callable[..., Path]


def _write(tmp: Path, name: str, text: str) -> Path:
    path = tmp / name
    path.write_text(text, encoding="utf-8")
    return path


def _req(
    tmp: Path,
    book: Path,
    policies: str,
    imap: dict[str, object],
    outputs: list[str],
    **kw: object,
) -> CalcRequest:
    return CalcRequest(
        workbook=book,
        policies=_write(tmp, "policies.csv", policies),
        input_map=_write(tmp, "map.json", json.dumps(imap)),
        out_dir=tmp / "out",
        outputs=outputs,
        **kw,  # type: ignore[arg-type]
    )


def _results(res: CalcResult) -> dict[tuple[str, str], dict[str, str]]:
    with res.results_csv.open(encoding="utf-8", newline="") as fh:
        return {(r["policy_id"], r["output"]): r for r in csv.DictReader(fh)}


# -- two entries on one cell ------------------------------------------------------

_TWO = {
    "Inputs": {"A1": "age", "B1": 40, "A2": "term", "B2": 10},
    "Calc": {"B1": F("=Inputs!B1+1", 41), "B2": F("=Inputs!B2+1", 11)},
    "Results": {"A2": "k1", "B2": 51, "C2": 999},
    "Table": {"A2": "k1", "B2": 50},
}


def _two_book(make_book: Book) -> Path:
    return make_book("two", _TWO, names={"Age": "Inputs!$B$1", "Out": "Calc!$B$1"})


@pytest.mark.parametrize(
    "fields",
    [
        {"age": "Inputs!B1", "term": "Inputs!$B$1"},  # by address
        {"age": "Age", "term": "Inputs!B1"},  # by name and by address
        {"age": "age", "term": "inputs!b1"},  # names and sheets ignore case
    ],
)
def test_two_map_fields_on_one_cell_are_an_input_error(
    make_book: Book, tmp_path: Path, fields: dict[str, str]
) -> None:
    # A policy gives each field its own value; the cell can hold only one, so one
    # field would be dropped (V2: age 50 and term 20 gave 41, not 51).
    book = _two_book(make_book)
    with pytest.raises(CalcInputError, match=r"'age' and 'term'.*Inputs!B1"):
        calc(_req(tmp_path, book, "id,age,term\np,50,20\n", {"fields": fields}, ["Calc!B1"]))


def test_the_id_field_and_another_field_on_one_cell_are_an_input_error(
    make_book: Book, tmp_path: Path
) -> None:
    book = _two_book(make_book)
    imap = {"id_field": "id", "fields": {"id": "Inputs!B1", "age": "Inputs!B1"}}
    with pytest.raises(CalcInputError, match=r"'age' and 'id'.*Inputs!B1"):
        calc(_req(tmp_path, book, "id,age\n7,50\n", imap, ["Calc!B1"]))


def _check(outputs: str) -> str:
    return (
        '{"results": {"sheet": "Results", "id_column": "A", "first_row": 2, '
        f'"outputs": {outputs}}}, "inputs": {{"table": {{"sheet": "Table", "id_column": "A", '
        '"first_row": 2, "fields": {"age": "B"}}}}'
    )


@pytest.mark.parametrize(
    "outputs",
    [
        '{"Calc!B1": "B", "Out": "C"}',
        '{"Out": "C", "Calc!B1": "B"}',
        '{"Calc!B1": "B", "calc!$B$1": "C"}',
    ],
)
def test_two_check_outputs_on_one_cell_are_an_input_error(
    make_book: Book, tmp_path: Path, outputs: str
) -> None:
    # One cell compared against two results columns: one comparison was dropped, so
    # whether the 999 mismatch was found depended on key order (V2).
    book = _two_book(make_book)
    check = _write(tmp_path, "check.json", _check(outputs))
    imap = {"fields": {"age": "Inputs!B1"}}
    with pytest.raises(CalcInputError, match=r"check outputs .* both .*Calc!B1"):
        calc(_req(tmp_path, book, "id,age\np,50\n", imap, ["Calc!B1"], check_against=check))


def test_a_map_field_on_the_selector_cell_is_accepted(make_book: Book, tmp_path: Path) -> None:
    # The selector is set per input set, map fields per policy: never at the same time.
    sheets: dict[str, dict[str, object]] = {
        "Sel": {"A1": 1},
        "T": {"A1": 40, "A2": 70},
        "Calc": {"B1": F("=INDEX(T!A1:A2,Sel!A1)", 40)},
        "Results": {"A2": 1, "B2": 40, "A3": 2, "B3": 70},
    }
    book = make_book("selfield", sheets)
    check = {
        "results": {
            "sheet": "Results",
            "id_column": "A",
            "first_row": 2,
            "outputs": {"Calc!B1": "B"},
        },
        "inputs": {"selector": "Sel!A1"},
    }
    res = calc(
        _req(
            tmp_path,
            book,
            "id,sel\n5,2\n",
            {"fields": {"sel": "Sel!A1"}},
            ["Calc!B1"],
            check_against=_write(tmp_path, "check.json", json.dumps(check)),
        )
    )
    assert _results(res)[("5", "Calc!B1")]["value"] == "70.0"


# -- policy values by Excel's entry rules (rulings D3, D4) -------------------------
#
# Hand values: Excel 1900 date system, where 1900-01-01 is serial 1 and the calendar
# gains a 29 February 1900, so 1900-03-01 is serial 61 and 2024-06-30 is 45473
# (2024-06-30 - 1899-12-30 = 45473 days). Typing ``TRUE`` into a cell makes a boolean.

_EMPTY, _NUMBER = None, 1.0


@pytest.mark.parametrize(
    ("raw", "stored", "want"),
    [
        ("2024-06-30", _EMPTY, 45473.0),
        ("1900-03-01", _EMPTY, 61.0),
        ("1900-03-01", _NUMBER, 61.0),
        ("TRUE", _EMPTY, True),
        ("true", _EMPTY, True),
        ("False", _EMPTY, False),
        ("40", _EMPTY, 40.0),
        (" 0042 ", _EMPTY, 42.0),
        ("4e1", _EMPTY, 40.0),
        ("-1.5", _EMPTY, -1.5),
        (".5", _EMPTY, 0.5),
        ("1e3", _NUMBER, 1000.0),
        ("-0", _NUMBER, 0.0),
        ("abc", _EMPTY, "abc"),
        ("1_000", _EMPTY, "1_000"),  # Excel keeps it as text
        ("A5%", _EMPTY, "A5%"),
        ("5%", "code", "5%"),  # a text cell takes text as written
    ],
)
def test_coerce_input_accepts(raw: str, stored: float | str | None, want: object) -> None:
    from xl_marinade.calc.run import coerce_input

    got = coerce_input(raw, stored)
    assert got == want and type(got) is type(want)
    if isinstance(got, float):
        assert str(got) == str(want)  # -0 is 0, not -0.0


@pytest.mark.parametrize(
    ("raw", "stored", "named"),
    [
        # Locale-dependent spellings (D3): refused by name rather than guessed.
        ("5%", _EMPTY, "percent"),
        ("1,000", _EMPTY, "comma"),
        ("1,5", _EMPTY, "comma"),
        ("$100", _EMPTY, "currency"),
        ("€5", _EMPTY, "currency"),
        ("(100)", _EMPTY, "parentheses"),
        ("5%", _NUMBER, "percent"),
        ("1,000", _NUMBER, "comma"),
        ("٤٠", _EMPTY, "digits other than 0-9"),
        ("１２３", _NUMBER, "digits other than 0-9"),
        ("30/06/2024", _EMPTY, "YYYY-MM-DD"),
        ("30/06/2024", _NUMBER, "YYYY-MM-DD"),
        ("1/2", _EMPTY, "YYYY-MM-DD"),  # Excel (en-US) makes it 2 January
        ("1 1/2", _EMPTY, "fraction"),  # Excel makes it 1.5
        ("4 000", _EMPTY, "separators"),
        ("2024-06-30 12:00:00", _NUMBER, "time"),
        ("2024-06-30T12:00", _EMPTY, "time"),
        ("12:30", _NUMBER, "time"),
        ("#N/A", _EMPTY, "error value"),
        ("#n/a", _NUMBER, "error value"),
        # Outside Excel's number grammar or range.
        ("1_000", _NUMBER, "not a number"),
        ("1e-400", _EMPTY, "smallest"),
        ("1e-400", _NUMBER, "smallest"),
        ("1e-310", _NUMBER, "smallest"),
        ("1e400", _NUMBER, "largest"),
        ("1e308", _EMPTY, "largest"),
        ("inf", _NUMBER, "not a number"),
        ("nan", _NUMBER, "not a number"),
        # ISO dates in Excel's 1900 leap-year region (D4).
        ("1900-01-01", _EMPTY, "1900-03-01"),
        ("1900-01-01", _NUMBER, "1900-03-01"),
        ("1900-02-28", _NUMBER, "1900-03-01"),
        ("1900-02-29", _NUMBER, "1900-03-01"),
        ("1899-12-31", _EMPTY, "1900-03-01"),
        ("2023-02-30", _NUMBER, "not a valid date"),
    ],
)
def test_coerce_input_refuses_by_name(raw: str, stored: float | None, named: str) -> None:
    from xl_marinade.calc.run import coerce_input

    with pytest.raises(CalcInputError, match=named):
        coerce_input(raw, stored)


def _empty_target_book(make_book: Book) -> Path:
    # Inputs!B1 and B3 were saved empty; B2 and B4 are siblings so both arms of each
    # IF family are exercised by the trust run.
    return make_book(
        "empty",
        {
            "Inputs": {"A1": "start", "B2": 60000, "A3": "flag", "B4": True},
            "Calc": {
                "B1": F("=IF(Inputs!B1>50000,1,2)", 2),
                "B2": F("=IF(Inputs!B2>50000,1,2)", 1),
                "B3": F("=IF(Inputs!B3=TRUE,1,2)", 2),
                "B4": F("=IF(Inputs!B4=TRUE,1,2)", 1),
                "C1": F("=COUNT(Inputs!B1:B2)", 1),
                "C2": F("=Inputs!B3*100", 0),
            },
            "Results": {"A2": "k1", "B2": True, "C2": 1},
        },
    )


def test_an_empty_target_takes_dates_and_booleans_as_excel_does(
    make_book: Book, tmp_path: Path
) -> None:
    # Typed into the empty cells: 2024-06-30 is 45473, not over 50000, so B1 = 2, and
    # COUNT sees two numbers; TRUE is a boolean, so B3 = 1 and TRUE*100 = 100.
    book = _empty_target_book(make_book)
    imap = {"fields": {"start": "Inputs!B1", "flag": "Inputs!B3"}}
    outs = ["Calc!B1:B4", "Calc!C1:C2"]
    got = _results(calc(_req(tmp_path, book, "id,start,flag\np,2024-06-30,TRUE\n", imap, outs)))
    values = {out: row["value"] for (_pid, out), row in got.items()}
    assert values["Calc!B1"] == "2.0"
    assert values["Calc!C1"] == "2.0"
    assert values["Calc!B3"] == "1.0"
    assert values["Calc!C2"] == "100.0"


def test_an_empty_target_refuses_a_locale_spelling_naming_policy_and_field(
    make_book: Book, tmp_path: Path
) -> None:
    book = _empty_target_book(make_book)
    imap = {"fields": {"start": "Inputs!B1"}}
    with pytest.raises(CalcInputError, match=r"policy 'p' field 'start'.*'1,000'.*comma"):
        calc(_req(tmp_path, book, 'id,start\np,"1,000"\n', imap, ["Calc!B1"]))


def test_a_1900_sentinel_date_is_refused_naming_the_leap_year_region(
    make_book: Book, tmp_path: Path
) -> None:
    # Excel holds 1900-01-01 as serial 1; the old conversion gave 2 and reported ok.
    book = make_book("dob", {"Inputs": {"B1": 29221}, "Calc": {"B1": F("=Inputs!B1+0", 29221)}})
    imap = {"fields": {"dob": "Inputs!B1"}}
    with pytest.raises(CalcInputError, match=r"policy 'p' field 'dob'.*1900-03-01.*leap"):
        calc(_req(tmp_path, book, "id,dob\np,1900-01-01\n", imap, ["Calc!B1"]))


def test_check_table_values_keep_their_stored_type(make_book: Book, tmp_path: Path) -> None:
    # The table holds the boolean TRUE for k1; set into the empty Inputs!B3 it makes
    # Calc!B3 = 1, which is what the results sheet stored. As the text "TRUE" it was a
    # false mismatch (engine 2) that refused the output.
    book = _empty_target_book(make_book)
    check = {
        "results": {
            "sheet": "Results",
            "id_column": "A",
            "first_row": 2,
            "outputs": {"Calc!B3": "C"},
        },
        "inputs": {
            "table": {"sheet": "Results", "id_column": "A", "first_row": 2, "fields": {"flag": "B"}}
        },
    }
    res = calc(
        _req(
            tmp_path,
            book,
            "id,flag\np,TRUE\n",
            {"fields": {"flag": "Inputs!B3"}},
            ["Calc!B3"],
            check_against=_write(tmp_path, "check.json", json.dumps(check)),
        )
    )
    report = json.loads(res.trust_report.read_text(encoding="utf-8"))
    assert report["check_against"]["mismatches"] == []
    row = _results(res)[("p", "Calc!B3")]
    assert (row["value"], row["status"]) == ("1.0", "ok")


# -- values labels its row with the id field calc uses ------------------------------


def _rows(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8", newline="") as fh:
        return list(csv.DictReader(fh))


@pytest.mark.parametrize("imap", [{}, {"id_field": "id"}])
def test_values_and_calc_rows_share_keys_under_the_default_id_field(
    make_book: Book, tmp_path: Path, imap: dict[str, object]
) -> None:
    # The id field defaults to "id" (guide): the stored input set's id is the value
    # in its cell, 555, and calc labels the same input set 555 from policies.csv.
    from xl_marinade.calc.values import ValuesRequest, values

    book = make_book("vid", {"In": {"A1": 555, "A2": 3}, "Out": {"B1": F("=In!A2*2", 6)}})
    variables = _write(tmp_path, "vars.json", json.dumps({"npv": {"cell": "Out!B1"}}))
    full = {**imap, "fields": {"id": "In!A1", "n": "In!A2"}}
    mp = _write(tmp_path, "vmap.json", json.dumps(full))
    stored = _rows(values(ValuesRequest(book, variables, tmp_path / "v", input_map=mp)))
    res = calc(_req(tmp_path, book, "id,n\n555,3\n", full, [], variables=variables))
    assert res.rows_csv is not None
    computed = _rows(res.rows_csv)
    key = ("policy", "variable", "period", "value")
    assert [tuple(r[k] for k in key) for r in stored] == [("555", "npv", "", "6.0")]
    assert [tuple(r[k] for k in key) for r in computed] == [("555", "npv", "", "6.0")]


def test_values_writes_a_boolean_id_as_calc_reads_it(make_book: Book, tmp_path: Path) -> None:
    # A policies.csv id is text; a boolean id cell reads TRUE there, as Excel shows it.
    from xl_marinade.calc.values import ValuesRequest, values

    book = make_book("vbool", {"In": {"A1": True}, "Out": {"B1": F("=In!A1", True)}})
    variables = _write(tmp_path, "vars.json", json.dumps({"x": {"cell": "Out!B1"}}))
    mp = _write(tmp_path, "vmap.json", json.dumps({"fields": {"id": "In!A1"}}))
    stored = _rows(values(ValuesRequest(book, variables, tmp_path / "v", input_map=mp)))
    assert stored[0]["policy"] == "TRUE"


# -- never touch the user's files or sheets ----------------------------------------

_SIMPLE = {"Inputs": {"B1": 3}, "Calc": {"B1": F("=IF(Inputs!B1>1,Inputs!B1+1,0)", 4)}}
_SIMPLE_MAP = {"fields": {"x": "Inputs!B1"}}


@pytest.mark.parametrize(
    ("kind", "name"),
    [
        ("policies", "results.csv"),
        ("input_map", "run.json"),
        ("variables", "rows.csv"),
        ("check_against", "trust_report.json"),
        ("ir", "paths.csv"),
        ("workbook", "results.csv"),
    ],
)
def test_calc_refuses_an_input_it_would_clear_and_deletes_nothing(
    make_book: Book, ir_for: Callable[[Path], Path], tmp_path: Path, kind: str, name: str
) -> None:
    # The run clears its known output names from -o first; an input stored under one
    # of those names was deleted, then the run failed to read it (data loss).
    book = make_book("keep", _SIMPLE)
    out = tmp_path / "out"
    out.mkdir()
    req = _req(tmp_path, book, "id,x\np,7\n", _SIMPLE_MAP, ["Calc!B1"])
    sources = {
        "policies": req.policies,
        "input_map": req.input_map,
        "variables": _write(tmp_path, "vars.json", json.dumps({"v": {"cell": "Calc!B1"}})),
        "check_against": _write(tmp_path, "check.json", "{}"),
        "ir": ir_for(book),
        "workbook": book,
    }
    moved = out / name
    moved.write_bytes(sources[kind].read_bytes())
    before = moved.read_bytes()
    setattr(req, kind, moved)
    with pytest.raises(CalcInputError, match=rf"{name}.*-o|-o.*{name}"):
        calc(req)
    assert moved.read_bytes() == before


def test_calc_refuses_a_hard_link_to_an_input_under_an_output_name(
    make_book: Book, tmp_path: Path
) -> None:
    book = make_book("keep", _SIMPLE)
    out = tmp_path / "out"
    out.mkdir()
    req = _req(tmp_path, book, "id,x\np,7\n", _SIMPLE_MAP, ["Calc!B1"])
    (out / "results.csv").hardlink_to(req.policies)
    with pytest.raises(CalcInputError, match="results.csv"):
        calc(req)
    assert req.policies.read_text(encoding="utf-8") == "id,x\np,7\n"


def test_calc_still_clears_a_symlinked_output_and_keeps_its_target(
    make_book: Book, tmp_path: Path
) -> None:
    book = make_book("keep", _SIMPLE)
    out = tmp_path / "out"
    out.mkdir()
    target = _write(tmp_path, "elsewhere.csv", "mine")
    (out / "results.csv").symlink_to(target)
    calc(_req(tmp_path, book, "id,x\np,7\n", _SIMPLE_MAP, ["Calc!B1"]))
    assert target.read_text(encoding="utf-8") == "mine"
    assert not (out / "results.csv").is_symlink()


def test_values_refuses_an_input_it_would_clear(make_book: Book, tmp_path: Path) -> None:
    from xl_marinade.calc.values import ValuesRequest, values

    book = make_book("keep", _SIMPLE)
    out = tmp_path / "out"
    out.mkdir()
    imap = _write(out, "run.json", json.dumps(_SIMPLE_MAP))
    variables = _write(tmp_path, "vars.json", json.dumps({"v": {"cell": "Calc!B1"}}))
    with pytest.raises(CalcInputError, match="run.json"):
        values(ValuesRequest(book, variables, out, input_map=imap))
    assert imap.exists()


def test_an_out_dir_that_is_a_file_is_refused_before_the_run(
    make_book: Book, tmp_path: Path
) -> None:
    from xl_marinade.calc.values import ValuesRequest, values

    book = make_book("keep", _SIMPLE)
    afile = _write(tmp_path, "afile", "x")
    req = _req(tmp_path, book, "id,x\np,7\n", _SIMPLE_MAP, ["Calc!B1"])
    req.out_dir = afile
    with pytest.raises(CalcInputError, match="afile.*not a directory"):
        calc(req)
    variables = _write(tmp_path, "vars.json", json.dumps({"v": {"cell": "Calc!B1"}}))
    with pytest.raises(CalcInputError, match="afile.*not a directory"):
        values(ValuesRequest(book, variables, afile))
    req.out_dir = afile / "sub"
    with pytest.raises(CalcInputError, match="afile.*not a directory"):
        calc(req)
    assert afile.read_text(encoding="utf-8") == "x"


def test_a_workbook_sheet_named_like_the_probe_sheet_is_an_input_error(
    make_book: Book, tmp_path: Path
) -> None:
    book = make_book("probe", {**_SIMPLE, "_Marinade_Probes": {"A1": 1}})
    with pytest.raises(CalcInputError, match="_Marinade_Probes"):
        calc(_req(tmp_path, book, "id,x\np,7\n", _SIMPLE_MAP, ["Calc!B1"]))


def test_an_engine_panic_is_a_calc_error_naming_the_workbook(
    make_book: Book, tmp_path: Path
) -> None:
    # IronCalc 0.9.1 panics (index out of bounds) on XNPV(1,1); the panic is a
    # BaseException, which used to escape as a raw traceback.
    from xl_marinade.errors import CalcError

    book = make_book(
        "xnpv", {"Inputs": {"B1": 1}, "Calc": {"B1": F("=XNPV(Inputs!B1,1)", "#NUM!")}}
    )
    with pytest.raises(CalcError, match=r"engine.*xnpv\.xlsx"):
        calc(_req(tmp_path, book, "id,x\np,2\n", _SIMPLE_MAP, ["Calc!B1"]))


def test_cli_exits_1_without_a_traceback_on_an_engine_panic(
    make_book: Book, tmp_path: Path
) -> None:
    from typer.testing import CliRunner

    from xl_marinade.cli.main import app

    book = make_book(
        "xnpv", {"Inputs": {"B1": 1}, "Calc": {"B1": F("=XNPV(Inputs!B1,1)", "#NUM!")}}
    )
    req = _req(tmp_path, book, "id,x\np,2\n", _SIMPLE_MAP, ["Calc!B1"])
    args = ["calc", str(book), "--policies", str(req.policies), "--map", str(req.input_map)]
    args += ["--outputs", "Calc!B1", "-o", str(tmp_path / "out")]
    result = CliRunner().invoke(app, args)
    assert result.exit_code == 1, result.output
    assert "Traceback" not in result.output
    assert "xnpv.xlsx" in result.stderr


def test_an_engine_panic_for_one_policy_names_the_policy(make_book: Book, tmp_path: Path) -> None:
    from xl_marinade.errors import CalcError

    book = make_book(
        "xnpv2", {"Inputs": {"B1": 3}, "Calc": {"B1": F("=IF(Inputs!B1>5,XNPV(1,1),0)", 0)}}
    )
    req = _req(tmp_path, book, "id,x\nfine,2\nboom,7\n", _SIMPLE_MAP, ["Calc!B1"])
    with pytest.raises(CalcError, match=r"policy 'boom'.*crashed.*xnpv2\.xlsx"):
        calc(req)
    assert not (tmp_path / "out" / "results.csv").exists()  # no partial results
