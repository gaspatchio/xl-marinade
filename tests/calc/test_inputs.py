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
