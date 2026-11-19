"""Tests for variable expansion and ``values()``."""

from __future__ import annotations

import json
from collections.abc import Callable
from pathlib import Path

import pytest

from tests.calc.conftest import F
from xl_marinade.calc.graph import WorkbookGraph, open_graph
from xl_marinade.calc.models import CellRef, Reason, ReasonKind, Status, VariableSpec
from xl_marinade.calc.output import write_rows
from xl_marinade.calc.rows import Row, VariableCells, expand_variables, rows_for
from xl_marinade.calc.values import ValuesRequest, values
from xl_marinade.errors import CalcInputError

Book = Callable[..., Path]
IrFor = Callable[[Path], Path]


def write_json(tmp_path: Path, data: object, name: str = "vars.json") -> Path:
    p = tmp_path / name
    p.write_text(json.dumps(data), encoding="utf-8")
    return p


def _book(make_book: Book) -> Path:
    return make_book(
        "v",
        {
            "Proj": {"C11": F("=1", 1), "C12": F("=2", 2), "C13": F("=3", 3)},
            "Summary": {"B4": F("=6", 6)},
        },
    )


def test_values_rows(make_book: Book, tmp_path: Path) -> None:
    p = _book(make_book)
    vars_ = write_json(
        tmp_path,
        {"premium": {"range": "Proj!C11:C13", "period_start": 1}, "npv": {"cell": "Summary!B4"}},
    )
    out = values(ValuesRequest(p, vars_, tmp_path / "o" / "deep"))
    assert out.read_text().splitlines() == [
        "policy,variable,period,value,status,reasons",
        "stored,premium,1,1.0,stored,",
        "stored,premium,2,2.0,stored,",
        "stored,premium,3,3.0,stored,",
        "stored,npv,,6.0,stored,",
    ]
    run_path = tmp_path / "o" / "deep" / "run.json"
    run = json.loads(run_path.read_text())
    assert run["variables"]["premium"] == {
        "range": "Proj!C11:C13",
        "period_start": 1,
        "period_0_cell": "Proj!C10",
    }
    assert run["variables"]["npv"]["period_start"] is None
    assert run["variables"]["npv"]["period_0_cell"] is None
    assert run_path.read_text().endswith("}\n")
    # Deterministic rerun.
    first = (out.read_bytes(), run_path.read_bytes())
    values(ValuesRequest(p, vars_, tmp_path / "o" / "deep"))
    assert first == (out.read_bytes(), run_path.read_bytes())


def test_values_period_0_cell_row_range_and_top_edge(make_book: Book, tmp_path: Path) -> None:
    p = make_book("h", {"S": {"B1": 1, "C1": 2, "D1": 3}})
    vars_ = write_json(
        tmp_path,
        {
            "r": {"range": "S!B1:D1", "period_start": 1},
            "top": {"range": "S!B1:B1", "period_start": 1},
        },
    )
    out = values(ValuesRequest(p, vars_, tmp_path / "o"))
    run = json.loads((out.parent / "run.json").read_text())
    assert run["variables"]["r"]["period_0_cell"] == "S!A1"
    assert run["variables"]["top"]["period_0_cell"] is None


def test_values_2d_range_rejected(make_book: Book, tmp_path: Path) -> None:
    p = make_book("d", {"S": {"A1": 1, "B2": 2}})
    vars_ = write_json(tmp_path, {"x": {"range": "S!A1:B2", "period_start": 0}})
    with pytest.raises(CalcInputError, match="2-D range"):
        values(ValuesRequest(p, vars_, tmp_path / "o"))


def test_values_ir_reuse_and_hash_mismatch(make_book: Book, ir_for: IrFor, tmp_path: Path) -> None:
    p = _book(make_book)
    other = make_book("other", {"Proj": {"C11": 9}, "Summary": {"B4": 1}})
    vars_ = write_json(tmp_path, {"npv": {"cell": "Summary!B4"}})
    out = values(ValuesRequest(p, vars_, tmp_path / "o", ir=ir_for(p)))
    assert out.read_text().splitlines()[1] == "stored,npv,,6.0,stored,"
    with pytest.raises(CalcInputError, match="does not match"):
        values(ValuesRequest(p, vars_, tmp_path / "o2", ir=ir_for(other)))
    assert isinstance(open_graph(p, ir_for(p), tmp_path / "t"), WorkbookGraph)


def test_values_needs_no_ironcalc(blocked_ironcalc: None, make_book: Book, tmp_path: Path) -> None:
    p = _book(make_book)
    vars_ = write_json(tmp_path, {"npv": {"cell": "Summary!B4"}})
    assert values(ValuesRequest(p, vars_, tmp_path / "o")).exists()


def _id_run(make_book: Book, tmp_path: Path, cells: dict[str, object], target: str) -> Path:
    p = make_book("i", {"In": cells, "S": {"A1": 5}})
    vars_ = write_json(tmp_path, {"v": {"cell": "S!A1"}})
    imap = tmp_path / "map.json"
    imap.write_text(json.dumps({"id_field": "pid", "fields": {"pid": target}}))
    return values(ValuesRequest(p, vars_, tmp_path / "o", input_map=imap))


def test_values_stored_id_text_cell(make_book: Book, tmp_path: Path) -> None:
    out = _id_run(make_book, tmp_path, {"A1": "P7"}, "In!A1")
    assert out.read_text().splitlines()[1] == "P7,v,,5.0,stored,"


def test_values_stored_id_numeric_cell(make_book: Book, tmp_path: Path) -> None:
    out = _id_run(make_book, tmp_path, {"A1": 7}, "In!A1")
    assert out.read_text().splitlines()[1] == "7,v,,5.0,stored,"


def test_values_stored_id_multi_cell_rejected(make_book: Book, tmp_path: Path) -> None:
    with pytest.raises(CalcInputError, match="single cell"):
        _id_run(make_book, tmp_path, {"A1": "a", "A2": "b"}, "In!A1:A2")


def test_values_stored_id_empty_cell_falls_back(make_book: Book, tmp_path: Path) -> None:
    out = _id_run(make_book, tmp_path, {"B1": 1}, "In!A1")
    assert out.read_text().splitlines()[1] == "stored,v,,5.0,stored,"


def test_values_text_bool_error_empty(make_book: Book, tmp_path: Path) -> None:
    p = make_book(
        "t",
        {"S": {"A1": F('="hi"', "hi"), "A2": F("=1=1", True), "A3": F("=NA()", "#N/A"), "A4": 0}},
    )
    vars_ = write_json(tmp_path, {"c": {"range": "S!A1:A5", "period_start": 0}})
    out = values(ValuesRequest(p, vars_, tmp_path / "o"))
    lines = out.read_text().splitlines()
    assert lines[1:] == [
        "stored,c,0,hi,stored,",
        "stored,c,1,TRUE,stored,",
        "stored,c,2,#N/A,stored,",
        "stored,c,3,0.0,stored,",
        "stored,c,4,,stored,",
    ]


def test_expand_variables_directions(make_book: Book, ir_for: IrFor) -> None:
    p = make_book("e", {"S": {"A1": 1}})
    g = WorkbookGraph.from_ir(ir_for(p))
    cols = expand_variables([VariableSpec("c", "S!B2:B4", 1)], g)[0]
    assert cols.cells == [CellRef("S", r, 2) for r in (2, 3, 4)]
    rws = expand_variables([VariableSpec("r", "S!D2:B2", 1)], g)[0]
    assert rws.cells == [CellRef("S", 2, c) for c in (2, 3, 4)]


def test_rows_for_and_writer_sorts_reasons(tmp_path: Path) -> None:
    a, b = CellRef("S", 1, 1), CellRef("S", 2, 1)
    reasons = [
        Reason(ReasonKind.near_comparison, b, "x"),
        Reason(ReasonKind.engine_error, None, "boom"),
        Reason(ReasonKind.near_comparison, a, "y"),
    ]

    def value_of(c: CellRef) -> tuple[float | None, Status, list[Reason]]:
        return 1.5, Status.UNCERTAIN, reasons

    rows = rows_for("p1", [VariableCells("v", [a, b], 1), VariableCells("w", [a], None)], value_of)
    assert [(r.variable, r.period) for r in rows] == [("v", 1), ("v", 2), ("w", None)]
    assert isinstance(rows[0], Row)
    path = tmp_path / "rows.csv"
    write_rows(path, rows[2:])
    assert path.read_text().splitlines()[1] == (
        "p1,w,,1.5,uncertain,engine_error:boom; near_comparison@S!A1:y; near_comparison@S!A2:x"
    )
