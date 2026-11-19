"""Tests for the calc input-file loaders."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from xl_marinade.calc.config import (
    load_check_against,
    load_input_map,
    load_policies,
    load_variables,
    parse_outputs,
)
from xl_marinade.calc.models import InputMap, InputTable, Policy
from xl_marinade.errors import CalcInputError


def _write(path: Path, data: object) -> Path:
    path.write_text(json.dumps(data), encoding="utf-8")
    return path


def test_policies_bom_header(tmp_path: Path) -> None:
    p = tmp_path / "p.csv"
    p.write_bytes("﻿id,age\n007,40\n".encode())
    pols = load_policies(p, InputMap(fields={"age": "S!B1"}, id_field=None))
    assert pols == [Policy(id="007", values={"age": "40"})]


def test_custom_id_field(tmp_path: Path) -> None:
    p = tmp_path / "p.csv"
    p.write_text("policy_id,age\nA1,40\n", encoding="utf-8")
    pols = load_policies(p, InputMap(fields={"age": "S!B1"}, id_field="policy_id"))
    assert pols[0].id == "A1"


def test_duplicate_policy_id_is_error(tmp_path: Path) -> None:
    p = tmp_path / "p.csv"
    p.write_text("id,age\n007,40\n007,41\n", encoding="utf-8")
    with pytest.raises(CalcInputError, match="duplicate.*007"):
        load_policies(p, InputMap(fields={"age": "S!B1"}, id_field=None))


def test_blank_mapped_value_is_error(tmp_path: Path) -> None:
    p = tmp_path / "p.csv"
    p.write_text("id,age\n7,\n", encoding="utf-8")
    with pytest.raises(CalcInputError, match=r"7.*age"):
        load_policies(p, InputMap(fields={"age": "S!B1"}, id_field=None))


def test_missing_id_column_and_field(tmp_path: Path) -> None:
    p = tmp_path / "p.csv"
    p.write_text("x,age\n7,1\n", encoding="utf-8")
    with pytest.raises(CalcInputError, match="id"):
        load_policies(p, InputMap(fields={"age": "S!B1"}, id_field=None))
    p.write_text("id,y\n7,1\n", encoding="utf-8")
    with pytest.raises(CalcInputError, match="age"):
        load_policies(p, InputMap(fields={"age": "S!B1"}, id_field=None))


def test_quoted_sheet_targets_survive(tmp_path: Path) -> None:
    m = _write(tmp_path / "m.json", {"fields": {"age": "'Life Table '!B3"}})
    assert load_input_map(m).fields["age"] == "'Life Table '!B3"


def test_input_map_id_field_in_fields_is_accepted(tmp_path: Path) -> None:
    m = _write(tmp_path / "m.json", {"id_field": "a", "fields": {"a": "S!A1"}})
    loaded = load_input_map(m)
    assert loaded.id_field == "a"
    assert loaded.fields == {"a": "S!A1"}


def test_input_map_bad_shape(tmp_path: Path) -> None:
    bad = tmp_path / "m.json"
    bad.write_text("{nope", encoding="utf-8")
    with pytest.raises(CalcInputError):
        load_input_map(bad)
    with pytest.raises(CalcInputError):
        load_input_map(_write(bad, {"fields": {}}))
    with pytest.raises(CalcInputError):
        load_input_map(_write(bad, {"fields": {"a": 3}}))


def test_variables_order_and_kinds(tmp_path: Path) -> None:
    v = _write(
        tmp_path / "v.json",
        {
            "premium": {"range": "Proj!C11:C130", "period_start": 0},
            "npv": {"cell": "Summary!B4"},
        },
    )
    specs = load_variables(v)
    assert [(s.name, s.target, s.period_start) for s in specs] == [
        ("premium", "Proj!C11:C130", 0),
        ("npv", "Summary!B4", None),
    ]


def test_range_variable_needs_period_start(tmp_path: Path) -> None:
    v = _write(tmp_path / "v.json", {"premium": {"range": "Proj!C11:C20"}})
    with pytest.raises(CalcInputError, match="period_start"):
        load_variables(v)


def test_cell_variable_rejects_period_start(tmp_path: Path) -> None:
    v = _write(tmp_path / "v.json", {"n": {"cell": "S!B4", "period_start": 0}})
    with pytest.raises(CalcInputError, match="period_start"):
        load_variables(v)


def test_period_start_must_be_nonneg_int(tmp_path: Path) -> None:
    for bad in (-1, "0", 1.5, True):
        v = _write(tmp_path / "v.json", {"p": {"range": "S!A1:A2", "period_start": bad}})
        with pytest.raises(CalcInputError, match="period_start"):
            load_variables(v)


def test_variable_needs_exactly_one_of_range_cell(tmp_path: Path) -> None:
    v = _write(tmp_path / "v.json", {"p": {"range": "S!A1:A2", "cell": "S!A1"}})
    with pytest.raises(CalcInputError):
        load_variables(v)
    with pytest.raises(CalcInputError):
        load_variables(_write(tmp_path / "v2.json", {"p": {}}))


_RESULTS = {"sheet": "Res", "id_column": "A", "first_row": 2, "outputs": {"S!B13": "C"}}


def test_check_against_table(tmp_path: Path) -> None:
    c = _write(
        tmp_path / "c.json",
        {
            "results": _RESULTS,
            "inputs": {
                "table": {
                    "sheet": "In",
                    "id_column": "A",
                    "first_row": 2,
                    "fields": {"age": "B"},
                }
            },
        },
    )
    ca = load_check_against(c)
    assert ca.results_sheet == "Res"
    assert ca.outputs == {"S!B13": "C"}
    assert ca.selector is None
    assert ca.table == InputTable("In", "A", 2, {"age": "B"})


def test_check_against_selector(tmp_path: Path) -> None:
    c = _write(
        tmp_path / "c.json",
        {"results": _RESULTS, "inputs": {"selector": "Control!C4"}},
    )
    ca = load_check_against(c)
    assert ca.selector == "Control!C4"
    assert ca.table is None


def test_check_against_needs_exactly_one_input_source(tmp_path: Path) -> None:
    table = {"sheet": "In", "id_column": "A", "first_row": 2, "fields": {"a": "B"}}
    for inputs in ({}, {"selector": "C!A1", "table": table}):
        c = _write(tmp_path / "c.json", {"results": _RESULTS, "inputs": inputs})
        with pytest.raises(CalcInputError, match="exactly one"):
            load_check_against(c)


def test_parse_outputs_dedupes_in_order() -> None:
    assert parse_outputs(["A!B1", "A!B2", "A!B1"]) == ["A!B1", "A!B2"]


def test_policies_not_utf8_is_input_error(tmp_path: Path) -> None:
    p = tmp_path / "p.csv"
    p.write_bytes("id,name\n1,Jos\xe9\n".encode("cp1252"))
    with pytest.raises(CalcInputError, match="UTF-8"):
        load_policies(p, InputMap(fields={"name": "S!B1"}, id_field=None))


def test_policies_malformed_csv_is_input_error(tmp_path: Path) -> None:
    p = tmp_path / "p.csv"
    p.write_text("id,age\n1," + "9" * 200_000 + "\n", encoding="utf-8")  # over csv's field limit
    with pytest.raises(CalcInputError, match="CSV"):
        load_policies(p, InputMap(fields={"age": "S!B1"}, id_field=None))


def test_policies_duplicate_header_is_input_error(tmp_path: Path) -> None:
    p = tmp_path / "p.csv"
    p.write_text("id,age,age\n1,40,41\n", encoding="utf-8")
    with pytest.raises(CalcInputError, match="duplicate column.*age"):
        load_policies(p, InputMap(fields={"age": "S!B1"}, id_field=None))


def test_json_not_utf8_is_input_error(tmp_path: Path) -> None:
    p = tmp_path / "m.json"
    p.write_bytes('{"fields": {"n\xe9": "S!B1"}}'.encode("cp1252"))
    with pytest.raises(CalcInputError, match="UTF-8"):
        load_input_map(p)
