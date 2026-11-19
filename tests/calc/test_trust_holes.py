"""Regression tests: ways a wrong or unchecked value could have come back ``ok``."""

from __future__ import annotations

import csv
import json
from collections.abc import Callable
from pathlib import Path

import pytest

pytest.importorskip("ironcalc")

from tests.calc.conftest import F, add_external_link  # noqa: E402
from xl_marinade.calc import CalcRequest, CalcResult, calc  # noqa: E402

Book = Callable[..., Path]


def _run(
    tmp: Path,
    book: Path,
    policies: str,
    fields: dict[str, str],
    outputs: list[str],
    *,
    check: dict[str, object] | None = None,
) -> CalcResult:
    (tmp / "p.csv").write_text(policies, encoding="utf-8")
    (tmp / "m.json").write_text(json.dumps({"id_field": "id", "fields": fields}), encoding="utf-8")
    check_path: Path | None = None
    if check is not None:
        check_path = tmp / "check.json"
        check_path.write_text(json.dumps(check), encoding="utf-8")
    return calc(
        CalcRequest(
            workbook=book,
            policies=tmp / "p.csv",
            input_map=tmp / "m.json",
            out_dir=tmp / "out",
            outputs=outputs,
            check_against=check_path,
        )
    )


def _rows(res: CalcResult) -> dict[tuple[str, str], dict[str, str]]:
    with res.results_csv.open(encoding="utf-8", newline="") as fh:
        return {(r["policy_id"], r["output"]): r for r in csv.DictReader(fh)}


def _paths(res: CalcResult) -> list[dict[str, str]]:
    with res.paths_csv.open(encoding="utf-8", newline="") as fh:
        return list(csv.DictReader(fh))


# -- C1: failing and refused cells vouch for nothing ---------------------------


def test_failing_family_member_does_not_exercise_its_branch(
    make_book: Book, tmp_path: Path
) -> None:
    # B1 takes TRUE in the trust run but its stored value is wrong; C1 (same family)
    # takes FALSE there and TRUE for the new policy.
    book = make_book(
        "fail",
        {
            "Inputs": {"B1": 10, "C1": 1},
            "Calc": {
                "B1": F("=IF(Inputs!B1>5,Inputs!B1*2,Inputs!B1)", 999),
                "C1": F("=IF(Inputs!C1>5,Inputs!C1*2,Inputs!C1)", 1),
            },
        },
    )
    # Both outputs requested, so B1 is probed too (probes go upstream of outputs only).
    res = _run(tmp_path, book, "id,c\n1,7\n", {"c": "Inputs!C1"}, ["Calc!B1", "Calc!C1"])
    row = _rows(res)[("1", "Calc!C1")]
    assert row["value"] == "14.0"
    assert row["status"] == "unverified_path"
    assert "unexercised_branch@Calc!C1" in row["reasons"]


def test_refused_pattern_cells_get_no_probes(make_book: Book, tmp_path: Path) -> None:
    book = make_book(
        "refprobe",
        {
            "Inputs": {"B1": 10},
            "Calc": {"B1": F("=IF(Inputs!B1>5,Inputs!B1*2,Inputs!B1)+RAND()*0", 20)},
        },
    )
    res = _run(tmp_path, book, "id,c\n1,7\n", {"c": "Inputs!B1"}, ["Calc!B1"])
    assert _rows(res)[("1", "Calc!B1")]["status"] == "refused"
    assert _paths(res) == []


# -- C2: a defined name over an external range ---------------------------------


def test_name_over_external_range_is_refused(make_book: Book, tmp_path: Path) -> None:
    base = make_book(
        "extname",
        {
            "Inputs": {"B1": 99},
            "Calc": {"B1": F("=IFERROR(VLOOKUP(Inputs!B1,Ext,2,FALSE),0)", 0)},
        },
        names={"Ext": "[1]Src!$A$1:$B$3"},
    )
    cache = {"A1": "1", "B1": "100", "A2": "2", "B2": "200", "A3": "3", "B3": "300"}
    book = add_external_link(base, tmp_path / "ext.xlsx", "Src", cache)
    res = _run(tmp_path, book, "id,x\n1,2\n", {"x": "Inputs!B1"}, ["Calc!B1"])
    row = _rows(res)[("1", "Calc!B1")]
    assert row["status"] == "refused"
    assert "refused_pattern@Calc!B1:external reference not held" in row["reasons"]


# -- I1: probes inside an arm that was not taken --------------------------------


def test_lookup_in_untaken_arm_is_not_evidence(make_book: Book, tmp_path: Path) -> None:
    # Trust run: B1's flag is 0, so its lookup (which would land "last") never ran;
    # C1 takes the lookup arm and lands "first".
    book = make_book(
        "arm",
        {
            "Inputs": {"B1": 0, "B2": 35, "C1": 1, "C2": 15},
            "Tbl": {"A1": 10, "B1": 1, "A2": 20, "B2": 2, "A3": 30, "B3": 3},
            "Calc": {
                "B1": F("=IF(Inputs!B1>0.5,VLOOKUP(Inputs!B2,Tbl!$A$1:$B$3,2),0)", 0),
                "C1": F("=IF(Inputs!C1>0.5,VLOOKUP(Inputs!C2,Tbl!$A$1:$B$3,2),0)", 1),
            },
        },
    )
    res = _run(
        tmp_path,
        book,
        "id,f,x\n1,1,35\n2,0,35\n",
        {"f": "Inputs!B1", "x": "Inputs!B2"},
        ["Calc!B1"],
    )
    rows = _rows(res)
    taken = rows[("1", "Calc!B1")]
    assert taken["value"] == "3.0"
    assert taken["status"] == "unverified_path"
    assert "unexercised_lookup_edge@Calc!B1" in taken["reasons"]
    # Policy 2 does not take the lookup arm: its lookup is not a path it took.
    assert rows[("2", "Calc!B1")]["status"] == "ok"
    kinds = {(p["policy_id"], p["kind"]) for p in _paths(res)}
    assert ("1", "lookup") in kinds
    assert ("2", "lookup") not in kinds


# -- I2: exact equality is not a near comparison --------------------------------


def test_exact_equality_comparison_is_ok(make_book: Book, tmp_path: Path) -> None:
    book = make_book(
        "teq",
        {
            "Inputs": {"B1": 100},
            "Proj": {
                "A1": 0,
                "B1": F("=IF(Proj!A1=0,Inputs!B1,0)", 100),
                "B2": F("=Proj!B1*1.1", 110.00000000000001),
            },
        },
    )
    res = _run(tmp_path, book, "id,p\n1,200\n", {"p": "Inputs!B1"}, ["Proj!B1:B2"])
    rows = _rows(res)
    assert rows[("1", "Proj!B1")]["status"] == "ok"
    assert rows[("1", "Proj!B2")]["status"] == "ok"


# -- I3: engine-only error codes ---------------------------------------------------


def test_engine_only_error_code_is_refused(make_book: Book, tmp_path: Path) -> None:
    book = make_book(
        "nimpl",
        {
            "Inputs": {"B1": "row"},
            "Data": {"C3": 5},
            "Calc": {
                "B1": F("=CELL(Inputs!B1,Data!C3)", 3),
                "B2": F('=IFERROR(CELL(Inputs!B1,Data!C3),"n/a")', 3),
            },
        },
    )
    res = _run(tmp_path, book, "id,k\n1,address\n2,col\n", {"k": "Inputs!B1"}, ["Calc!B1:B2"])
    rows = _rows(res)
    out = rows[("1", "Calc!B1")]
    assert out["value"] == "#N/IMPL!"
    assert out["status"] == "refused"
    assert "engine_error@Calc!B1" in out["reasons"]
    hidden = rows[("1", "Calc!B2")]
    assert hidden["status"] == "refused"
    assert "engine_error@Calc!B2" in hidden["reasons"]
    assert rows[("2", "Calc!B1")]["status"] == "ok"
    assert rows[("2", "Calc!B2")]["status"] == "ok"


def test_excel_error_code_stays_a_value(make_book: Book, tmp_path: Path) -> None:
    book = make_book(
        "div0",
        {"Inputs": {"B1": 2}, "Calc": {"B1": F("=IFERROR(1/Inputs!B1,-1)", 0.5)}},
    )
    res = _run(tmp_path, book, "id,k\n1,4\n", {"k": "Inputs!B1"}, ["Calc!B1"])
    assert _rows(res)[("1", "Calc!B1")]["status"] == "ok"


def test_text_starting_with_hash_is_not_an_error(make_book: Book, tmp_path: Path) -> None:
    book = make_book(
        "hash",
        {"Inputs": {"B1": "#tag"}, "Calc": {"B1": F('=Inputs!B1&"!"', "#tag!")}},
    )
    res = _run(tmp_path, book, "id,k\n1,#N/IMPL\n", {"k": "Inputs!B1"}, ["Calc!B1"])
    row = _rows(res)[("1", "Calc!B1")]
    assert (row["value"], row["status"]) == ("#N/IMPL!", "ok")


# -- I4: a check-against set vouches only for what it compares ------------------


def test_check_set_credits_only_probes_upstream_of_compared_outputs(
    make_book: Book, tmp_path: Path
) -> None:
    book = make_book(
        "chkup",
        {
            "Inputs": {"B1": 40},
            "Calc": {
                "B1": F("=IF(Inputs!B1>60,Inputs!B1*2,Inputs!B1)", 40),
                "D1": F("=IF(Inputs!B1>65,1,2)", 2),
            },
            "Results": {"A2": 1, "B2": 140},
            "Table": {"A2": 1, "B2": 70},
        },
    )
    check = {
        "results": {
            "sheet": "Results",
            "id_column": "A",
            "first_row": 2,
            "outputs": {"Calc!B1": "B"},
        },
        "inputs": {
            "table": {"sheet": "Table", "id_column": "A", "first_row": 2, "fields": {"age": "B"}}
        },
    }
    res = _run(
        tmp_path,
        book,
        "id,age\n9,75\n",
        {"age": "Inputs!B1"},
        ["Calc!B1", "Calc!D1"],
        check=check,
    )
    rows = _rows(res)
    assert rows[("9", "Calc!B1")]["status"] == "ok"
    assert rows[("9", "Calc!D1")]["status"] == "unverified_path"


# -- rulings: approximate lookups whose answer can change per policy -----------


def test_approximate_lookup_on_input_dependent_keys_is_refused(
    make_book: Book, tmp_path: Path
) -> None:
    book = make_book(
        "inkeys",
        {
            "Inputs": {"B1": 10},
            "Tbl": {"A1": F("=Inputs!B1", 10), "B1": 1, "A2": 20, "B2": 2, "A3": 30, "B3": 3},
            "Calc": {"B1": F("=VLOOKUP(25,Tbl!$A$1:$B$3,2)", 2)},
        },
    )
    res = _run(tmp_path, book, "id,k\n1,40\n", {"k": "Inputs!B1"}, ["Calc!B1"])
    row = _rows(res)[("1", "Calc!B1")]
    assert row["status"] == "refused"
    assert "approximate lookup on input-dependent keys" in row["reasons"]
    report = json.loads(res.trust_report.read_text(encoding="utf-8"))
    assert {"cell": "Calc!B1", "detail": "approximate lookup on input-dependent keys"} in report[
        "refused_patterns"
    ]


def test_approximate_lookup_on_constant_keys_is_not_refused(
    make_book: Book, tmp_path: Path
) -> None:
    book = make_book(
        "constkeys",
        {
            "Inputs": {"B1": 25},
            "Tbl": {"A1": 10, "B1": 1, "A2": 20, "B2": 2, "A3": 30, "B3": 3},
            "Calc": {"B1": F("=VLOOKUP(Inputs!B1,Tbl!$A$1:$B$3,2)", 2)},
        },
    )
    res = _run(tmp_path, book, "id,k\n1,26\n", {"k": "Inputs!B1"}, ["Calc!B1"])
    assert _rows(res)[("1", "Calc!B1")]["status"] == "ok"


def test_lookup_value_type_differs_from_keys_is_refused(make_book: Book, tmp_path: Path) -> None:
    book = make_book(
        "ltype",
        {
            "Inputs": {"B1": 0},
            "Tbl": {"A1": 10, "B1": 1, "A2": 20, "B2": 2, "A3": 30, "B3": 3},
            "Calc": {
                "A1": F('=IF(Inputs!B1=1,"x",25)', 25),
                "B1": F("=VLOOKUP(Calc!A1,Tbl!$A$1:$B$3,2)", 2),
            },
        },
    )
    res = _run(tmp_path, book, "id,k\n1,1\n2,0\n", {"k": "Inputs!B1"}, ["Calc!B1"])
    rows = _rows(res)
    bad = rows[("1", "Calc!B1")]
    assert bad["status"] == "refused"
    assert "refused_pattern@Calc!B1:lookup value type differs from keys" in bad["reasons"]
    assert rows[("2", "Calc!B1")]["status"] == "ok"
