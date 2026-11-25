"""XLOOKUP and XMATCH are edge-probed like MATCH, and XLOOKUP's if_not_found arm is a
branch.

Hand values from Excel's documentation: ``XLOOKUP(v, keys, values, [if_not_found],
[match_mode], [search_mode])`` and ``XMATCH(v, keys, [match_mode], [search_mode])``;
match_mode 0 exact (default), -1 exact or next smaller, 1 exact or next larger, 2
wildcard; search_mode 1 first to last (default), -1 last to first. Not found is
``#N/A`` unless XLOOKUP's if_not_found is given; that arm is used only when nothing is
found.
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

Book = Callable[..., Path]

_KEYS = "Rates!$A$1:$A$3"
_VALS = "Rates!$B$1:$B$3"
_RATES = {"A1": 20, "A2": 21, "A3": 22, "B1": 0.01, "B2": 0.02, "B3": 0.03}
_RATES.update({"C1": "apple", "C2": "banana", "C3": "cherry"})


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


def _paths(res: CalcResult) -> list[dict[str, str]]:
    with res.paths_csv.open(encoding="utf-8", newline="") as fh:
        return list(csv.DictReader(fh))


# Stored age 21 lands on the interior row for every mode below, so the trust run
# exercises only "interior" (and "found").
_CELLS: dict[str, object] = {
    # exact; Excel: 20 -> 0.01, 22 -> 0.03, 99 -> #N/A
    "B1": F(f"=_xlfn.XLOOKUP(Inputs!B1,{_KEYS},{_VALS})", 0.02),
    # exact with if_not_found: 99 -> 0
    "B2": F(f"=_xlfn.XLOOKUP(Inputs!B1,{_KEYS},{_VALS},0)", 0.02),
    # XMATCH exact: 20 -> 1, 22 -> 3, 99 -> #N/A
    "B3": F(f"=_xlfn.XMATCH(Inputs!B1,{_KEYS})", 2.0),
    # next smaller: 99 -> 22 (last row, 0.03), 21.5 -> 21 (interior, 0.02), 19 -> "none"
    "B4": F(f'=_xlfn.XLOOKUP(Inputs!B1,{_KEYS},{_VALS},"none",-1)', 0.02),
    # next larger: 99 -> "none", 21.5 -> 22 (last row, 0.03), 19 -> 20 (first row)
    "B5": F(f'=_xlfn.XLOOKUP(Inputs!B1,{_KEYS},{_VALS},"none",1)', 0.02),
    # last to first: same positions as first to last for distinct keys
    "B6": F(f"=_xlfn.XMATCH(Inputs!B1,{_KEYS},0,-1)", 2.0),
    # wildcard on the names: "b*" -> banana (interior), "a*" -> apple (first), "z*" none
    "B7": F(f'=_xlfn.XLOOKUP(Inputs!B2,Rates!$C$1:$C$3,{_VALS},"none",2)', 0.02),
}
_OUTPUTS = [f"Calc!B{r}" for r in range(1, 8)]


def _book(make_book: Book) -> Path:
    return make_book(
        "xl",
        {"Inputs": {"B1": 21, "B2": "b*"}, "Rates": _RATES, "Calc": dict(_CELLS)},
    )


def _edge(
    rows: dict[tuple[str, str], dict[str, str]], pid: str, out: str, fn: str, edge: str
) -> None:
    row = rows[(pid, out)]
    assert row["status"] == "unverified_path", (pid, out, row)
    assert f"unexercised_lookup_edge@{out}:{fn} #0 landed {edge}, not exercised" in row["reasons"]


def test_xlookup_and_xmatch_edges_are_not_ok_unless_exercised(
    make_book: Book, tmp_path: Path
) -> None:
    res = _run(
        tmp_path,
        _book(make_book),
        "id,age,name\nA21,21,b*\nA20,20,a*\nA22,22,b*\nA99,99,z*\nA21h,21.5,b*\nA19,19,b*\n",
        {"age": "Inputs!B1", "name": "Inputs!B2"},
        _OUTPUTS,
    )
    rows = _rows(res)
    for out in _OUTPUTS:
        assert rows[("A21", out)]["status"] == "ok", (out, rows[("A21", out)])
    # first row
    assert rows[("A20", "Calc!B1")]["value"] == "0.01"
    for out, fn in (("Calc!B1", "XLOOKUP"), ("Calc!B2", "XLOOKUP"), ("Calc!B3", "XMATCH")):
        _edge(rows, "A20", out, fn, "first")
    _edge(rows, "A20", "Calc!B4", "XLOOKUP", "first")
    _edge(rows, "A20", "Calc!B5", "XLOOKUP", "first")
    _edge(rows, "A20", "Calc!B6", "XMATCH", "first")
    _edge(rows, "A20", "Calc!B7", "XLOOKUP", "first")
    # last row
    assert rows[("A22", "Calc!B3")]["value"] == "3.0"
    for out, fn in (("Calc!B1", "XLOOKUP"), ("Calc!B3", "XMATCH"), ("Calc!B6", "XMATCH")):
        _edge(rows, "A22", out, fn, "last")
    # not found
    assert rows[("A99", "Calc!B1")]["value"] == "#N/A"
    assert rows[("A99", "Calc!B2")]["value"] == "0.0"
    for out, fn in (("Calc!B1", "XLOOKUP"), ("Calc!B2", "XLOOKUP"), ("Calc!B3", "XMATCH")):
        _edge(rows, "A99", out, fn, "not_found")
    _edge(rows, "A99", "Calc!B5", "XLOOKUP", "not_found")
    _edge(rows, "A99", "Calc!B7", "XLOOKUP", "not_found")
    assert (
        "unexercised_branch@Calc!B2:XLOOKUP #0 took 'not_found'"
        in rows[("A99", "Calc!B2")]["reasons"]
    )
    # next smaller from 99 lands on the last row
    assert rows[("A99", "Calc!B4")]["value"] == "0.03"
    _edge(rows, "A99", "Calc!B4", "XLOOKUP", "last")
    # 21.5: next smaller is 21 (interior: exercised), next larger is 22 (last row)
    assert rows[("A21h", "Calc!B4")]["value"] == "0.02"
    assert rows[("A21h", "Calc!B4")]["status"] == "ok"
    assert rows[("A21h", "Calc!B5")]["value"] == "0.03"
    _edge(rows, "A21h", "Calc!B5", "XLOOKUP", "last")
    # 19: next smaller finds nothing, next larger finds the first row
    assert rows[("A19", "Calc!B4")]["value"] == "none"
    _edge(rows, "A19", "Calc!B4", "XLOOKUP", "not_found")
    _edge(rows, "A19", "Calc!B5", "XLOOKUP", "first")
    kinds = {(r["policy_id"], r["cell"], r["kind"], r["outcome"]) for r in _paths(res)}
    assert ("A99", "Calc!B2", "lookup", "not_found") in kinds
    assert ("A99", "Calc!B2", "branch", "not_found") in kinds
    assert ("A21", "Calc!B2", "branch", "found") in kinds


def test_an_edge_a_family_member_exercised_counts(make_book: Book, tmp_path: Path) -> None:
    # B1 and B2 are one family (B2 reads Inputs!B2 = 20, stored: the first row).
    book = make_book(
        "fam",
        {
            "Inputs": {"B1": 21, "B2": 20},
            "Rates": _RATES,
            "Calc": {
                "B1": F(f"=_xlfn.XLOOKUP(Inputs!B1,{_KEYS},{_VALS})", 0.02),
                "B2": F(f"=_xlfn.XLOOKUP(Inputs!B2,{_KEYS},{_VALS})", 0.01),
            },
        },
    )
    res = _run(
        tmp_path, book, "id,age\nA20,20\nA22,22\n", {"age": "Inputs!B1"}, ["Calc!B1", "Calc!B2"]
    )
    rows = _rows(res)
    assert rows[("A20", "Calc!B1")]["status"] == "ok"
    _edge(rows, "A22", "Calc!B1", "XLOOKUP", "last")


def test_binary_search_is_probed_on_sorted_keys_and_refused_on_unsorted(
    make_book: Book, tmp_path: Path
) -> None:
    # search_mode 2 (binary, ascending) finds what a linear search finds only on keys
    # sorted ascending: 20,21,22 are; 22,20,21 are not, and Excel's answer there is
    # undefined. Age 20 on the sorted keys is the first row, never exercised.
    book = make_book(
        "bin",
        {
            "Inputs": {"B1": 21},
            "Rates": {**_RATES, "D1": 22, "D2": 20, "D3": 21},
            "Calc": {
                "B1": F(f"=_xlfn.XLOOKUP(Inputs!B1,{_KEYS},{_VALS},0,0,2)", 0.02),
                "B2": F(f"=_xlfn.XLOOKUP(Inputs!B1,Rates!$D$1:$D$3,{_VALS},0,0,2)", 0.03),
            },
        },
    )
    res = _run(tmp_path, book, "id,age\nA20,20\n", {"age": "Inputs!B1"}, ["Calc!B1", "Calc!B2"])
    rows = _rows(res)
    _edge(rows, "A20", "Calc!B1", "XLOOKUP", "first")
    assert rows[("A20", "Calc!B2")]["status"] == "refused"
    assert "approximate lookup on unsorted column" in rows[("A20", "Calc!B2")]["reasons"]


def test_if_not_found_arm_counts_only_when_nothing_was_found(
    make_book: Book, tmp_path: Path
) -> None:
    # The if_not_found arm holds IF(Inputs!B3>0, 1, 2), B3 stored 5: in the trust run
    # the age (21) is found, so the arm is never used and its IF exercises nothing.
    # Age 99: not found, the arm gives IF(5>0) = 1 -- neither the not_found outcome nor
    # the IF's TRUE arm was checked against Excel.
    book = make_book(
        "nf",
        {
            "Inputs": {"B1": 21, "B3": 5},
            "Rates": _RATES,
            "Calc": {
                "B1": F(f"=_xlfn.XLOOKUP(Inputs!B1,{_KEYS},{_VALS},IF(Inputs!B3>0,1,2))", 0.02),
            },
        },
    )
    res = _run(tmp_path, book, "id,age\nA21,21\nA99,99\n", {"age": "Inputs!B1"}, ["Calc!B1"])
    rows = _rows(res)
    assert rows[("A21", "Calc!B1")]["status"] == "ok"
    row = rows[("A99", "Calc!B1")]
    assert row["value"] == "1.0"
    assert row["status"] == "unverified_path"
    assert "unexercised_branch@Calc!B1:XLOOKUP #0 took 'not_found'" in row["reasons"]
    assert "unexercised_branch@Calc!B1:IF #1 took 'TRUE'" in row["reasons"]
