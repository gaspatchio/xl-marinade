"""Lookups written against defined names or LET variables must resolve like their range twins.

The lookup resolver reads values by address and could not dereference a name,
so `INDEX(MP_Value, MATCH(PointID, MP_ID, 0))` stopped at partial_resolved,
dropped its key drivers and emitted `UNRESOLVED:MP_Value`, while the same
lookup written with explicit ranges resolved to the exact row. Every lookup in
the lifelib example workbooks is written against names (issue #42, defect 1).
"""

import sqlite3
from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

import xl_marinade
from xl_marinade.core.new_arch.lookup_names import expand_names

NAMES = {
    "MP_ID": "Data!$A$2:$A$11",
    "MP_Value": "Data!$B$2:$B$11",
    "MP_Table": "Data!$A$2:$C$11",
    "PointID": "Calc!$B$1",
}


def _workbook(path: Path, *, named: bool, key: int = 3) -> Path:
    """Three lookups keyed by Calc!B1, written with names (named=True) or ranges."""

    def ref(name: str) -> str:
        return name if named else NAMES[name]

    wb = openpyxl.Workbook()
    calc = wb.active
    calc.title = "Calc"
    calc["A1"], calc["B1"] = "Selected ID", key
    calc["B4"] = f"=INDEX({ref('MP_Value')},MATCH({ref('PointID')},{ref('MP_ID')},0))"
    calc["B5"] = f"=VLOOKUP({ref('PointID')},{ref('MP_Table')},3,FALSE)"
    calc["B6"] = (
        f"=_xlfn.LET(_xlpm.k,{ref('PointID')},"
        f"INDEX({ref('MP_Value')},MATCH(_xlpm.k,{ref('MP_ID')},0)))"
    )
    data = wb.create_sheet("Data")
    for i in range(1, 11):
        data[f"A{i + 1}"], data[f"B{i + 1}"], data[f"C{i + 1}"] = i, i * 100, i / 10
    for name, target in NAMES.items():
        wb.defined_names[name] = DefinedName(name, attr_text=target)
    wb.save(path)
    return path


def _edges(db: Path) -> dict[str, list[tuple]]:
    conn = sqlite3.connect(db)
    try:
        return {
            "internal": sorted(
                conn.execute(
                    """
                    SELECT fs.sheet_name || '!' || f.a1, ts.sheet_name || '!' || t.a1
                    FROM cell_edges_internal e
                    JOIN cells f ON f.cell_id = e.from_cell_id
                    JOIN sheets fs ON fs.sheet_id = f.sheet_id
                    JOIN cells t ON t.cell_id = e.to_cell_id
                    JOIN sheets ts ON ts.sheet_id = t.sheet_id
                    """
                )
            ),
            "range": sorted(
                conn.execute(
                    """
                    SELECT fs.sheet_name || '!' || f.a1, ts.sheet_name || '!' || r.to_range_a1
                    FROM range_edges r
                    JOIN cells f ON f.cell_id = r.from_cell_id
                    JOIN sheets fs ON fs.sheet_id = f.sheet_id
                    JOIN sheets ts ON ts.sheet_id = r.to_sheet_id
                    """
                )
            ),
            "external": sorted(conn.execute("SELECT external_ref FROM cell_edges_external")),
            "metrics": sorted(conn.execute("SELECT * FROM resolution_metrics")),
        }
    finally:
        conn.close()


@pytest.fixture(scope="module")
def twins(tmp_path_factory: pytest.TempPathFactory) -> dict[bool, dict[str, list[tuple]]]:
    root = tmp_path_factory.mktemp("named_lookups")
    return {
        named: _edges(
            xl_marinade.extract(
                _workbook(root / f"{named}.xlsx", named=named), root / f"{named}.db"
            )
        )
        for named in (False, True)
    }


def test_named_lookups_produce_the_edges_of_their_range_twins(
    twins: dict[bool, dict[str, list[tuple]]],
) -> None:
    assert twins[True]["internal"] == twins[False]["internal"]
    assert twins[True]["range"] == twins[False]["range"]


def test_named_lookup_resolves_to_the_row_its_key_selects(
    twins: dict[bool, dict[str, list[tuple]]],
) -> None:
    internal = twins[True]["internal"]
    # PointID = 3 selects Data row 4, via INDEX/MATCH (B4) and the LET variable (B6).
    assert ("Calc!B4", "Data!B4") in internal
    assert ("Calc!B6", "Data!B4") in internal
    # ...and the key itself is a driver of every lookup.
    assert {("Calc!B4", "Calc!B1"), ("Calc!B5", "Calc!B1"), ("Calc!B6", "Calc!B1")} <= set(internal)


def test_no_name_is_left_as_an_unresolved_edge(
    twins: dict[bool, dict[str, list[tuple]]],
) -> None:
    unresolved = [ref for (ref,) in twins[True]["external"] if ref.startswith("UNRESOLVED:")]
    assert unresolved == []


def test_named_lookups_count_as_resolved(twins: dict[bool, dict[str, list[tuple]]]) -> None:
    assert twins[True]["metrics"] == twins[False]["metrics"]
    assert ("INDEX", "partial_resolved") not in {(f, s) for f, s, _ in twins[True]["metrics"]}


def test_a_key_edit_on_a_named_lookup_is_one_workbook_edit(tmp_path: Path) -> None:
    """The moved lookup edges are the key edit's consequence, not an edit (issue #42)."""
    base = xl_marinade.extract(_workbook(tmp_path / "a.xlsx", named=True), tmp_path / "a.db")
    edited = xl_marinade.extract(
        _workbook(tmp_path / "b.xlsx", named=True, key=5), tmp_path / "b.db"
    )
    result = xl_marinade.diff(base, edited)

    moved = [c for c in result["changes"] if c["type"].startswith("CELL_EDGE_")]
    assert {(c["type"], c["from"], c["to"]) for c in moved} >= {
        ("CELL_EDGE_REMOVED", "Calc!R4C2", "Data!R4C2"),
        ("CELL_EDGE_ADDED", "Calc!R4C2", "Data!R6C2"),
    }
    assert {c["layer"] for c in moved} == {"ir_inference"}
    edits = [
        (c["type"], c.get("cell"))
        for c in result["changes"]
        if c["layer"] == "workbook" and c["type"] != "IR_METADATA_CHANGED"
    ]
    assert edits == [("VALUE_CHANGED", "Calc!1,2")]


# --- expand_names, directly -----------------------------------------------------------

RANGES = {
    ("MP_Age", "Calc"): ["Model_Points!B2:B10001"],
    ("PointID", "Calc"): ["'Control'!C4"],
    ("Rate", "Other"): ["Other!A1"],  # sheet-scoped: only visible on Other
    ("Split", "Calc"): ["Data!A1:A5", "Data!C1:C5"],
}


def _resolve(name: str, sheet: str) -> list[str] | None:
    return RANGES.get((name, sheet))


def _ref(ref: str) -> dict:
    return {"type": "Ref", "ref": ref}


def test_expand_replaces_a_single_range_name() -> None:
    node = {"type": "Function", "name": "INDEX", "args": [_ref("MP_Age"), _ref("PointID")]}
    assert expand_names(node, "Calc", _resolve)["args"] == [
        _ref("Model_Points!B2:B10001"),
        _ref("'Control'!C4"),
    ]


@pytest.mark.parametrize("ref", ["B10", "Data!A1:A5", "Unknown", "Split"])
def test_expand_leaves_addresses_unknown_and_multi_area_names(ref: str) -> None:
    assert expand_names(_ref(ref), "Calc", _resolve) == _ref(ref)


def test_expand_uses_the_qualifying_sheet_for_scope() -> None:
    assert expand_names(_ref("Other!Rate"), "Calc", _resolve) == _ref("Other!A1")
    assert expand_names(_ref("Rate"), "Calc", _resolve) == _ref("Rate")


def test_expand_inlines_let_variables_in_order_and_shadows_names() -> None:
    # LET(PointID, B1, n, PointID + 1, INDEX(MP_Age, n)): PointID is the LET
    # variable here, not the defined name, and n is built from it.
    let = {
        "type": "Function",
        "name": "_xlfn.LET",
        "args": [
            _ref("PointID"),
            _ref("B1"),
            _ref("n"),
            {"type": "Binary", "op": "+", "left": _ref("PointID"), "right": {"type": "Const"}},
            {"type": "Function", "name": "INDEX", "args": [_ref("MP_Age"), _ref("n")]},
        ],
    }
    assert expand_names(let, "Calc", _resolve) == {
        "type": "Function",
        "name": "INDEX",
        "args": [
            _ref("Model_Points!B2:B10001"),
            {"type": "Binary", "op": "+", "left": _ref("B1"), "right": {"type": "Const"}},
        ],
    }
