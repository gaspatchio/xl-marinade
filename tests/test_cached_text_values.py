"""Cached formula text is stored exactly as Excel saved it (K1).

Issue: the extractor stripped leading and trailing whitespace from every cached
``<v>``, and treated an empty one as no value. A formula's text result is stored
as ``<c t="str"><v>...</v></c>`` (ECMA-376 Part 1, 18.18.11 ST_CellType ``str``:
"Cell containing a formula string"); the ``<v>`` text *is* the result, so
``=" "&"x"`` caches `` x`` and ``=IF(A1>1,"",A1)`` caches the empty string. Excel
keeps both: ``LEN`` of the first is 2, and ``=B4=""`` on the second is TRUE while
``ISBLANK(B4)`` is FALSE -- the cell holds a value. Trimming made `values` report
the wrong text and the calc trust run refuse cells it reproduced exactly; dropping
``""`` refused every projection that blanks periods beyond term with
``IF(t>Term,"",x)``.

Numeric, boolean and error caches are still read with surrounding whitespace
removed, and constants (shared strings, inline strings) were never trimmed.

The workbook is written by hand so each cached ``<v>`` is exactly what the test
says it is.
"""

from __future__ import annotations

import json
import sqlite3
import zipfile
from pathlib import Path

import openpyxl
import pytest

import xl_marinade
from xl_marinade.core.api import extract
from xl_marinade.core.new_arch.fast_parser import stream_worksheet_cells

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_SST_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"

SHARED_STRINGS = [" key ", "\tlead tab"]

# address -> the <c> element as it appears in the sheet XML (the r attribute is added).
CELLS: dict[str, str] = {
    "A1": "<c><v>2</v></c>",
    "B1": '<c t="str"><f>" "&amp;"x"</f><v> x</v></c>',
    "B2": '<c t="str"><f>"x"&amp;" "</f><v>x </v></c>',
    "B3": '<c t="str"><f>"  "</f><v>  </v></c>',
    "B4": '<c t="str"><f>IF(A1&gt;1,"",A1)</f><v></v></c>',
    "B5": '<c t="str"><f>IF(A1&gt;1,"",A1)</f><v/></c>',
    "B6": '<c t="str"><f>CHAR(9)&amp;"x"</f><v>\tx</v></c>',
    "B7": '<c t="str"><f>"x"&amp;CHAR(10)</f><v>x\n</v></c>',
    # CR as an XML character reference: this models the XML layer. Excel itself writes
    # CR in cell text as _x000D_, which the extractor does not decode (a separate gap).
    "B8": '<c t="str"><f>"x"&amp;CHAR(13)</f><v>x&#13;</v></c>',
    "B9": '<c t="str"><f>" a b "</f><v> a b </v></c>',
    "B10": '<c t="str"><f>"never calculated"</f></c>',
    # A filled-down IF(t>Term,"",x): Excel writes it as a shared formula.
    "B11": '<c t="str"><f t="shared" ref="B11:B12" si="0">IF(A1&gt;1,"",A1)</f><v></v></c>',
    "B12": '<c t="str"><f t="shared" si="0"/><v></v></c>',
    "C1": "<c><f>A1*2</f><v> 4 </v></c>",
    "C2": "<c><v> 7 </v></c>",
    "C3": '<c t="b"><f>A1&gt;1</f><v> 1 </v></c>',
    "C4": '<c t="e"><f>1/0</f><v> #DIV/0! </v></c>',
    "D1": '<c t="s"><v>0</v></c>',
    "D2": '<c t="s"><v>1</v></c>',
    "D3": '<c t="inlineStr"><is><t xml:space="preserve"> pad </t></is></c>',
    # An array (CSE) formula over E1:E6: only the anchor holds <f>; Excel writes each
    # other member as <c t="str"><v>...</v></c> holding that member's result.
    "E1": '<c t="str"><f t="array" ref="E1:E6">IF(A1:A6&gt;1," a","")</f><v> a</v></c>',
    "E2": '<c t="str"><v>b </v></c>',
    "E3": '<c t="str"><v>  </v></c>',
    "E4": '<c t="str"><v></v></c>',
    "E5": '<c t="str"><v/></c>',
    "E6": '<c t="str"><v>\tc</v></c>',
    # A dynamic-array (spill) anchor is written the same way, with cm="1".
    "F1": '<c t="str" cm="1"><f t="array" ref="F1:F3">IF(A1:A3&gt;1,"","x ")</f><v></v></c>',
    "F2": '<c t="str"><v/></c>',
    "F3": '<c t="str"><v>x </v></c>',
}


def _sheet_xml(cells: dict[str, str]) -> str:
    rows: dict[int, list[str]] = {}
    for ref, body in cells.items():
        row = int(ref.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
        rows.setdefault(row, []).append(body.replace("<c", f'<c r="{ref}"', 1))
    data = "".join(f'<row r="{r}">{"".join(cs)}</row>' for r, cs in sorted(rows.items()))
    return (
        f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{_MAIN}">'
        f"<sheetData>{data}</sheetData></worksheet>"
    )


def _sst_xml() -> str:
    items = "".join(f'<si><t xml:space="preserve">{s}</t></si>' for s in SHARED_STRINGS)
    n = len(SHARED_STRINGS)
    return (
        f'<?xml version="1.0" encoding="UTF-8"?>'
        f'<sst xmlns="{_MAIN}" count="{n}" uniqueCount="{n}">{items}</sst>'
    )


def _write_book(path: Path, cells: dict[str, str]) -> Path:
    """A one-sheet workbook ("S") whose sheet XML is exactly ``cells``."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "S"
    raw = path.with_suffix(".raw.xlsx")
    wb.save(raw)
    with zipfile.ZipFile(raw) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = _sheet_xml(cells).encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                rel = f'<Relationship Id="rIdSst" Type="{_REL}/sharedStrings" Target="sharedStrings.xml"/>'
                data = data.replace(b"</Relationships>", rel.encode() + b"</Relationships>")
            elif item.filename == "[Content_Types].xml":
                ct = f'<Override PartName="/xl/sharedStrings.xml" ContentType="{_SST_CT}"/>'
                data = data.replace(b"</Types>", ct.encode() + b"</Types>")
            zout.writestr(item, data)
        zout.writestr("xl/sharedStrings.xml", _sst_xml())
    raw.unlink()
    return path


@pytest.fixture(scope="module")
def workbook(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return _write_book(tmp_path_factory.mktemp("cached_text") / "t.xlsx", CELLS)


@pytest.fixture(scope="module")
def stored(workbook: Path, tmp_path_factory: pytest.TempPathFactory) -> dict[str, object]:
    """Address -> the decoded value blob ('<none>' when the cell has no stored value)."""
    db = extract(workbook, tmp_path_factory.mktemp("cached_text_db") / "t.db")
    conn = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
    try:
        rows = conn.execute("SELECT cell_address, value FROM agent_cells").fetchall()
    finally:
        conn.close()
    return {addr.split("!")[1]: "<none>" if v is None else json.loads(v) for addr, v in rows}


@pytest.mark.parametrize(
    ("cell", "expected"),
    [
        ("B1", " x"),
        ("B2", "x "),
        ("B3", "  "),
        ("B4", ""),
        ("B5", ""),
        ("B6", "\tx"),
        ("B7", "x\n"),
        ("B8", "x\r"),
        ("B9", " a b "),
        ("B11", ""),
        ("B12", ""),
        ("E1", " a"),
        ("F1", ""),
    ],
)
def test_formula_text_result_is_stored_exactly(
    stored: dict[str, object], cell: str, expected: str
) -> None:
    assert stored[cell] == expected


@pytest.mark.parametrize(
    ("cell", "expected"),
    [
        ("E2", "b "),
        ("E3", "  "),
        ("E4", ""),
        ("E5", ""),
        ("E6", "\tc"),
        ("F2", ""),
        ("F3", "x "),
    ],
)
def test_array_and_spill_member_text_is_stored_exactly(
    stored: dict[str, object], cell: str, expected: str
) -> None:
    # Members carry no <f>, but their <v> is still a formula's text result: "" and
    # whitespace-only results are values, so the member is a cell with that value.
    assert stored[cell] == expected


def test_formula_without_a_cached_result_has_no_stored_value(stored: dict[str, object]) -> None:
    # No <v> at all: never calculated, so there is nothing to store.
    assert stored["B10"] == "<none>"


@pytest.mark.parametrize(
    ("cell", "expected"),
    [("C1", "4"), ("C2", "7"), ("C3", "1"), ("C4", "#DIV/0!")],
)
def test_numeric_logical_and_error_caches_are_still_trimmed(
    stored: dict[str, object], cell: str, expected: str
) -> None:
    assert stored[cell] == expected


@pytest.mark.parametrize(
    ("cell", "expected"),
    [("D1", " key "), ("D2", "\tlead tab"), ("D3", " pad ")],
)
def test_text_constants_are_unchanged(stored: dict[str, object], cell: str, expected: str) -> None:
    assert stored[cell] == expected


def test_parser_yields_the_empty_string_for_an_empty_text_result(workbook: Path) -> None:
    cells = {
        (r, c): (v, t) for r, c, _f, v, t, _si, _m, _s in stream_worksheet_cells(str(workbook), "S")
    }
    assert cells[(4, 2)] == ("", "str")  # B4: <v></v>
    assert cells[(5, 2)] == ("", "str")  # B5: <v/>
    assert cells[(10, 2)] == (None, "str")  # B10: no <v>
    assert cells[(4, 5)] == ("", "str")  # E4: array member, <v></v>
    assert cells[(5, 5)] == ("", "str")  # E5: array member, <v/>


# -- characterisation: the whole text-cache surface of the workbook -------------
#
# Captured from the code before this fix and regenerated after it. The diff between
# the two is the t="str" cells above and what follows from them:
# - the values of B1-B9, B11, B12, E1, E2, E6, F1 and F3;
# - the array and spill members whose result is whitespace-only or "" (E3, E4, E5,
#   F2), which were empty cells before and now hold that value;
# - the constant bindings those members join (C2:E2 -> C2:F2, D3 and F3 -> D3:F3,
#   E6 -> E4:E6), and the placeholder rows for empty cells around them.
# No formula binding is relabelled or retyped, and no edge changes.
GOLDEN = Path(__file__).parent / "golden" / "cached_text_surface.json"


def _surface(db: Path) -> dict[str, list[list[object]]]:
    """Cells, bindings and edges keyed by address (binding ids hash the file's bytes)."""
    conn = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
    try:
        cells = conn.execute(
            "SELECT cell_address, data_type, formula, value FROM agent_cells ORDER BY 1"
        ).fetchall()
        bindings = conn.execute(
            "SELECT address, binding_type, label, is_orphan FROM agent_bindings ORDER BY 1"
        ).fetchall()
        edges = conn.execute(
            """
            SELECT fs.sheet_name || '!' || fc.a1, e.to_cell_id
            FROM cell_edges_internal e
            JOIN cells fc ON fc.cell_id = e.from_cell_id
            JOIN sheets fs ON fs.sheet_id = fc.sheet_id
            ORDER BY 1, 2
            """
        ).fetchall()
    finally:
        conn.close()
    return {
        "cells": [list(r) for r in cells],
        "bindings": [list(r) for r in bindings],
        "edges": [list(r) for r in edges],
    }


def test_text_cache_surface_matches_golden(
    workbook: Path, tmp_path_factory: pytest.TempPathFactory
) -> None:
    db = extract(workbook, tmp_path_factory.mktemp("cached_text_golden") / "t.db")
    assert _surface(db) == json.loads(GOLDEN.read_text(encoding="utf-8"))


# -- marinade diff: "" is a value like any other -------------------------------


def _diff_values(
    tmp: Path, a: dict[str, str], b: dict[str, str]
) -> dict[str, tuple[object, object]]:
    """VALUE_CHANGED cells of diff(a, b) -> (old value_json, new value_json)."""
    db_a = extract(_write_book(tmp / "a.xlsx", a), tmp / "a.db")
    db_b = extract(_write_book(tmp / "b.xlsx", b), tmp / "b.db")
    out = xl_marinade.diff(db_a, db_b)
    return {
        c["cell"]: (c["old_value"], c["new_value"])
        for c in out["changes"]
        if c["type"] == "VALUE_CHANGED"
    }


_BLANK_BEYOND_TERM = '<c t="str"><f>IF(A1&gt;1,"","x")</f><v></v></c>'


def test_diff_treats_a_saved_empty_text_result_like_any_other_value(tmp_path: Path) -> None:
    # Same formula, saved once calculated ("") and once never calculated (no <v>): a
    # recalculation gap, not an edit -- as for a cell whose result is " x" or 5.
    calculated = {"A1": "<c><v>2</v></c>", "B1": _BLANK_BEYOND_TERM}
    never = {"A1": "<c><v>2</v></c>", "B1": '<c t="str"><f>IF(A1&gt;1,"","x")</f></c>'}
    assert _diff_values(tmp_path, calculated, never) == {}


def test_diff_reports_an_empty_text_result_turning_into_text(tmp_path: Path) -> None:
    # A1 2 -> 0: the unchanged formula's result goes from "" to "x". Both are values.
    before = {"A1": "<c><v>2</v></c>", "B1": _BLANK_BEYOND_TERM}
    after = {"A1": "<c><v>0</v></c>", "B1": '<c t="str"><f>IF(A1&gt;1,"","x")</f><v>x</v></c>'}
    assert _diff_values(tmp_path, before, after) == {
        "S!1,1": ('"2"', '"0"'),
        "S!1,2": ('""', '"x"'),
    }
