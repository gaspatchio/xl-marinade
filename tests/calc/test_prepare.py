import re
import zipfile
from datetime import date

import pytest

pytest.importorskip("ironcalc")

from tests.calc.conftest import F, add_external_link, patch_book  # noqa: E402
from xl_marinade.calc.engine import Engine  # noqa: E402
from xl_marinade.calc.graph import WorkbookGraph  # noqa: E402
from xl_marinade.calc.models import CellRef  # noqa: E402
from xl_marinade.calc.prepare import prepare_engine_copy  # noqa: E402
from xl_marinade.calc.trust import run_trust  # noqa: E402


def _sheet_xml(path, part="xl/worksheets/sheet1.xml"):
    with zipfile.ZipFile(path) as z:
        return z.read(part).decode("utf-8")


def _trust(copy, graph):
    eng = Engine.load(copy.path)
    eng.evaluate()
    return eng, run_trust(eng, graph, tolerance=1e-9, skip=set())


def test_strips_cached_values_of_formula_cells_only(make_book, tmp_path):
    book = make_book("b", {"S": {"A1": 2, "B1": F("=A1*3", 6), "C1": F('="x"&A1', "x2")}})
    copy = prepare_engine_copy(book, tmp_path / "e.xlsx", as_of=None)
    assert copy.stripped == 2
    assert copy.path == tmp_path / "e.xlsx"
    xml = _sheet_xml(copy.path)
    assert "<v>6</v>" not in xml and "x2" not in xml
    assert re.search(r'<c r="A1"[^>]*><v>2</v></c>', xml)
    assert copy.data_table_cells == [] and copy.held_externals == {}


def test_prepare_as_of(make_book, ir_for, tmp_path):
    book = make_book(
        "b", {"S": {"A1": F("=TODAY()", 46023), "B1": F('=IF(A1>0,"TODAY()","no")', "TODAY()")}}
    )
    copy = prepare_engine_copy(book, tmp_path / "e.xlsx", as_of=date(2026, 1, 1))
    xml = _sheet_xml(copy.path)
    assert "DATE(2026,1,1)" in xml
    assert "&quot;TODAY()&quot;" in xml or '"TODAY()"' in xml  # literal left alone
    eng, res = _trust(copy, WorkbookGraph.from_ir(ir_for(book)))
    assert eng.value(CellRef("S", 1, 1)) == 46023.0
    assert (res.checked, res.reproduced, res.failures) == (2, 2, [])


def test_prepare_now_rewritten_and_without_as_of_left(make_book, tmp_path):
    book = make_book("b", {"S": {"A1": F("=now()+1", 1.0)}})
    assert "DATE(2026,1,1)+1" in _sheet_xml(
        prepare_engine_copy(book, tmp_path / "e.xlsx", as_of=date(2026, 1, 1)).path
    )
    assert "now()+1" in _sheet_xml(prepare_engine_copy(book, tmp_path / "f.xlsx", as_of=None).path)


def _with_data_table(xml: str) -> str:
    cell = re.compile(r'<c r="B2"[^>]*>.*?</c>', re.S)
    xml = cell.sub(
        '<c r="B2"><f t="dataTable" ref="B2:B3" dt2D="0" dtr="0" r1="A1"/><v>10</v></c>', xml
    )
    return re.sub(r'<c r="B3"([^>]*)>.*?</c>', '<c r="B3"><v>20</v></c>', xml, flags=re.S)


def test_prepare_data_table_removed(make_book, ir_for, tmp_path):
    base = make_book(
        "base", {"S": {"A1": 2, "B1": F("=A1*10", 20), "A2": 1, "A3": 2, "B2": 0, "B3": 0}}
    )
    book = patch_book(
        base, tmp_path / "dt.xlsx", edit={"xl/worksheets/sheet1.xml": _with_data_table}
    )
    assert 't="dataTable"' in _sheet_xml(book)
    copy = prepare_engine_copy(book, tmp_path / "e.xlsx", as_of=None)
    xml = _sheet_xml(copy.path)
    assert "dataTable" not in xml and "<v>10</v>" not in xml and "<v>20</v>" not in xml
    assert copy.data_table_cells == [CellRef("S", 2, 2), CellRef("S", 3, 2)]
    eng = Engine.load(copy.path)  # loads: IronCalc refuses data-table formulas
    eng.evaluate()
    assert eng.value(CellRef("S", 1, 2)) == 20.0


def test_prepare_external_held(make_book, ir_for, tmp_path):
    base = make_book(
        "base",
        {
            "S": {
                "A1": F("=[1]Other!B2*2", 84),
                "A2": F("='[1]Other'!$C$1&\"!\"", "hi!"),
                "A3": F('="[1]Other!B2"', "[1]Other!B2"),
                "A4": F("=[1]Other!Z9", 7),
            }
        },
    )
    book = add_external_link(base, tmp_path / "ext.xlsx", "Other", {"B2": "42", "C1": "hi"})
    copy = prepare_engine_copy(book, tmp_path / "e.xlsx", as_of=None)
    assert copy.held_externals == {
        CellRef("S", 1, 1): ["[1]Other!B2"],
        CellRef("S", 2, 1): ["'[1]Other'!$C$1"],
    }
    eng = Engine.load(copy.path)
    eng.evaluate()
    assert eng.value(CellRef("S", 1, 1)) == 84.0
    assert eng.value(CellRef("S", 2, 1)) == "hi!"
    assert eng.value(CellRef("S", 3, 1)) == "[1]Other!B2"  # string literal untouched
    # No cached value: left as is, so the engine errors and the trust run reports it.
    # (IronCalc 0.9.1 gives #ERROR! here, #REF! when the book has no external link part.)
    assert "[1]Other!Z9" in _sheet_xml(copy.path)
    assert str(eng.value(CellRef("S", 4, 1))).startswith("#")


def test_prepare_external_negative_number_parenthesised(make_book, tmp_path):
    base = make_book("base", {"S": {"A1": F("=2^[1]Other!A1", 0.5)}})
    book = add_external_link(base, tmp_path / "ext.xlsx", "Other", {"A1": "-1"})
    copy = prepare_engine_copy(book, tmp_path / "e.xlsx", as_of=None)
    eng = Engine.load(copy.path)
    eng.evaluate()
    assert eng.value(CellRef("S", 1, 1)) == 0.5


def test_trust_on_external_book(make_book, ir_for, tmp_path):
    base = make_book("base", {"S": {"A1": F("=[1]Other!B2*2", 84), "A2": F("=[1]Other!Z9", 7)}})
    book = add_external_link(base, tmp_path / "ext.xlsx", "Other", {"B2": "42"})
    copy = prepare_engine_copy(book, tmp_path / "e.xlsx", as_of=None)
    _, res = _trust(copy, WorkbookGraph.from_ir(ir_for(book)))
    assert (res.checked, res.reproduced) == (2, 1)
    [failure] = res.failures
    assert failure.cell == CellRef("S", 2, 1) and failure.reason.kind.value == "engine_error"


def _shared(xml: str) -> str:
    xml = re.sub(
        r'<c r="A1"([^>]*)>.*?</c>',
        r'<c r="A1"\1><f t="shared" ref="A1:A2" si="0">[1]Other!$B$2+B1+[1]Other!C1</f>'
        r"<v>0</v></c>",
        xml,
        flags=re.S,
    )
    return re.sub(
        r'<c r="A2"([^>]*)>.*?</c>', r'<c r="A2"\1><f t="shared" si="0"/><v>0</v></c>', xml
    )


def test_prepare_shared_master_replaces_only_absolute_externals(make_book, tmp_path):
    base = make_book("base", {"S": {"A1": 0, "A2": 0, "B1": 1, "B2": 2}})
    shared = patch_book(base, tmp_path / "sh.xlsx", edit={"xl/worksheets/sheet1.xml": _shared})
    book = add_external_link(shared, tmp_path / "ext.xlsx", "Other", {"B2": "40", "C1": "5"})
    copy = prepare_engine_copy(book, tmp_path / "e.xlsx", as_of=None)
    xml = _sheet_xml(copy.path)
    # $B$2 is the same cell for every sharer; C1 shifts to C2 in A2, so it stays.
    assert "40+B1+[1]Other!C1" in xml
    assert copy.stripped == 2
    assert copy.held_externals == {
        CellRef("S", 1, 1): ["[1]Other!$B$2"],
        CellRef("S", 2, 1): ["[1]Other!$B$2"],
    }
