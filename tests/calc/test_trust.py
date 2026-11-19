import pytest

pytest.importorskip("ironcalc")

from tests.calc.conftest import F, patch_book  # noqa: E402
from xl_marinade.calc.engine import Engine  # noqa: E402
from xl_marinade.calc.graph import WorkbookGraph  # noqa: E402
from xl_marinade.calc.models import CellRef, ReasonKind  # noqa: E402
from xl_marinade.calc.prepare import prepare_engine_copy  # noqa: E402
from xl_marinade.calc.trust import TrustResult, run_trust, same  # noqa: E402


def test_same():
    assert same(1.0, 1.0 + 1e-12, 1e-9) and not same(1.0, 1.001, 1e-9)
    assert same("#N/A", "#N/A", 1e-9) and not same(True, 1.0, 1e-9) and same("", None, 1e-9)
    assert same(1e12, 1e12 + 1.0, 1e-9) and not same(0.0, 1e-8, 1e-9)
    assert not same(1.0, True, 1e-9) and same(False, False, 1e-9) and not same(True, False, 1e-9)
    assert not same("#N/A", "#VALUE!", 1e-9) and not same("a", "A", 1e-9)
    assert not same(0.0, None, 1e-9) and not same(None, 0.0, 1e-9)
    assert not same(float("nan"), float("nan"), 1e-9) and not same(1.0, "1", 1e-9)


def _run(book, ir_for, tmp_path, skip=frozenset()):
    copy = prepare_engine_copy(book, tmp_path / f"{book.stem}.engine.xlsx", as_of=None)
    eng = Engine.load(copy.path)
    eng.evaluate()
    return run_trust(eng, WorkbookGraph.from_ir(ir_for(book)), tolerance=1e-9, skip=set(skip))


def _only(res, cell):
    [failure] = res.failures
    assert failure.cell == cell and failure.reason.root == cell
    return failure.reason.kind


def test_trust_reproduces_simple_book(make_book, ir_for, tmp_path):
    book = make_book(
        "b",
        {
            "S": {
                "A1": 2,
                "B1": F("=A1*3", 6),
                "C1": F('=IF(A1>1,"hi","lo")', "hi"),
                "D1": F("=NA()", "#N/A"),
                "E1": F("=A1>1", True),
                "F1": F("=SUM(A1:B1)/7", 8 / 7),
            }
        },
    )
    res = _run(book, ir_for, tmp_path)
    assert (res.checked, res.reproduced, res.failures, res.share) == (5, 5, [], 1.0)


def test_trust_flags_silently_wrong(make_book, ir_for, tmp_path):
    cells: dict[str, object] = {"A1": " "}
    for i in range(1, 6):
        cells[f"B{i}"] = i
        cells[f"C{i}"] = i * 10
    cells["D1"] = F("=VLOOKUP(A1,B1:C5,2)", "#N/A")
    res = _run(make_book("b", {"S": cells}), ir_for, tmp_path)
    assert _only(res, CellRef("S", 1, 4)) is ReasonKind.silently_wrong
    assert (res.checked, res.reproduced, res.share) == (1, 0, 0.0)


def test_trust_flags_wrong_number_as_silently_wrong(make_book, ir_for, tmp_path):
    res = _run(make_book("b", {"S": {"A1": 2, "B1": F("=A1*3", 7)}}), ir_for, tmp_path)
    kind = _only(res, CellRef("S", 1, 2))
    assert kind is ReasonKind.silently_wrong
    assert "7" in res.failures[0].reason.detail and "6" in res.failures[0].reason.detail


def test_trust_flags_engine_error_sheet_qualified_range(make_book, ir_for, tmp_path):
    book = make_book("b", {"S": {"A1": 1, "A2": 2, "A3": 3, "B1": F("=SUM(S!$A$1:S!$A3)", 6)}})
    res = _run(book, ir_for, tmp_path)
    assert _only(res, CellRef("S", 1, 2)) is ReasonKind.engine_error


def test_trust_different_error_code_is_silently_wrong(make_book, ir_for, tmp_path):
    res = _run(make_book("b", {"S": {"A1": F("=1/0", "#N/A")}}), ir_for, tmp_path)
    assert _only(res, CellRef("S", 1, 1)) is ReasonKind.silently_wrong


def test_trust_unsupported_function(make_book, ir_for, tmp_path):
    res = _run(make_book("b", {"S": {"A1": F("=NOSUCHFN(1)", 3)}}), ir_for, tmp_path)
    assert _only(res, CellRef("S", 1, 1)) is ReasonKind.unsupported_function


def test_trust_macro_function(make_book, ir_for, tmp_path):
    book = make_book("b", {"S": {"A1": F("=MyUdf(1)", 3)}})
    copy = prepare_engine_copy(book, tmp_path / "e.xlsx", as_of=None)
    eng = Engine.load(copy.path)
    eng.evaluate()
    graph = WorkbookGraph.from_ir(ir_for(book))
    graph.udf_cells = frozenset({CellRef("S", 1, 1)})
    res = run_trust(eng, graph, tolerance=1e-9, skip=set())
    assert _only(res, CellRef("S", 1, 1)) is ReasonKind.macro_function


def test_stored_missing_is_no_stored_value(make_book, ir_for, tmp_path):
    base = make_book("base", {"S": {"A1": 2, "B1": F("=A1*3", 6), "C1": F("=A1", 2)}})

    def drop_v(xml: str) -> str:
        return xml.replace("<v>6</v>", "", 1)

    book = patch_book(base, tmp_path / "nov.xlsx", edit={"xl/worksheets/sheet1.xml": drop_v})
    res = _run(book, ir_for, tmp_path)
    assert _only(res, CellRef("S", 1, 2)) is ReasonKind.no_stored_value
    assert (res.checked, res.reproduced) == (2, 1)


def test_trust_skip_cells_not_compared(make_book, ir_for, tmp_path):
    book = make_book("b", {"S": {"A1": 2, "B1": F("=A1*3", 7), "C1": F("=A1", 2)}})
    res = _run(book, ir_for, tmp_path, skip={CellRef("S", 1, 2)})
    assert (res.checked, res.reproduced, res.failures) == (1, 1, [])


def test_share_of_empty_run():
    assert TrustResult(checked=0, reproduced=0, failures=[]).share == 1.0
