import pytest

pytest.importorskip("ironcalc")

from tests.calc.conftest import F  # noqa: E402
from xl_marinade.calc.engine import Engine  # noqa: E402
from xl_marinade.calc.models import CellRef  # noqa: E402
from xl_marinade.errors import CalcError, CalcInputError  # noqa: E402


def test_engine_reads_sets_and_recalculates(make_book):
    book = make_book("b", {"S": {"A1": 2, "B1": F("=A1*3", 6), "C1": F("=1/0", "#DIV/0!")}})
    eng = Engine.load(book)
    eng.evaluate()
    a1, b1 = CellRef("S", 1, 1), CellRef("S", 1, 2)
    assert eng.value(b1) == 6.0
    assert eng.value(CellRef("S", 1, 3)) == "#DIV/0!"
    assert eng.value(CellRef("S", 9, 9)) is None
    eng.set_value(a1, 5.0)
    eng.evaluate()
    assert eng.value(b1) == 15.0
    eng.set_value(a1, "x")
    eng.evaluate()
    assert eng.value(a1) == "x" and eng.value(b1) == "#VALUE!"
    eng.set_value(a1, True)
    eng.evaluate()
    assert eng.value(a1) is True and eng.value(b1) == 3.0
    eng.set_formula(CellRef("S", 2, 1), "=B1+1")
    eng.set_formula(CellRef("S", 3, 1), "B1+2")
    eng.evaluate()
    assert eng.value(CellRef("S", 2, 1)) == 4.0 and eng.value(CellRef("S", 3, 1)) == 5.0
    assert eng.version == "0.9.1"


def test_engine_hidden_sheet(make_book):
    eng = Engine.load(make_book("b", {"S": {"A1": 2}}))
    eng.add_hidden_sheet("_marinade_probes")
    probe = CellRef("_marinade_probes", 1, 1)
    eng.set_formula(probe, "=S!A1*2")
    eng.evaluate()
    assert eng.value(probe) == 4.0
    props = {p["name"]: p["state"] for p in eng._model.get_worksheets_properties()}
    assert props["_marinade_probes"] == "hidden"


def test_engine_unknown_sheet_is_input_error(make_book):
    eng = Engine.load(make_book("b", {"S": {"A1": 2}}))
    with pytest.raises(CalcInputError, match="Nope"):
        eng.value(CellRef("Nope", 1, 1))


def test_engine_load_failure_is_typed(tmp_path):
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a zip")
    with pytest.raises(CalcError, match="bad.xlsx"):
        Engine.load(bad)
