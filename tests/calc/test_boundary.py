"""Tier boundary: calc imports without ironcalc; core and docs never import calc."""

import builtins
import subprocess
import sys

import pytest

from xl_marinade.calc.models import CellRef


@pytest.fixture
def blocked_ironcalc(monkeypatch):
    real_import = builtins.__import__

    def guard(name, *args, **kwargs):
        if name == "ironcalc" or name.startswith("ironcalc."):
            raise ModuleNotFoundError(f"blocked in test: {name}")
        return real_import(name, *args, **kwargs)

    for mod in list(sys.modules):
        if mod == "ironcalc" or mod.startswith("ironcalc."):
            monkeypatch.delitem(sys.modules, mod, raising=False)
    monkeypatch.setattr(builtins, "__import__", guard)


def test_calc_imports_without_ironcalc(blocked_ironcalc):
    import xl_marinade.calc  # noqa: F401


def test_engine_load_without_ironcalc_raises(blocked_ironcalc, tmp_path):
    from xl_marinade.calc.engine import Engine
    from xl_marinade.errors import CalcUnavailable

    with pytest.raises(CalcUnavailable, match=r"xl-marinade\[calc\]"):
        Engine.load(tmp_path / "x.xlsx")


def test_core_and_docs_do_not_import_calc():
    code = (
        "import sys, xl_marinade, xl_marinade.docs\n"
        "bad = [m for m in sys.modules if m.startswith(('xl_marinade.calc', 'ironcalc'))]\n"
        "assert not bad, bad\n"
    )
    proc = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True)
    assert proc.returncode == 0, proc.stderr


def test_cellref_quotes_sheet_with_trailing_space():
    assert CellRef("Life Table ", 9, 1).a1() == "'Life Table '!A9"
    assert CellRef.parse("'It''s'!$B$3") == CellRef("It's", 3, 2)


def test_cellref_parse_rejects_unqualified():
    from xl_marinade.errors import CalcInputError

    with pytest.raises(CalcInputError):
        CellRef.parse("A1")


def test_workbook_refused_carries_reasons():
    from xl_marinade.errors import CalcError, WorkbookRefused

    err = WorkbookRefused("refused", ["circular"])
    assert isinstance(err, CalcError)
    assert err.reasons == ["circular"]
