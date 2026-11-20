"""Tier boundary: calc imports without ironcalc; core and docs never import calc."""

import subprocess
import sys

import pytest

from xl_marinade.calc.models import CellRef


def test_calc_imports_without_ironcalc():
    # Subprocess so xl_marinade.calc is imported fresh (never already in sys.modules);
    # a None entry makes any `import ironcalc` raise ImportError.
    code = (
        "import sys\n"
        "sys.modules['ironcalc'] = None\n"
        "import xl_marinade.calc, xl_marinade.calc.models, xl_marinade.calc.engine\n"
    )
    proc = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True)
    assert proc.returncode == 0, proc.stderr


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
