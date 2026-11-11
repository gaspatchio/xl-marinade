"""The database records whether VBA was read (issues #48, #57).

Without `oletools` every VBA table was left empty and nothing in the database
said so, so a macro workbook extracted that way was indistinguishable from one
with no macros. `ir_metadata.vba_extraction` now records one of:

- `extracted` -- the VBA tables hold the workbook's VBA;
- `no_vba_project` -- the workbook has no VBA project (e.g. any `.xlsx`),
  whether or not oletools is installed;
- `skipped_no_oletools` -- the workbook has a VBA project, but the `[vba]`
  extra is not installed, so it was not read;
- `failed` -- reading or storing the VBA raised; the tables may be partial.

The extraction progress line says the same thing: agents read it, not the
database, and read "0 procs" as "no macros" (#57).
"""

import importlib
import sqlite3
import sys
import zipfile
from pathlib import Path
from types import ModuleType

import openpyxl
import pytest

import xl_marinade


def _module(name: str) -> ModuleType:
    # Resolved per test, not at import: other tests purge and re-import
    # xl_marinade modules, and the pipeline imports whatever is current.
    return importlib.import_module(name)


def _xlsx(path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["B1"] = 2, "=A1*3"
    wb.save(path)
    return path


def _status(db: Path) -> str:
    with sqlite3.connect(db) as conn:
        row = conn.execute("SELECT value FROM ir_metadata WHERE key = 'vba_extraction'").fetchone()
    assert row is not None, "ir_metadata has no vba_extraction key"
    return row[0]


def _xlsm(path: Path, vba_project: bytes = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1") -> Path:
    """A workbook carrying xl/vbaProject.bin (by default, just an OLE header)."""
    src = _xlsx(path.with_suffix(".src.xlsx"))
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(path, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("xl/vbaProject.bin", vba_project)
    return path


def _no_oletools(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setitem(sys.modules, "oletools", None)
    monkeypatch.setitem(sys.modules, "oletools.olevba", None)


def _vba_line(err: str) -> str:
    return next(line.strip() for line in err.splitlines() if line.strip().startswith("VBA:"))


def test_workbook_without_vba_project(tmp_path: Path, capfd: pytest.CaptureFixture[str]) -> None:
    pytest.importorskip("oletools.olevba")
    db = xl_marinade.extract(_xlsx(tmp_path / "book.xlsx"), tmp_path / "book.db")
    assert _status(db) == "no_vba_project"
    assert _vba_line(capfd.readouterr().err) == "VBA: none (no_vba_project)"


def test_workbook_without_vba_project_needs_no_oletools(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capfd: pytest.CaptureFixture[str]
) -> None:
    """The #57 reproduction: a plain .xlsx read as "macros not read" without oletools.

    Its status matched a macro workbook that was never read, so neither could
    be told from the other; an agent took both to mean "no macros".
    """
    _no_oletools(monkeypatch)
    db = xl_marinade.extract(_xlsx(tmp_path / "book.xlsx"), tmp_path / "book.db")
    assert _status(db) == "no_vba_project"
    err = capfd.readouterr().err
    assert _vba_line(err) == "VBA: none (no_vba_project)"
    assert "oletools" not in err  # nothing to warn about


def test_vba_project_without_oletools_says_not_read(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capfd: pytest.CaptureFixture[str]
) -> None:
    """The #48 case: a macro workbook on a plain install must not read as "no macros"."""
    _no_oletools(monkeypatch)
    db = xl_marinade.extract(_xlsm(tmp_path / "book.xlsm"), tmp_path / "book.db")
    assert _status(db) == "skipped_no_oletools"
    line = _vba_line(capfd.readouterr().err)
    assert line.startswith("VBA: present, NOT READ")
    assert "xl-marinade[vba]" in line
    assert "0 procs" not in line


def test_vba_project_with_oletools_is_extracted(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capfd: pytest.CaptureFixture[str]
) -> None:
    """With oletools importable, the VBA project is read (oletools stubbed)."""

    class StubParser:
        type = "OLE"

        def __init__(self, filename: str, data: bytes | None = None) -> None:
            pass

        def extract_all_macros(self) -> list[tuple[str, str, str, str]]:
            code = 'Attribute VB_Name = "Module1"\nSub Hello()\n    x = 1\nEnd Sub\n'
            return [("vbaProject.bin", "VBA/Module1", "Module1.bas", code)]

        def detect_vba_macros(self) -> bool:
            return False

        def close(self) -> None:
            pass

    olevba = ModuleType("oletools.olevba")
    olevba.VBA_Parser = StubParser  # type: ignore[attr-defined]
    monkeypatch.setitem(sys.modules, "oletools", ModuleType("oletools"))
    monkeypatch.setitem(sys.modules, "oletools.olevba", olevba)

    db = xl_marinade.extract(_xlsm(tmp_path / "book.xlsm"), tmp_path / "book.db")
    assert _status(db) == "extracted"
    with sqlite3.connect(db) as conn:
        assert conn.execute("SELECT name FROM vba_modules").fetchall() == [("Module1.bas",)]
    line = _vba_line(capfd.readouterr().err)
    assert line.startswith("VBA: 1 procs")
    assert line.endswith("(extracted)")


def test_extracted(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    # No macro workbook ships with the tests; stand in for a successful read.
    extractor = _module("xl_marinade.core.vba.extractor")
    monkeypatch.setattr(extractor, "extract_vba", lambda _path: extractor.VBAExtraction())
    db = xl_marinade.extract(_xlsx(tmp_path / "book.xlsx"), tmp_path / "book.db")
    assert _status(db) == "extracted"


def test_unreadable_vba_project_is_failed(tmp_path: Path) -> None:
    """A vbaProject.bin olevba cannot parse is a failure, not an absence."""
    pytest.importorskip("oletools.olevba")
    xlsm = _xlsm(tmp_path / "book.xlsm", b"not an OLE compound file")
    db = xl_marinade.extract(xlsm, tmp_path / "book.db")
    assert _status(db) == "failed"


def test_storage_failure_is_failed(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    extractor = _module("xl_marinade.core.vba.extractor")
    fep = _module("xl_marinade.core.new_arch.fast_extraction_pipeline")

    def boom(*_args: object) -> None:
        raise sqlite3.OperationalError("disk I/O error")

    monkeypatch.setattr(extractor, "extract_vba", lambda _path: extractor.VBAExtraction())
    monkeypatch.setattr(fep, "_store_vba_phase3_tables", boom)
    db = xl_marinade.extract(_xlsx(tmp_path / "book.xlsx"), tmp_path / "book.db")
    assert _status(db) == "failed"
