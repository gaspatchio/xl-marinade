"""The database records whether VBA was read (issues #48, #57).

`ir_metadata.vba_extraction` records one of:

- `extracted` -- the VBA tables hold the workbook's VBA;
- `no_vba_project` -- the workbook has no VBA project (e.g. any `.xlsx`);
- `failed` -- reading or storing the VBA raised; the tables may be partial.

Since 0.7.0 VBA is read with olefile alone, so `skipped_no_oletools` -- the
`[vba]` extra not installed -- is no longer produced. The extraction progress
line says the same thing as the status: agents read it, not the database.
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


def _hide_oletools(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setitem(sys.modules, "oletools", None)
    monkeypatch.setitem(sys.modules, "oletools.olevba", None)


def _vba_line(err: str) -> str:
    return next(line.strip() for line in err.splitlines() if line.strip().startswith("VBA:"))


def test_workbook_without_vba_project(tmp_path: Path, capfd: pytest.CaptureFixture[str]) -> None:
    """The #57 reproduction: a plain .xlsx is "no macros", never "not read"."""
    db = xl_marinade.extract(_xlsx(tmp_path / "book.xlsx"), tmp_path / "book.db")
    assert _status(db) == "no_vba_project"
    assert _vba_line(capfd.readouterr().err) == "VBA: none (no_vba_project)"


def test_vba_project_is_read_without_oletools(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capfd: pytest.CaptureFixture[str]
) -> None:
    """A VBA project is read with oletools unimportable: the [vba] extra is gone.

    The repo ships no macro workbook, so the module list is stood in for at
    the reader's boundary (the reader itself is tested in test_ovba.py); the
    rest of the extraction -- procedure parsing, storage, status -- is real.
    """
    _hide_oletools(monkeypatch)
    extractor = _module("xl_marinade.core.vba.extractor")
    ovba = _module("xl_marinade.core.vba.ovba")
    code = 'Attribute VB_Name = "Module1"\r\nSub Hello()\r\n    x = 1\r\nEnd Sub\r\n'
    monkeypatch.setattr(extractor, "is_ole", lambda _data: True)
    monkeypatch.setattr(
        extractor,
        "read_vba_modules",
        lambda _src: [ovba.VBAModuleSource("Module1", "Module1.bas", "VBA/Module1", code)],
    )

    db = xl_marinade.extract(_xlsm(tmp_path / "book.xlsm"), tmp_path / "book.db")
    assert _status(db) == "extracted"
    with sqlite3.connect(db) as conn:
        assert conn.execute("SELECT name FROM vba_modules").fetchall() == [("Module1.bas",)]
        assert conn.execute("SELECT name FROM vba_procedures").fetchall() == [("Hello",)]
    line = _vba_line(capfd.readouterr().err)
    assert line.startswith("VBA: 1 procs")
    assert line.endswith("(extracted)")
    assert "oletools" not in sys.modules or sys.modules["oletools"] is None


def test_damaged_vba_project_is_failed(tmp_path: Path, capfd: pytest.CaptureFixture[str]) -> None:
    """An OLE header with nothing behind it: present but unreadable."""
    db = xl_marinade.extract(_xlsm(tmp_path / "book.xlsm"), tmp_path / "book.db")
    assert _status(db) == "failed"
    assert _vba_line(capfd.readouterr().err).startswith("VBA: extraction FAILED")


def test_skipped_no_oletools_is_never_recorded(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _hide_oletools(monkeypatch)
    for name, build in (("a.xlsx", _xlsx), ("b.xlsm", _xlsm)):
        db = xl_marinade.extract(build(tmp_path / name), tmp_path / f"{name}.db")
        assert _status(db) != "skipped_no_oletools"


def test_unreadable_vba_project_is_failed(tmp_path: Path) -> None:
    """A vbaProject.bin that is not OLE at all is a failure, not an absence."""
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
