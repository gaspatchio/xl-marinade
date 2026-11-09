"""The database records whether VBA was read (issue #48).

Without `oletools` every VBA table was left empty and nothing in the database
said so, so a macro workbook extracted that way was indistinguishable from one
with no macros. `ir_metadata.vba_extraction` now records one of:

- `extracted` -- the VBA tables hold the workbook's VBA;
- `no_vba_project` -- the workbook has no VBA project (e.g. any `.xlsx`);
- `skipped_no_oletools` -- the `[vba]` extra is not installed, VBA not read;
- `failed` -- reading or storing the VBA raised; the tables may be partial.
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


def test_workbook_without_vba_project(tmp_path: Path) -> None:
    pytest.importorskip("oletools.olevba")
    db = xl_marinade.extract(_xlsx(tmp_path / "book.xlsx"), tmp_path / "book.db")
    assert _status(db) == "no_vba_project"


def test_missing_oletools_is_recorded(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """The #48 reproduction: a plain install must not read as "no macros"."""
    monkeypatch.setitem(sys.modules, "oletools", None)
    monkeypatch.setitem(sys.modules, "oletools.olevba", None)
    db = xl_marinade.extract(_xlsx(tmp_path / "book.xlsx"), tmp_path / "book.db")
    assert _status(db) == "skipped_no_oletools"


def test_extracted(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    # No macro workbook ships with the tests; stand in for a successful read.
    extractor = _module("xl_marinade.core.vba.extractor")
    monkeypatch.setattr(extractor, "extract_vba", lambda _path: extractor.VBAExtraction())
    db = xl_marinade.extract(_xlsx(tmp_path / "book.xlsx"), tmp_path / "book.db")
    assert _status(db) == "extracted"


def test_unreadable_vba_project_is_failed(tmp_path: Path) -> None:
    """A vbaProject.bin olevba cannot parse is a failure, not an absence."""
    pytest.importorskip("oletools.olevba")
    src = _xlsx(tmp_path / "book.xlsx")
    xlsm = tmp_path / "book.xlsm"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(xlsm, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("xl/vbaProject.bin", b"not an OLE compound file")
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
