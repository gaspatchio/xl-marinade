"""The public library surface: xl_marinade.extract() and the error types."""

from pathlib import Path

import xl_marinade
from test_workbook_generator.cli import create_comprehensive_test_workbook


def test_public_surface():
    assert callable(xl_marinade.extract)
    assert callable(xl_marinade.diff)
    assert issubclass(xl_marinade.errors.ExtractionError, xl_marinade.errors.MarinadeError)
    assert issubclass(xl_marinade.errors.DiffError, xl_marinade.errors.MarinadeError)


def test_extract_writes_a_database(tmp_path):
    xlsx = tmp_path / "s.xlsx"
    create_comprehensive_test_workbook(xlsx)

    out = xl_marinade.extract(xlsx, tmp_path / "ir.db")

    assert Path(out).exists()
    assert out == tmp_path / "ir.db"
