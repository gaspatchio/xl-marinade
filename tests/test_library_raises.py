"""Library code raises typed exceptions instead of calling sys.exit."""

import pytest

from xl_marinade.core.new_arch.memory_budget import (
    MemoryBudgetConfig,
    MemoryBudgetController,
)
from xl_marinade.errors import MemoryBudgetExceeded


def test_memory_budget_raises_not_exits():
    # max 1 MB with immediate sampling: the test process RSS is always well above
    # 1 MB, so check() must raise a typed error rather than sys.exit the process.
    controller = MemoryBudgetController(MemoryBudgetConfig(max_memory_mb=1, check_interval_rows=1))
    if not controller.config.enabled:
        pytest.skip("psutil not available; memory monitoring disabled")

    with pytest.raises(MemoryBudgetExceeded):
        controller.check(row_count=10)


def test_diff_bad_db_raises_marinade_error(tmp_path):
    """core.api.diff wraps loader/sqlite failures as a typed MarinadeError."""
    from xl_marinade.core.api import diff
    from xl_marinade.errors import MarinadeError

    garbage = tmp_path / "not_a.db"
    garbage.write_text("this is not a sqlite database")

    with pytest.raises(MarinadeError):
        diff(str(garbage), str(garbage))


def test_diff_verification_error_is_marinade_error():
    """DiffVerificationError joins the typed hierarchy so the CLI can map it to exit 1."""
    from xl_marinade.core.ir_diff.model import DiffVerificationError
    from xl_marinade.errors import MarinadeError

    assert issubclass(DiffVerificationError, MarinadeError)


def test_document_bad_db_raises_marinade_error(tmp_path):
    """docs.document rejects a non-SQLite input with a typed MarinadeError.

    Regression: a garbage file surfaced as a raw sqlite3.DatabaseError
    traceback from deep inside the labeller, while `diff` on the same input
    produced a clean typed error.
    """
    from xl_marinade.docs import document
    from xl_marinade.errors import MarinadeError

    garbage = tmp_path / "not_a.db"
    garbage.write_text("this is not a sqlite database")

    with pytest.raises(MarinadeError):
        document(garbage, tmp_path / "out")


def test_document_non_ir_db_raises_marinade_error(tmp_path):
    """docs.document rejects a SQLite file that is not an IR database.

    Regression: pointing document at an arbitrary SQLite file surfaced as a
    raw RuntimeError traceback ("agent_bindings view not found") instead of a
    typed error.
    """
    import sqlite3

    from xl_marinade.docs import document
    from xl_marinade.errors import MarinadeError

    other = tmp_path / "other.db"
    conn = sqlite3.connect(other)
    conn.execute("CREATE TABLE t (x)")
    conn.commit()
    conn.close()

    with pytest.raises(MarinadeError):
        document(other, tmp_path / "out")


def test_document_partial_db_raises_marinade_error(tmp_path):
    """docs.document rejects a partially-written IR database with a typed error.

    Regression: an extraction interrupted mid-write (Ctrl-C, kill) can leave a
    database where the agent_bindings VIEW exists but its backing tables do
    not. Existence-checking the view passed validation and the labeller then
    surfaced a raw RuntimeError traceback; validation must probe the view with
    a real query.
    """
    import sqlite3

    from xl_marinade.docs import document
    from xl_marinade.errors import MarinadeError

    partial = tmp_path / "partial.db"
    conn = sqlite3.connect(partial)
    conn.execute("CREATE VIEW agent_bindings AS SELECT * FROM bindings")
    conn.commit()
    conn.close()

    with pytest.raises(MarinadeError):
        document(partial, tmp_path / "out")


def test_document_zero_binding_extraction_raises_marinade_error(tmp_path):
    """docs.document gives a typed error for an extraction with no bindings.

    Regression: documenting a values-only or macro-only workbook (extracts
    fine, zero bindings — e.g. randwalk2.xlsm from corpus testing) crashed
    with a raw ValueError traceback from the labeller.
    """
    import openpyxl

    from xl_marinade.core.api import extract
    from xl_marinade.docs import document
    from xl_marinade.errors import MarinadeError

    wb = openpyxl.Workbook()
    xlsx = tmp_path / "empty.xlsx"
    wb.save(xlsx)
    ir_db = extract(xlsx, tmp_path / "ir.db")

    with pytest.raises(MarinadeError):
        document(ir_db, tmp_path / "out")
