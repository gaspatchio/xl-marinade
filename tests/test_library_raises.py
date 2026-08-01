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
