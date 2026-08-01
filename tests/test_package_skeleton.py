"""Skeleton-level guarantees for the xl_marinade package."""

import sys


def test_import_exposes_version():
    sys.modules.pop("xl_marinade", None)
    import xl_marinade

    assert isinstance(xl_marinade.__version__, str)
    assert xl_marinade.__version__


def test_errors_base_class_exists():
    from xl_marinade.errors import MarinadeError

    assert issubclass(MarinadeError, Exception)
