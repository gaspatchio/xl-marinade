"""Skeleton-level guarantees for the xl_marinade package."""

import importlib.metadata
import sys


def test_import_exposes_version():
    sys.modules.pop("xl_marinade", None)
    import xl_marinade

    assert isinstance(xl_marinade.__version__, str)
    # Pin the constant to the installed distribution so the two cannot drift:
    # the wheel reporting one version while __version__ reports another is a
    # provenance bug, not an annoyance (review finding on the 0.2.0 bump).
    assert xl_marinade.__version__ == importlib.metadata.version("xl-marinade")


def test_errors_base_class_exists():
    from xl_marinade.errors import MarinadeError

    assert issubclass(MarinadeError, Exception)
