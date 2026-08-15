"""XL Marinade — deterministic Excel formula-graph extraction.

The public API is re-exported here. Importing this package has no side effects
and makes no network calls (see tests/test_import_cleanliness.py).
"""

__version__ = "0.3.0"

from xl_marinade import errors
from xl_marinade.core.api import diff, extract

__all__ = ["diff", "extract", "errors", "__version__"]
