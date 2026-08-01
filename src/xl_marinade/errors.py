"""Typed error hierarchy for XL Marinade.

Library code raises these; the CLI (``xl_marinade.cli``) is the only layer that
maps them to process exit codes.
"""


class MarinadeError(Exception):
    """Base class for all XL Marinade errors."""


class UnsupportedInput(MarinadeError):
    """The input is not a supported or readable Excel workbook."""


class MemoryBudgetExceeded(MarinadeError):
    """Extraction exceeded the configured memory budget."""


class ExtractionError(MarinadeError):
    """Extraction failed for a reason not covered by a more specific error."""


class DiffError(MarinadeError):
    """IR diff failed for a reason not covered by a more specific error.

    Wraps unexpected failures from the diff pipeline (e.g. an unreadable or
    corrupt IR database) so callers see a typed error instead of a raw
    ``sqlite3.Error``.
    """


class LLMUnavailable(MarinadeError):
    """An LLM feature was requested but no provider is configured/installed."""
