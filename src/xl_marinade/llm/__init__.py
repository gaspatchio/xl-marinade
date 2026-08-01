"""XL Marinade LLM add-on (Tier 1, bring-your-own-key).

Optional layer over the deterministic core. Install with::

    pip install xl-marinade[llm]

Importing this package fails fast with an actionable message if the extra's
dependencies (``openai``) are not installed, so the free core never depends on it.
"""

try:  # the [llm] extra guard
    import openai as _openai  # noqa: F401
except ImportError as _exc:  # pragma: no cover - exercised in a subprocess test
    raise ImportError(
        "The XL Marinade LLM features require the optional add-on: pip install xl-marinade[llm]"
    ) from _exc

from typing import Any

# `document` is provided lazily via __getattr__ (PEP 562), so static analysers
# can't see it as a module attribute — the export is intentional.
__all__ = ["document"]  # pyright: ignore[reportUnsupportedDunderAll]


def __getattr__(name: str) -> Any:  # PEP 562 — lazy attribute access
    """Expose ``document`` without eagerly importing the docs + sprint7 + jinja2 chain."""
    if name == "document":
        from xl_marinade.llm._document import document

        globals()["document"] = document  # cache so later lookups skip __getattr__
        return document
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
