"""The single LLM-client seam (bring-your-own-key).

All LLM calls in the add-on go through :func:`make_llm_client`, so provider
configuration lives in exactly one place. The ``LLM_BASE_URL`` override is the
whole multi-provider story for v0.1: it points the OpenAI-compatible client at
Azure OpenAI, a local vLLM/Ollama server, or a LiteLLM proxy.
"""

from __future__ import annotations

import os
from typing import Any

from xl_marinade.errors import LLMUnavailable


def provider_name() -> str:
    """The configured provider id (for audit logs)."""
    return os.getenv("LLM_PROVIDER", "openai")


def resolve_api_key(api_key: str | None = None) -> str | None:
    """Resolve the BYOK key: explicit arg, else ``LLM_API_KEY``, else ``OPENAI_API_KEY``."""
    return api_key or os.getenv("LLM_API_KEY") or os.getenv("OPENAI_API_KEY")


def make_llm_client(api_key: str | None = None) -> Any | None:
    """Return a configured OpenAI-compatible client, or ``None`` if no key is set.

    Env: ``LLM_PROVIDER`` (``openai`` | ``azure`` | ``openai_compatible``),
    ``LLM_API_KEY`` (falls back to ``OPENAI_API_KEY``), ``LLM_BASE_URL`` (optional
    endpoint override). Returning ``None`` on no key lets callers degrade to the
    deterministic path instead of failing.
    """
    key = resolve_api_key(api_key)
    if not key:
        return None
    try:
        import openai
    except ImportError as exc:  # pragma: no cover - exercised via the [llm] guard
        raise LLMUnavailable(
            "LLM features require the optional add-on: pip install xl-marinade[llm]"
        ) from exc

    kwargs: dict[str, Any] = {"api_key": key}
    base_url = os.getenv("LLM_BASE_URL")
    if base_url:
        kwargs["base_url"] = base_url
    return openai.OpenAI(**kwargs)


def require_llm_client(api_key: str | None = None) -> Any:
    """Like :func:`make_llm_client` but raises ``LLMUnavailable`` if unconfigured."""
    client = make_llm_client(api_key)
    if client is None:
        raise LLMUnavailable(
            "No LLM API key configured. Set LLM_API_KEY (or OPENAI_API_KEY), "
            "and optionally LLM_BASE_URL for Azure/local endpoints."
        )
    return client
