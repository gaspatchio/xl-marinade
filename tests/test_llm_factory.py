"""The BYOK provider factory: one seam, base_url override, None-without-key."""


def test_factory_returns_none_without_key(monkeypatch):
    monkeypatch.delenv("LLM_API_KEY", raising=False)
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)

    from xl_marinade.llm.factory import make_llm_client

    assert make_llm_client() is None


def test_require_raises_without_key(monkeypatch):
    import pytest

    monkeypatch.delenv("LLM_API_KEY", raising=False)
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)

    from xl_marinade.errors import LLMUnavailable
    from xl_marinade.llm.factory import require_llm_client

    with pytest.raises(LLMUnavailable):
        require_llm_client()


def test_factory_honours_base_url(monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test-not-used")
    monkeypatch.setenv("LLM_BASE_URL", "http://localhost:11434/v1")

    from xl_marinade.llm.factory import make_llm_client

    client = make_llm_client()
    assert client is not None
    assert str(client.base_url).startswith("http://localhost:11434")


def test_resolve_api_key_is_public(monkeypatch):
    monkeypatch.delenv("LLM_API_KEY", raising=False)
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)

    from xl_marinade.llm.factory import resolve_api_key

    assert resolve_api_key("explicit") == "explicit"
    assert resolve_api_key() is None
    monkeypatch.setenv("LLM_API_KEY", "sk-env")
    assert resolve_api_key() == "sk-env"
