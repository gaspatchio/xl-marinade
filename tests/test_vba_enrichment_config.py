"""`extract --enrich` (VBA enrichment) must honour the same BYOK config as
`document --enrich`.

VBA enrichment used to build its own ``openai.OpenAI`` client from a hardcoded
model and ``OPENAI_API_KEY`` only, bypassing the shared factory seam. That made
``LLM_BASE_URL`` (the local/Azure/proxy endpoint override) and ``OPENAI_MODEL``
silently ineffective for the ``extract --enrich`` path — a user who pointed
``LLM_BASE_URL`` at a local endpoint would still egress to OpenAI. These tests
pin the fix: it routes through ``factory.make_llm_client`` and honours
``OPENAI_MODEL`` (keeping a cheap default when unset).

The module is patched by object (not by dotted string): several sibling tests
delete ``xl_marinade.*`` from ``sys.modules`` to prove import-cleanliness, which
would break a string-path monkeypatch target that runs after them.
"""

import json


class _FakeCompletions:
    def __init__(self, captured: dict) -> None:
        self._captured = captured

    def create(self, **kwargs: object):
        self._captured["model"] = kwargs.get("model")

        class _Msg:
            content = json.dumps({"reads": [], "writes": [], "description": "does a thing"})

        class _Choice:
            message = _Msg()

        class _Resp:
            choices = [_Choice()]

        return _Resp()


class _FakeChat:
    def __init__(self, captured: dict) -> None:
        self.completions = _FakeCompletions(captured)


class _FakeClient:
    def __init__(self, captured: dict) -> None:
        self.chat = _FakeChat(captured)


def test_enrich_procedure_routes_through_the_factory_and_honours_model(monkeypatch):
    """It obtains its client from the shared factory seam and uses OPENAI_MODEL."""
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    monkeypatch.delenv("LLM_API_KEY", raising=False)

    from xl_marinade.llm import vba_enrichment

    captured: dict = {}
    made: dict = {}

    def _fake_make_llm_client(api_key=None):
        made["api_key_arg"] = api_key
        return _FakeClient(captured)

    monkeypatch.setattr(vba_enrichment, "make_llm_client", _fake_make_llm_client)
    monkeypatch.setenv("OPENAI_MODEL", "local-model-x")

    result = vba_enrichment.enrich_procedure("Mod.Proc", "Selection.Copy", "wb ctx", "static ctx")

    assert "api_key_arg" in made, "must obtain its client from factory.make_llm_client"
    assert captured.get("model") == "local-model-x", "must honour OPENAI_MODEL"
    assert result is not None
    assert result.model_used == "local-model-x"


def test_enrich_procedure_defaults_to_the_cheap_model_when_unset(monkeypatch):
    """With OPENAI_MODEL unset it keeps the cheap per-procedure default."""
    monkeypatch.delenv("OPENAI_MODEL", raising=False)

    from xl_marinade.llm import vba_enrichment

    captured: dict = {}
    monkeypatch.setattr(
        vba_enrichment, "make_llm_client", lambda api_key=None: _FakeClient(captured)
    )

    result = vba_enrichment.enrich_procedure("Mod.Proc", "Selection.Copy", "wb ctx", "static ctx")

    assert captured.get("model") == "gpt-4.1-nano"
    assert result is not None


def test_enrich_procedure_degrades_without_a_client(monkeypatch):
    """No configured client (no key) degrades to None — enrichment never raises."""
    from xl_marinade.llm import vba_enrichment

    monkeypatch.setattr(vba_enrichment, "make_llm_client", lambda api_key=None: None)

    result = vba_enrichment.enrich_procedure("Mod.Proc", "Selection.Copy", "wb ctx", "static ctx")

    assert result is None
