"""The core makes no LLM/network call unless enrichment is explicitly opted into.

A key in the environment must never, on its own, trigger egress.
"""

from test_workbook_generator.cli import create_comprehensive_test_workbook


def test_no_openai_call_when_key_present_but_not_opted_in(tmp_path, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-should-never-be-used")

    calls = []
    import xl_marinade.llm.vba_enrichment as le

    monkeypatch.setattr(le, "enrich_and_store", lambda *a, **k: calls.append(1) or {})

    xlsx = tmp_path / "s.xlsx"
    create_comprehensive_test_workbook(xlsx)

    import xl_marinade

    xl_marinade.extract(xlsx, tmp_path / "ir.db")  # enrich defaults False

    assert calls == [], "VBA LLM enrichment must not run without explicit opt-in"
