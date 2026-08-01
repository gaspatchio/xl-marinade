"""llm.document() orchestrates enrich-in-the-middle and degrades without a key."""

import json

from test_workbook_generator.cli import create_comprehensive_test_workbook
from xl_marinade.core.api import extract


def test_llm_document_degrades_without_key(tmp_path, monkeypatch):
    monkeypatch.delenv("LLM_API_KEY", raising=False)
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)

    xlsx = tmp_path / "wb.xlsx"
    create_comprehensive_test_workbook(xlsx)
    ir_db = extract(xlsx, tmp_path / "ir.db")

    from xl_marinade.llm import document  # requires the [llm] extra (openai installed in dev)

    out = tmp_path / "out"
    md = document(ir_db, out)  # no key -> deterministic, must NOT raise

    assert md == out / "documentation.md"
    assert md.exists() and md.stat().st_size > 0
    assert (out / "model_spec.json").exists()


def test_llm_document_enriched_path_runs_end_to_end(tmp_path):
    """document() with a real EnrichmentProvider must drive the enriched path for real.

    ``test_llm_document_degrades_without_key`` above only exercises the *degraded*
    branch (no provider -> deterministic docs only). It never runs
    ``_build_overlay -> run_sprint7_pipeline -> _render`` with a provider present, so
    two things go unverified: (a) whether the paid tier's headline feature actually
    completes, and (b) whether the IR that ``xl_marinade.core.api.extract`` produces
    satisfies ``run_sprint7_pipeline``'s hard requirement
    (``validate_fast_extraction`` raises ``ValueError`` unless
    ``ir_metadata.build_mode == "fast"``).

    We use ``FixtureProvider`` pointed at a fixture directory that has **no fixture
    files in it** (deliberately, not by omission): every ``EnrichmentProvider``
    method then falls through to ``FixtureProvider``'s own built-in deterministic
    defaults --- "no_change" for label/structural proposals, and a stable heuristic
    tag for segmentation (which must return a tag from ``allowed_tags`` on every
    call since Stage 3 is not triage-gated). This is the simplest deterministic,
    no-network provider that satisfies the real response-shape contracts documented
    in ``enrichment_service.py`` without hand-guessing them, and it is the same
    class real callers pass via ``--llm-provider fixture``.

    A prior manual run against this same fixture-workbook (see task report)
    confirmed the provider is genuinely exercised at volume: all 39 bindings run
    through ``enrich_label``, 20 through ``propose_structural_fix``, and all 39
    through ``propose_segmentation_tag`` -- not a trivial 0-selected no-op path.
    """
    from xl_marinade.llm import document  # requires the [llm] extra
    from xl_marinade.llm.enrichment_service import FixtureProvider

    xlsx = tmp_path / "wb.xlsx"
    create_comprehensive_test_workbook(xlsx)
    ir_db = extract(xlsx, tmp_path / "ir.db")

    # Deliberately empty/nonexistent fixture dir: FixtureProvider degrades each
    # method to its own built-in deterministic default (see class docstring above).
    provider = FixtureProvider(tmp_path / "fixtures_empty")

    out = tmp_path / "out"
    md = document(ir_db, out, provider=provider)  # must run the enriched composition, not raise

    assert md == out / "documentation.md"
    assert md.exists() and md.stat().st_size > 0
    assert (out / "model_spec.json").exists() and (out / "model_spec.json").stat().st_size > 0

    # Prove the sprint7 stages actually ran against a nonempty binding set (not a
    # 0-selected/early-return no-op): triage scoring and segmentation tagging both
    # produce artifacts sized to the number of bindings in the IR.
    triage_label = json.loads((out / "triage_label.json").read_text())
    assert len(triage_label) > 0

    segmentation_db = out / "segmentation.db"
    assert segmentation_db.exists()

    # No provider-response-shape mismatches were rejected by validation.
    rejected = json.loads((out / "rejected_proposals.json").read_text())
    assert rejected == []


def test_default_provider_honours_llm_api_key(monkeypatch):
    """A key in LLM_API_KEY alone must build a provider, not raise (never-raise invariant)."""
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    monkeypatch.setenv("LLM_API_KEY", "sk-dummy-not-used")

    from xl_marinade.llm._document import _default_provider

    prov = _default_provider()  # must NOT raise ValueError
    assert prov is not None
    assert prov.get_provider_name()


def test_enrichment_failure_degrades_to_deterministic(tmp_path, monkeypatch):
    """If sprint7 raises mid-run, we still render the deterministic doc from the built overlay."""
    import xl_marinade.llm._document as docmod
    from xl_marinade.llm.enrichment_service import FixtureProvider

    xlsx = tmp_path / "wb.xlsx"
    create_comprehensive_test_workbook(xlsx)
    ir_db = extract(xlsx, tmp_path / "ir.db")

    def _boom(_config):
        raise RuntimeError("sprint7 exploded")

    monkeypatch.setattr(docmod, "run_sprint7_pipeline", _boom)

    out = tmp_path / "out"
    md = docmod.document(ir_db, out, provider=FixtureProvider(tmp_path / "empty"))

    assert md == out / "documentation.md"
    assert md.exists() and md.stat().st_size > 0
    assert (out / "model_spec.json").exists()


def test_document_rejects_non_provider(tmp_path):
    """A non-EnrichmentProvider passed as `provider` fails fast at the boundary."""
    import pytest

    import xl_marinade.llm._document as docmod

    xlsx = tmp_path / "wb.xlsx"
    create_comprehensive_test_workbook(xlsx)
    ir_db = extract(xlsx, tmp_path / "ir.db")

    with pytest.raises(TypeError):
        docmod.document(ir_db, tmp_path / "out", provider=object())
