"""Enriched documentation: labelling -> sprint7 enrichment -> render.

Degrades to the deterministic docs pipeline when no LLM key/provider is available.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from loguru import logger

from xl_marinade.docs.pipeline import _build_overlay, _render
from xl_marinade.llm.sprint7_pipeline import Sprint7PipelineConfig, run_sprint7_pipeline

if TYPE_CHECKING:
    from xl_marinade.llm.enrichment_service import EnrichmentProvider


def _default_provider() -> EnrichmentProvider | None:
    """Build the default EnrichmentProvider if a key is configured, else None.

    Probes with ``resolve_api_key`` (no client is built-then-discarded). No key -> None ->
    the caller degrades to the deterministic docs pipeline.
    """
    from xl_marinade.llm.factory import resolve_api_key

    if resolve_api_key() is None:
        return None
    from xl_marinade.llm.enrichment_service import OpenAIProvider

    return OpenAIProvider()


def document(ir_db: Path, out_dir: Path, *, provider: EnrichmentProvider | None = None) -> Path:
    """Generate enriched documentation for an extracted IR database.

    ``provider`` is an ``EnrichmentProvider`` (e.g. ``enrichment_service.OpenAIProvider`` or
    ``FixtureProvider``). When ``None``, a default provider is built from the configured key;
    if no key is configured this degrades to the deterministic docs pipeline and NEVER raises.
    Writes ``documentation.md`` + ``model_spec.json``; returns the markdown path.
    """
    ir_db = Path(ir_db)
    out_dir = Path(out_dir)

    if provider is not None:
        from xl_marinade.llm.enrichment_service import EnrichmentProvider

        if not isinstance(provider, EnrichmentProvider):
            raise TypeError(
                f"provider must be an EnrichmentProvider, got {type(provider).__name__}"
            )

    engine, overlay_db, _ = _build_overlay(ir_db, out_dir)

    prov = provider if provider is not None else _default_provider()
    if prov is None:
        logger.info("No LLM provider — deterministic documentation only")
    else:
        logger.info("Enriching overlay via sprint7 ({})", prov.get_provider_name())
        config = Sprint7PipelineConfig(
            ir_db_path=str(ir_db),
            output_dir=out_dir,
            no_llm=False,
            llm_provider=prov,
        )
        try:
            run_sprint7_pipeline(config)
        except Exception as exc:  # noqa: BLE001 - never lose the recoverable deterministic doc
            logger.warning(
                "Enrichment failed ({}); falling back to deterministic documentation", exc
            )

    return _render(engine, overlay_db, ir_db, out_dir)
