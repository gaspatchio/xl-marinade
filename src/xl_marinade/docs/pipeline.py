"""ABOUTME: Deterministic documentation pipeline — labelling, classification, confidence, spec, markdown.

Pass-4 confidence runs AFTER classification (and after enrichment, when the llm
tier composes _build_overlay/_render around sprint7).
"""

from __future__ import annotations

from pathlib import Path

from loguru import logger

from xl_marinade.docs.apply_actuarial_classification import apply_actuarial_classification
from xl_marinade.docs.generators.markdown import MarkdownGenerator
from xl_marinade.docs.json_spec_generator import generate_json_spec
from xl_marinade.docs.two_pass_labeller import TwoPassLabellingEngine


def _build_overlay(ir_db: Path, out_dir: Path) -> tuple[TwoPassLabellingEngine, Path, Path]:
    """Run labelling (Passes 1-3) + actuarial classification. Returns (engine, overlay_db, mutations_json)."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    ir_db = Path(ir_db)
    mutations_json = out_dir / "mutations.json"
    overlay_db = out_dir / "semantic_overlay.db"

    logger.info("Labelling (Passes 1-3) from IR {}", ir_db)
    engine = TwoPassLabellingEngine(str(ir_db))
    engine.load_bindings()
    engine.build_graph()
    # Pass 4 confidence is deferred to _render (it needs classification data).
    engine.run_labelling(
        str(mutations_json),
        str(overlay_db),
        enable_sheet_refinement=True,
        run_confidence_assessment=False,
    )

    logger.info("Actuarial classification")
    apply_actuarial_classification(str(overlay_db), str(ir_db), mutations_path=str(mutations_json))
    return engine, overlay_db, mutations_json


def _render(engine: TwoPassLabellingEngine, overlay_db: Path, ir_db: Path, out_dir: Path) -> Path:
    """Pass-4 confidence + JSON spec + markdown. Returns the documentation.md path."""
    out_dir = Path(out_dir)
    ir_db = Path(ir_db)
    mutations_json = out_dir / "mutations.json"
    spec_json = out_dir / "model_spec.json"
    doc_md = out_dir / "documentation.md"

    logger.info("Pass-4 confidence assessment")
    engine.run_pass_4_confidence_assessment(
        str(overlay_db), str(mutations_json), confidence_threshold=0.7
    )

    logger.info("JSON spec generation")
    generate_json_spec(str(overlay_db), str(ir_db), str(spec_json))

    logger.info("Markdown documentation")
    MarkdownGenerator(None).generate_to_file(str(spec_json), str(doc_md))
    return doc_md


def document(ir_db: Path, out_dir: Path) -> Path:
    """Generate deterministic documentation for an extracted IR database.

    Writes ``documentation.md`` and ``model_spec.json`` under ``out_dir``.
    Returns the path to ``documentation.md``. No network, no LLM.
    """
    engine, overlay_db, _ = _build_overlay(ir_db, out_dir)
    return _render(engine, overlay_db, ir_db, out_dir)


if __name__ == "__main__":  # thin shim; real CLI subcommand lands in Step 3
    import argparse

    parser = argparse.ArgumentParser(description="Run deterministic documentation pipeline")
    parser.add_argument("--ir-db", required=True)
    parser.add_argument("--output-dir", required=True)
    ns = parser.parse_args()
    print(document(Path(ns.ir_db), Path(ns.output_dir)))
