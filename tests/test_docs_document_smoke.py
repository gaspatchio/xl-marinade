"""docs.document() produces documentation.md + model_spec.json deterministically."""

from test_workbook_generator.cli import create_comprehensive_test_workbook
from xl_marinade.core.api import extract  # Tier-0 extractor -> ir.db


def test_document_writes_markdown_and_spec(tmp_path):
    xlsx = tmp_path / "wb.xlsx"
    create_comprehensive_test_workbook(xlsx)
    ir_db = extract(xlsx, tmp_path / "ir.db")

    from xl_marinade.docs import document

    out = tmp_path / "out"
    md = document(ir_db, out)

    assert md == out / "documentation.md"
    assert md.exists() and md.stat().st_size > 0
    assert (out / "model_spec.json").exists()
