"""The Typer CLI is a thin adapter: it extracts, and maps errors to exit codes."""

import json

from typer.testing import CliRunner

from test_workbook_generator.cli import create_comprehensive_test_workbook
from xl_marinade.cli.main import app

runner = CliRunner()


def test_cli_extract_ok(tmp_path):
    xlsx = tmp_path / "s.xlsx"
    create_comprehensive_test_workbook(xlsx)

    result = runner.invoke(app, ["extract", str(xlsx), "-o", str(tmp_path / "ir.db")])

    assert result.exit_code == 0, result.output
    assert (tmp_path / "ir.db").exists()


def test_cli_bad_input_nonzero_exit(tmp_path):
    bad = tmp_path / "nope.xlsx"
    bad.write_text("not a workbook")

    result = runner.invoke(app, ["extract", str(bad), "-o", str(tmp_path / "ir.db")])

    assert result.exit_code != 0


def test_cli_error_output_preserves_bracketed_text(tmp_path, monkeypatch):
    """Error messages must reach the user verbatim — rich markup swallowed
    bracketed text like the ``xl-marinade[llm]`` extras name, rendering
    install hints as just ``xl-marinade``."""
    from xl_marinade.errors import MarinadeError

    def boom(*args, **kwargs):
        raise MarinadeError("requires the optional add-on: pip install xl-marinade[llm]")

    import xl_marinade

    monkeypatch.setattr(xl_marinade, "extract", boom)
    wb = tmp_path / "s.xlsx"
    wb.write_bytes(b"")

    result = runner.invoke(app, ["extract", str(wb)])

    assert result.exit_code == 1
    assert "xl-marinade[llm]" in result.output


def test_cli_document_deterministic(tmp_path):
    xlsx = tmp_path / "s.xlsx"
    create_comprehensive_test_workbook(xlsx)
    assert runner.invoke(app, ["extract", str(xlsx), "-o", str(tmp_path / "ir.db")]).exit_code == 0

    out = tmp_path / "docs"
    result = runner.invoke(app, ["document", str(tmp_path / "ir.db"), "-o", str(out)])

    assert result.exit_code == 0, result.output
    assert (out / "documentation.md").exists()
    assert (out / "model_spec.json").exists()


def test_cli_document_enrich_degrades_without_key(tmp_path, monkeypatch):
    monkeypatch.delenv("LLM_API_KEY", raising=False)
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    xlsx = tmp_path / "s.xlsx"
    create_comprehensive_test_workbook(xlsx)
    assert runner.invoke(app, ["extract", str(xlsx), "-o", str(tmp_path / "ir.db")]).exit_code == 0

    out = tmp_path / "docs"
    result = runner.invoke(app, ["document", "--enrich", str(tmp_path / "ir.db"), "-o", str(out)])

    assert result.exit_code == 0, result.output  # no key -> degrades, never fails
    assert (out / "documentation.md").exists()


def test_cli_document_enrich_without_extra_degrades_to_docs(tmp_path, monkeypatch):
    """--enrich with the [llm] extra absent degrades to the docs tier (exit 0), not the llm tier."""
    import importlib.util

    import xl_marinade.docs

    real_find_spec = importlib.util.find_spec

    def fake_find_spec(name, *args, **kwargs):
        # Simulate the [llm] extra being absent even though openai is installed in dev.
        if name == "openai":
            return None
        return real_find_spec(name, *args, **kwargs)

    monkeypatch.setattr(importlib.util, "find_spec", fake_find_spec)

    docs_called = {}
    real_docs_document = xl_marinade.docs.document

    def spy_docs_document(ir_db, out_dir):
        docs_called["yes"] = True
        return real_docs_document(ir_db, out_dir)

    # The llm tier calls docs.pipeline internals, never this package-level function,
    # so this spy fires only when the CLI took the deterministic docs branch.
    monkeypatch.setattr(xl_marinade.docs, "document", spy_docs_document)

    xlsx = tmp_path / "s.xlsx"
    create_comprehensive_test_workbook(xlsx)
    assert runner.invoke(app, ["extract", str(xlsx), "-o", str(tmp_path / "ir.db")]).exit_code == 0

    out = tmp_path / "docs"
    result = runner.invoke(app, ["document", "--enrich", str(tmp_path / "ir.db"), "-o", str(out)])

    assert result.exit_code == 0, result.output
    assert docs_called.get("yes"), "extra-absent --enrich must degrade to the docs tier"
    assert (out / "documentation.md").exists()


def test_cli_diff_emits_json(tmp_path):
    xlsx = tmp_path / "s.xlsx"
    create_comprehensive_test_workbook(xlsx)
    a, b = tmp_path / "a.db", tmp_path / "b.db"
    assert runner.invoke(app, ["extract", str(xlsx), "-o", str(a)]).exit_code == 0
    assert runner.invoke(app, ["extract", str(xlsx), "-o", str(b)]).exit_code == 0

    out = tmp_path / "diff.json"
    result = runner.invoke(app, ["diff", str(a), str(b), "-o", str(out)])

    assert result.exit_code == 0, result.output
    data = json.loads(out.read_text())
    assert data["version"] == "1.0"
    assert "changes" in data
    assert data["changes"] == []  # identical DBs -> no changes


def test_cli_diff_bad_db_exits_cleanly(tmp_path):
    """A corrupt/non-IR DB fed to `diff` yields a clean exit 1, not an unhandled traceback."""
    garbage = tmp_path / "garbage.db"
    garbage.write_text("not a sqlite database")

    result = runner.invoke(app, ["diff", str(garbage), str(garbage)])

    assert result.exit_code == 1, result.output
    # The command must map the failure itself (typed error -> typer.Exit), not let a
    # raw sqlite3.Error / DiffVerificationError escape as an unhandled exception.
    assert result.exception is None or isinstance(result.exception, SystemExit), (
        f"unhandled {type(result.exception).__name__}: {result.exception}"
    )


def test_cli_extract_max_memory_flag_plumbed(tmp_path, monkeypatch):
    """--max-memory-mb reaches the library; without it a user who hits the
    1800 MB default budget on a large workbook has no CLI remedy at all
    (found on a real 40 MB / 2.3M-cell model that peaks at ~3.4 GB)."""
    import xl_marinade

    captured = {}

    def fake_extract(workbook, out, enrich=False, max_memory_mb=1800):
        captured["max_memory_mb"] = max_memory_mb
        return out

    monkeypatch.setattr(xl_marinade, "extract", fake_extract)
    wb = tmp_path / "s.xlsx"
    wb.write_bytes(b"")

    assert runner.invoke(app, ["extract", str(wb), "--max-memory-mb", "6000"]).exit_code == 0
    assert captured["max_memory_mb"] == 6000

    assert runner.invoke(app, ["extract", str(wb)]).exit_code == 0
    assert captured["max_memory_mb"] == 1800
