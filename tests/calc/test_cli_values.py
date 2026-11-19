"""CLI tests for ``marinade values``."""

from __future__ import annotations

import json
from collections.abc import Callable
from pathlib import Path

from typer.testing import CliRunner

from tests.calc.conftest import F
from xl_marinade.cli.main import app

runner = CliRunner()


def test_cli_values_exit_codes(make_book: Callable[..., Path], tmp_path: Path) -> None:
    p = make_book("c", {"S": {"A1": F("=1", 1)}})
    good = tmp_path / "good.json"
    good.write_text(json.dumps({"v": {"cell": "S!A1"}}))
    bad = tmp_path / "bad.json"
    bad.write_text("{not json")
    ok = runner.invoke(app, ["values", str(p), "--variables", str(good), "-o", str(tmp_path / "o")])
    assert ok.exit_code == 0, ok.output
    assert (tmp_path / "o" / "rows.csv").exists()
    fail = runner.invoke(
        app, ["values", str(p), "--variables", str(bad), "-o", str(tmp_path / "o2")]
    )
    assert fail.exit_code == 1
    assert "not valid JSON" in fail.output


def test_cli_values_help_is_experimental() -> None:
    res = runner.invoke(app, ["values", "--help"])
    assert res.exit_code == 0
    assert "Experimental" in res.output
