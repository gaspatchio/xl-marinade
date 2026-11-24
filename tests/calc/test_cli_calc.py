"""Tests for ``marinade calc``: exit codes, deterministic reruns, clean stdout."""

from __future__ import annotations

import json
from collections.abc import Callable
from pathlib import Path

import pytest
from typer.testing import CliRunner

pytest.importorskip("ironcalc")

from tests.calc.conftest import F  # noqa: E402
from xl_marinade.cli.main import app  # noqa: E402

Book = Callable[..., Path]

SHEETS: dict[str, dict[str, object]] = {
    "Inputs": {"A1": "age", "B1": 40},
    "Calc": {"B1": F("=IF(Inputs!B1>60,Inputs!B1*2,Inputs!B1)", 40)},
}
MAP = {"id_field": "id", "fields": {"age": "Inputs!B1"}}
ARTIFACTS = ("results.csv", "trust_report.json", "paths.csv", "run.json")


def _args(tmp: Path, book: Path, out: Path, *extra: str) -> list[str]:
    policies = tmp / "policies.csv"
    policies.write_text("id,age\n1,50\n2,70\n", encoding="utf-8")
    imap = tmp / "map.json"
    imap.write_text(json.dumps(MAP), encoding="utf-8")
    return [
        "calc",
        str(book),
        "--policies",
        str(policies),
        "--map",
        str(imap),
        "--outputs",
        "Calc!B1",
        "-o",
        str(out),
        *extra,
    ]


def test_calc_succeeds_and_stdout_is_clean(make_book: Book, tmp_path: Path) -> None:
    book = make_book("b", SHEETS)
    result = CliRunner().invoke(app, _args(tmp_path, book, tmp_path / "out"))
    assert result.exit_code == 0, result.output
    assert result.stdout == ""
    assert "wrote" in result.stderr
    for name in ARTIFACTS:
        assert (tmp_path / "out" / name).exists()


def test_two_runs_are_byte_identical(make_book: Book, tmp_path: Path) -> None:
    book = make_book("b", SHEETS)
    runner = CliRunner()
    for d in ("a", "b"):
        res = runner.invoke(app, _args(tmp_path, book, tmp_path / d, "--as-of", "2025-01-31"))
        assert res.exit_code == 0, res.output
    for name in ARTIFACTS:
        assert (tmp_path / "a" / name).read_bytes() == (tmp_path / "b" / name).read_bytes(), name


def test_bad_input_exits_1(make_book: Book, tmp_path: Path) -> None:
    book = make_book("b", SHEETS)
    args = _args(tmp_path, book, tmp_path / "out")
    args[args.index("Calc!B1")] = "Nope!A1"
    assert CliRunner().invoke(app, args).exit_code == 1


def test_bad_as_of_exits_1(make_book: Book, tmp_path: Path) -> None:
    book = make_book("b", SHEETS)
    res = CliRunner().invoke(app, _args(tmp_path, book, tmp_path / "out", "--as-of", "31/01/2025"))
    assert res.exit_code == 1
    assert "as-of" in res.stderr


def test_refused_workbook_exits_6(make_book: Book, tmp_path: Path) -> None:
    book = make_book("b1904", SHEETS, date1904=True)
    res = CliRunner().invoke(app, _args(tmp_path, book, tmp_path / "out"))
    assert res.exit_code == 6
    assert "1904" in res.stderr


@pytest.mark.usefixtures("blocked_ironcalc")
def test_missing_engine_exits_5(make_book: Book, tmp_path: Path) -> None:
    book = make_book("b", SHEETS)
    res = CliRunner().invoke(app, _args(tmp_path, book, tmp_path / "out"))
    assert res.exit_code == 5
    assert "xl-marinade[calc]" in res.stderr


def test_help_is_experimental() -> None:
    res = CliRunner().invoke(app, ["calc", "--help"])
    assert res.exit_code == 0
    assert "Experimental:" in res.stdout
