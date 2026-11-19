"""Tests for reading workbook calculation settings."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import pytest

from xl_marinade.calc.settings import WorkbookSettings, read_settings, refusal_reasons
from xl_marinade.errors import CalcInputError

Book = Callable[..., Path]
SHEETS: dict[str, dict[str, object]] = {"S": {"A1": 1}}


def test_default_book(make_book: Book) -> None:
    s = read_settings(make_book("d", SHEETS))
    assert s == WorkbookSettings(
        date1904=False, full_precision=True, iterate=False, manual_calc=False
    )
    assert refusal_reasons(s) == []


def test_date1904(make_book: Book) -> None:
    s = read_settings(make_book("d", SHEETS, date1904=True))
    assert s.date1904 is True
    assert refusal_reasons(s) == ["1904 date system"]


def test_precision_as_displayed(make_book: Book) -> None:
    s = read_settings(make_book("p", SHEETS, full_precision=False))
    assert s.full_precision is False
    assert refusal_reasons(s) == ["precision as displayed"]


def test_iterate_is_read_but_not_refused(make_book: Book) -> None:
    s = read_settings(make_book("i", SHEETS, iterate=True))
    assert s.iterate is True
    assert refusal_reasons(s) == []


def test_manual_calc_not_refused(make_book: Book) -> None:
    s = read_settings(make_book("m", SHEETS, manual=True))
    assert s.manual_calc is True
    assert refusal_reasons(s) == []


def test_both_reasons() -> None:
    s = WorkbookSettings(date1904=True, full_precision=False, iterate=False, manual_calc=False)
    assert refusal_reasons(s) == ["1904 date system", "precision as displayed"]


def test_not_a_zip(tmp_path: Path) -> None:
    p = tmp_path / "x.xlsx"
    p.write_text("nope")
    with pytest.raises(CalcInputError):
        read_settings(p)
