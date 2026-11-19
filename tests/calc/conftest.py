"""Fixtures for calc tests: synthetic workbooks that look like Excel saved them."""

from __future__ import annotations

import builtins
import re
import shutil
import sys
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import NamedTuple
from xml.sax.saxutils import escape

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import xl_marinade


class F(NamedTuple):
    """A formula cell with the value Excel would have cached for it."""

    formula: str
    cached: object


_CELL_RE = re.compile(r'<c r="(?P<ref>[A-Z]+[0-9]+)"(?P<attrs>[^>]*?)>(?P<body>.*?)</c>', re.S)


def _is_formula_cell(v: object) -> bool:
    # Duck-typed: tests import F via ``tests.calc.conftest`` while pytest loads this
    # file as ``conftest``, so ``isinstance`` could see two distinct classes.
    return hasattr(v, "formula") and hasattr(v, "cached")


def _patch_sheet_xml(xml: str, cached: dict[str, object]) -> str:
    def repl(m: re.Match[str]) -> str:
        ref = m.group("ref")
        if ref not in cached:
            return m.group(0)
        value = cached[ref]
        attrs = re.sub(r'\s+t="[^"]*"', "", m.group("attrs"))
        body = re.sub(r"<v>.*?</v>|<v\s*/>", "", m.group("body"), flags=re.S)
        if isinstance(value, bool):
            attrs += ' t="b"'
            v = "1" if value else "0"
        elif isinstance(value, str) and value.startswith("#"):
            attrs += ' t="e"'
            v = escape(value)
        elif isinstance(value, str):
            attrs += ' t="str"'
            v = escape(value)
        else:
            v = repr(value)
        return f'<c r="{ref}"{attrs}>{body}<v>{v}</v></c>'

    return _CELL_RE.sub(repl, xml)


@pytest.fixture
def make_book(tmp_path: Path) -> Callable[..., Path]:
    def _make(
        name: str,
        sheets: dict[str, dict[str, object]],
        *,
        names: dict[str, str] | None = None,
        date1904: bool = False,
        full_precision: bool = True,
        iterate: bool = False,
        manual: bool = False,
        validations: list[tuple[str, list[str]]] | None = None,
    ) -> Path:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        cached_by_sheet: dict[str, dict[str, object]] = {}
        for sheet_name, cells in sheets.items():
            ws = wb.create_sheet(sheet_name)
            cached: dict[str, object] = {}
            for ref, val in cells.items():
                if _is_formula_cell(val):
                    ws[ref] = val.formula  # type: ignore[attr-defined]
                    cached[ref] = val.cached  # type: ignore[attr-defined]
                else:
                    ws[ref] = val
            cached_by_sheet[sheet_name] = cached
        for key, ref in (names or {}).items():
            wb.defined_names[key] = DefinedName(key, attr_text=ref)
        if date1904:
            wb.epoch = openpyxl.utils.datetime.CALENDAR_MAC_1904
        wb.calculation.fullPrecision = full_precision
        wb.calculation.iterate = iterate
        wb.calculation.calcMode = "manual" if manual else "auto"
        for target, allowed in validations or []:
            sheet_name, _, cell = target.rpartition("!")
            dv = DataValidation(type="list", formula1='"' + ",".join(allowed) + '"')
            wb[sheet_name.strip("'").replace("''", "'")].add_data_validation(dv)
            dv.add(cell)

        path = tmp_path / f"{name}.xlsx"
        raw = tmp_path / f"{name}.raw.xlsx"
        wb.save(raw)

        # Map sheet names to part names, then rewrite those parts with cached values.
        order = list(sheets)
        with zipfile.ZipFile(raw) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                m = re.fullmatch(r"xl/worksheets/sheet([0-9]+)\.xml", item.filename)
                if m is not None:
                    cached = cached_by_sheet[order[int(m.group(1)) - 1]]
                    data = _patch_sheet_xml(data.decode("utf-8"), cached).encode("utf-8")
                zout.writestr(item, data)
        raw.unlink()
        return path

    return _make


@pytest.fixture
def ir_for(tmp_path: Path) -> Callable[[Path], Path]:
    def _ir(path: Path) -> Path:
        out_dir = tmp_path / f"ir_{path.stem}"
        shutil.rmtree(out_dir, ignore_errors=True)
        out_dir.mkdir()
        return xl_marinade.extract(path, out_dir / "ir.db")

    return _ir


@pytest.fixture
def blocked_ironcalc(monkeypatch):
    real_import = builtins.__import__

    def guard(name, *args, **kwargs):
        if name == "ironcalc" or name.startswith("ironcalc."):
            raise ModuleNotFoundError(f"blocked in test: {name}")
        return real_import(name, *args, **kwargs)

    for mod in list(sys.modules):
        if mod == "ironcalc" or mod.startswith("ironcalc."):
            monkeypatch.delitem(sys.modules, mod, raising=False)
    monkeypatch.setattr(builtins, "__import__", guard)
