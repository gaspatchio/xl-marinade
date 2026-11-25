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
        local_names: dict[tuple[str, str], str] | None = None,
        list_sources: list[tuple[str, str]] | None = None,
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
        # Sheet-scoped names, keyed (sheet, name): Excel's ``localSheetId`` names.
        for (sheet_name, key), ref in (local_names or {}).items():
            wb[sheet_name].defined_names[key] = DefinedName(key, attr_text=ref)
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
        # List validations whose source is a reference, a name or a formula (formula1
        # as Excel writes it, without ``=``).
        for target, source in list_sources or []:
            sheet_name, _, cell = target.rpartition("!")
            dv = DataValidation(type="list", formula1=source)
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


def patch_book(
    src: Path,
    dst: Path,
    *,
    edit: dict[str, Callable[[str], str]] | None = None,
    add: dict[str, str] | None = None,
) -> Path:
    """Copy an xlsx, rewriting parts with ``edit`` callables and adding ``add`` parts.

    For XML openpyxl cannot write: data tables, external-link caches.
    """
    edit = edit or {}
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in edit:
                data = edit[item.filename](data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
        for name, text in (add or {}).items():
            zout.writestr(name, text.encode("utf-8"))
    return dst


_SSML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RELNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def add_external_link(src: Path, dst: Path, sheet: str, cells: dict[str, str]) -> Path:
    """Attach ``externalLink1.xml`` caching ``cells`` (``A1`` -> ``<v>`` text) for ``sheet``.

    Values that parse as numbers are cached as numbers, anything else as text.
    """

    def cell_xml(ref: str, v: str) -> str:
        try:
            float(v)
            return f'<cell r="{ref}"><v>{v}</v></cell>'
        except ValueError:
            return f'<cell r="{ref}" t="str"><v>{escape(v)}</v></cell>'

    rows: dict[int, list[str]] = {}
    for ref, v in cells.items():
        row = int(re.sub(r"[A-Z]+", "", ref))
        rows.setdefault(row, []).append(cell_xml(ref, v))
    sheet_data = "".join(f'<row r="{r}">{"".join(cs)}</row>' for r, cs in sorted(rows.items()))
    link = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<externalLink xmlns="{_SSML}" xmlns:r="{_RELNS}"><externalBook r:id="rId1">'
        f'<sheetNames><sheetName val="{escape(sheet)}"/></sheetNames>'
        f'<sheetDataSet><sheetData sheetId="0">{sheet_data}</sheetData></sheetDataSet>'
        f"</externalBook></externalLink>"
    )
    link_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/externalLinkPath" Target="other.xlsx" TargetMode="External"/>'
        "</Relationships>"
    )

    def wb(xml: str) -> str:
        root_tag = xml[xml.index("<workbook") : xml.index(">", xml.index("<workbook"))]
        if "xmlns:r=" not in root_tag:
            xml = xml.replace("<workbook ", f'<workbook xmlns:r="{_RELNS}" ', 1)
        refs = '<externalReferences><externalReference r:id="rIdExt1"/></externalReferences>'
        return xml.replace("</sheets>", "</sheets>" + refs, 1)

    def rels(xml: str) -> str:
        rel = (
            '<Relationship Id="rIdExt1" Type="http://schemas.openxmlformats.org/officeDocument/'
            '2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>'
        )
        return xml.replace("</Relationships>", rel + "</Relationships>", 1)

    def types(xml: str) -> str:
        ov = (
            '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>'
        )
        return xml.replace("</Types>", ov + "</Types>", 1)

    return patch_book(
        src,
        dst,
        edit={
            "xl/workbook.xml": wb,
            "xl/_rels/workbook.xml.rels": rels,
            "[Content_Types].xml": types,
        },
        add={
            "xl/externalLinks/externalLink1.xml": link,
            "xl/externalLinks/_rels/externalLink1.xml.rels": link_rels,
        },
    )
