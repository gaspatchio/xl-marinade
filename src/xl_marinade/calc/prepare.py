"""The engine copy: a workbook IronCalc must recalculate rather than read back.

Every formula cell's cached ``<v>`` is deleted, What-If data tables are removed
(IronCalc will not load them), ``TODAY()``/``NOW()`` become a fixed ``DATE`` when an
as-of date is given, and external references are replaced by the value Excel cached
for them inside the workbook.
"""

from __future__ import annotations

import html
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from loguru import logger

from xl_marinade.calc.models import CellRef, _col_to_number
from xl_marinade.errors import CalcInputError

_NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

_C_RE = re.compile(r"<c\b([^>]*?)(/>|>(.*?)</c>)", re.S)
_V_RE = re.compile(r"<v\b[^>]*>.*?</v>|<v\b[^>]*/>", re.S)
_T_RE = re.compile(r'\st="[^"]*"')
_F_RE = re.compile(r"<f\b([^>]*?)(?:/>|>(.*?)</f>)", re.S)
_ATTR_RE = re.compile(r'\b([A-Za-z0-9_:]+)="([^"]*)"')
_A1_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")
_NUM_RE = re.compile(r"^-?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][-+]?[0-9]+)?$")

# One scan over formula text. String literals and quoted sheet names are matched first so
# nothing inside them is rewritten; then single-cell external references; then TODAY/NOW.
_TOKEN_RE = re.compile(
    r'(?P<str>"(?:[^"]|"")*")'
    r"|(?P<ext>(?:'\[(?P<n1>[0-9]+)\](?P<s1>(?:[^']|'')+)'|\[(?P<n2>[0-9]+)\](?P<s2>[\w.]+))"
    r"!(?P<ref>\$?[A-Za-z]{1,3}\$?[0-9]+)(?P<range>:\$?[A-Za-z]{1,3}\$?[0-9]+)?(?![\w(])"
    r")"
    r"|(?P<qname>'(?:[^']|'')*')"
    r"|(?P<today>(?<![\w.])(?i:TODAY|NOW)\s*\(\s*\))"
)

_ExtCache = dict[int, dict[str, dict[tuple[int, int], str]]]


@dataclass(slots=True)
class PreparedCopy:
    """The cache-free engine copy and what was changed to make it."""

    path: Path
    stripped: int
    data_table_cells: list[CellRef] = field(default_factory=list)
    held_externals: dict[CellRef, list[str]] = field(default_factory=dict)


def _xml_unescape(text: str) -> str:
    """Decode named and numeric character references (``&amp;``, ``&#10;``, ...)."""
    return html.unescape(text)


def _xml_escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _a1(text: str) -> tuple[int, int] | None:
    m = _A1_RE.match(text)
    if m is None:
        return None
    return int(m.group(2)), _col_to_number(m.group(1))


def _rels(z: zipfile.ZipFile, part: str) -> dict[str, str]:
    """Relationship id -> target part path (resolved against ``part``'s folder)."""
    folder, name = posixpath.split(part)
    rels_path = posixpath.join(folder, "_rels", f"{name}.rels")
    if rels_path not in z.namelist():
        return {}
    root = ET.fromstring(z.read(rels_path))
    out: dict[str, str] = {}
    for rel in root.findall("pr:Relationship", _NS):
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target", "")
        resolved = target.lstrip("/") if target.startswith("/") else posixpath.join(folder, target)
        out[rel.get("Id", "")] = posixpath.normpath(resolved)
    return out


def _workbook_parts(z: zipfile.ZipFile) -> tuple[dict[str, str], list[str]]:
    """(worksheet part -> sheet name, external-link parts in formula ``[n]`` order)."""
    root = ET.fromstring(z.read("xl/workbook.xml"))
    rels = _rels(z, "xl/workbook.xml")
    sheets: dict[str, str] = {}
    for el in root.findall("m:sheets/m:sheet", _NS):
        part = rels.get(el.get(_R_ID, ""))
        if part is not None:
            sheets[part] = el.get("name", "")
    links = [
        rels.get(el.get(_R_ID, ""), "")
        for el in root.findall("m:externalReferences/m:externalReference", _NS)
    ]
    return sheets, links


def _literal(value: str, cell_type: str | None) -> str:
    """The formula literal for a cached external value."""
    if cell_type in {"str", "s", "inlineStr"}:
        return '"' + value.replace('"', '""') + '"'
    if cell_type == "b":
        return "TRUE" if value.strip() in {"1", "true", "TRUE"} else "FALSE"
    if cell_type == "e":
        return value.strip()
    text = value.strip()
    if not _NUM_RE.match(text):
        return '"' + value.replace('"', '""') + '"'
    return f"({text})" if text.startswith("-") else text


def _external_cache(z: zipfile.ZipFile, link_parts: list[str]) -> _ExtCache:
    """``[n]`` -> sheet name (casefolded) -> (row, col) -> formula literal."""
    names = set(z.namelist())
    cache: _ExtCache = {}
    count = max(len(link_parts), sum(1 for n in names if n.startswith("xl/externalLinks/ext")))
    for n in range(1, count + 1):
        part = link_parts[n - 1] if n <= len(link_parts) and link_parts[n - 1] else ""
        if part not in names:
            part = f"xl/externalLinks/externalLink{n}.xml"
        if part not in names:
            continue
        root = ET.fromstring(z.read(part))
        book = root.find("m:externalBook", _NS)
        if book is None:
            continue
        sheet_names = [s.get("val", "") for s in book.findall("m:sheetNames/m:sheetName", _NS)]
        per_sheet: dict[str, dict[tuple[int, int], str]] = {}
        for data in book.findall("m:sheetDataSet/m:sheetData", _NS):
            sid = int(data.get("sheetId", "-1"))
            if not 0 <= sid < len(sheet_names):
                continue
            values: dict[tuple[int, int], str] = {}
            for cell in data.findall("m:row/m:cell", _NS):
                pos = _a1(cell.get("r", ""))
                v = cell.find("m:v", _NS)
                if pos is None or v is None or v.text is None:
                    continue
                values[pos] = _literal(v.text, cell.get("t"))
            per_sheet[sheet_names[sid].casefold()] = values
        cache[n] = per_sheet
    return cache


class _Rewriter:
    """Rewrites formula text: TODAY/NOW and cached external references."""

    def __init__(self, as_of: date | None, cache: _ExtCache) -> None:
        self._date = f"DATE({as_of.year},{as_of.month},{as_of.day})" if as_of else None
        self._cache = cache

    def rewrite(self, formula: str, *, shared: bool) -> tuple[str, list[str]]:
        """Return the new text and the external references replaced.

        In a shared-formula master only fully absolute references are replaced: a
        relative one shifts in every cell that shares the formula.
        """
        held: list[str] = []

        def repl(m: re.Match[str]) -> str:
            if m.group("today") is not None:
                return self._date if self._date is not None else m.group(0)
            if m.group("ext") is None or m.group("range") is not None:
                return m.group(0)
            ref = m.group("ref")
            if shared and ref.count("$") < 2:
                return m.group(0)
            n = int(m.group("n1") or m.group("n2"))
            sheet = (m.group("s1") or "").replace("''", "'") or m.group("s2")
            pos = _a1(ref)
            literal = self._cache.get(n, {}).get(sheet.casefold(), {}).get(pos or (0, 0))
            if literal is None:
                return m.group(0)
            held.append(m.group("ext"))
            return literal

        return _TOKEN_RE.sub(repl, formula), held


@dataclass(slots=True)
class _SheetResult:
    xml: str
    stripped: int
    data_table_cells: list[CellRef]
    held: dict[CellRef, list[str]]


def _rect_cells(sheet: str, ref: str) -> list[CellRef]:
    first, _, last = ref.partition(":")
    a, b = _a1(first), _a1(last or first)
    if a is None or b is None:
        return []
    return [
        CellRef(sheet, r, c)
        for r in range(min(a[0], b[0]), max(a[0], b[0]) + 1)
        for c in range(min(a[1], b[1]), max(a[1], b[1]) + 1)
    ]


def _cell_pos(attrs: str) -> tuple[int, int] | None:
    for key, val in _ATTR_RE.findall(attrs):
        if key == "r":
            return _a1(val)
    return None


def _prepare_sheet(xml: str, sheet: str, rewriter: _Rewriter) -> _SheetResult:
    # Pass 1: data-table rectangles, and array/spill rectangles. Only an array's anchor
    # holds the <f>; its other cells hold Excel's results as bare <v>, which the engine
    # would otherwise load as constants.
    table_cells: set[CellRef] = set()
    array_cells: set[CellRef] = set()
    for f_attrs, _ in _F_RE.findall(xml):
        attrs = dict(_ATTR_RE.findall(f_attrs))
        if attrs.get("t") == "dataTable" and "ref" in attrs:
            table_cells.update(_rect_cells(sheet, attrs["ref"]))
        if attrs.get("t") == "array" and "ref" in attrs:
            array_cells.update(_rect_cells(sheet, attrs["ref"]))

    stripped = 0
    held: dict[CellRef, list[str]] = {}
    shared_held: dict[str, list[str]] = {}
    shared_children: dict[str, list[CellRef]] = defaultdict(list)

    def fix_cell(m: re.Match[str]) -> str:
        nonlocal stripped
        attrs, body = m.group(1), m.group(3)
        pos = _cell_pos(attrs)
        cell = CellRef(sheet, *pos) if pos is not None else None
        if body is None:
            return m.group(0)
        if cell is not None and cell in table_cells:
            body = _F_RE.sub(lambda fm: "" if 't="dataTable"' in fm.group(1) else fm.group(0), body)
            return "<c" + _T_RE.sub("", attrs) + ">" + _V_RE.sub("", body) + "</c>"
        if "<f" not in body:
            if cell is None or cell not in array_cells:
                return m.group(0)
            stripped += 1
            return "<c" + _T_RE.sub("", attrs) + ">" + _V_RE.sub("", body) + "</c>"
        stripped += 1

        def fix_f(fm: re.Match[str]) -> str:
            f_attrs, text = fm.group(1), fm.group(2)
            attrs_d = dict(_ATTR_RE.findall(f_attrs))
            shared = attrs_d.get("t") == "shared"
            si = attrs_d.get("si", "")
            if text is None or text == "":
                if shared and cell is not None:
                    shared_children[si].append(cell)
                return fm.group(0)
            plain = _xml_unescape(text)
            new, refs = rewriter.rewrite(plain, shared=shared)
            if refs and cell is not None:
                held[cell] = refs
                if shared:
                    shared_held[si] = refs
            if new == plain:
                return fm.group(0)  # unchanged: keep the original bytes
            return f"<f{f_attrs}>{_xml_escape(new)}</f>"

        body = _F_RE.sub(fix_f, body)
        return "<c" + _T_RE.sub("", attrs) + ">" + _V_RE.sub("", body) + "</c>"

    out = _C_RE.sub(fix_cell, xml)
    for si, refs in shared_held.items():
        for child in shared_children.get(si, []):
            held.setdefault(child, list(refs))
    return _SheetResult(out, stripped, sorted(table_cells), held)


_DEFINED_NAME_RE = re.compile(r"(<definedName\b[^>]*>)(.*?)(</definedName>)", re.S)


def _prepare_workbook_xml(xml: str, rewriter: _Rewriter) -> str:
    """Apply the formula rewrites to every ``<definedName>`` (e.g. ``AsOf = TODAY()``)."""

    def fix(m: re.Match[str]) -> str:
        plain = _xml_unescape(m.group(2))
        new, refs = rewriter.rewrite(plain, shared=False)
        if new == plain:
            return m.group(0)
        if refs:
            logger.debug("defined name {} holds external references {}", m.group(1), refs)
        return m.group(1) + _xml_escape(new) + m.group(3)

    return _DEFINED_NAME_RE.sub(fix, xml)


def prepare_engine_copy(src: Path, dst: Path, *, as_of: date | None) -> PreparedCopy:
    """Write the engine copy of ``src`` to ``dst``.

    The same rewrites apply to defined names in ``xl/workbook.xml``. External references
    with no cached value are left as they are: IronCalc then gives an error (``#REF!``,
    or ``#ERROR!`` when the book has an external-link part) and the trust run reports
    the cell. ``stripped`` counts formula cells plus the non-anchor cells of array ranges.

    Raises:
        CalcInputError: if ``src`` is not a readable xlsx package.
    """
    try:
        zin = zipfile.ZipFile(src)
    except (OSError, zipfile.BadZipFile) as exc:
        raise CalcInputError(f"cannot read workbook {src}: {exc}") from exc
    stripped = 0
    table_cells: list[CellRef] = []
    held: dict[CellRef, list[str]] = {}
    with zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        try:
            sheets, link_parts = _workbook_parts(zin)
            rewriter = _Rewriter(as_of, _external_cache(zin, link_parts))
        except (KeyError, ET.ParseError) as exc:
            raise CalcInputError(f"cannot read workbook structure of {src}: {exc}") from exc
        for item in zin.infolist():
            data = zin.read(item.filename)
            sheet = sheets.get(item.filename)
            if item.filename == "xl/workbook.xml":
                data = _prepare_workbook_xml(data.decode("utf-8"), rewriter).encode("utf-8")
            elif sheet is not None:
                res = _prepare_sheet(data.decode("utf-8"), sheet, rewriter)
                data = res.xml.encode("utf-8")
                stripped += res.stripped
                table_cells.extend(res.data_table_cells)
                held.update(res.held)
            zout.writestr(item, data)
    logger.debug(
        "engine copy {}: {} formula cells stripped, {} data-table cells, {} held externals",
        dst,
        stripped,
        len(table_cells),
        len(held),
    )
    return PreparedCopy(dst, stripped, sorted(table_cells), dict(sorted(held.items())))
