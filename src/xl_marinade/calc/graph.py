"""The formula graph and stored values, loaded from the extraction IR."""

from __future__ import annotations

import hashlib
import json
import re
import sqlite3
from bisect import bisect_left
from collections import defaultdict, deque
from collections.abc import Iterable
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Literal

from xl_marinade.calc.formula_ast import NAME, SHEET
from xl_marinade.calc.models import CellRef, _col_to_number
from xl_marinade.calc.prepare import relative_names
from xl_marinade.core.db_uri import connect_read_only
from xl_marinade.errors import CalcInputError

StoredValue = float | str | bool | None

_BUCKET = 256
_TEXT_TYPES = frozenset({"s", "str", "inlineStr"})
_RANGE_RE = re.compile(
    r"^(?:(?P<sheet>'(?:[^']|'')*'|[^'!]+)!)?"
    r"\$?(?P<c1>[A-Za-z]{1,3})\$?(?P<r1>[0-9]+)"
    r"(?::\$?(?P<c2>[A-Za-z]{1,3})\$?(?P<r2>[0-9]+))?$"
)

# Whole-column / whole-row references (``S!$A:$B``, ``S!1:3``): references, not formulas.
_WHOLE_RE = re.compile(
    r"^(?:(?P<sheet>'(?:[^']|'')*'|[^'!]+)!)?"
    r"(?:\$?(?P<c1>[A-Za-z]{1,3}):\$?(?P<c2>[A-Za-z]{1,3})|\$?(?P<r1>[0-9]+):\$?(?P<r2>[0-9]+))$"
)
_MAX_ROW = 1_048_576
_MAX_COL = 16_384

# ``Name`` or ``Sheet!Name``.
_NAME_REF_RE = re.compile(rf"^(?:(?P<sheet>{SHEET})!)?(?P<name>{NAME})$")

# (sheet, r1, c1, r2, c2) rectangle
_Rect = tuple[str, int, int, int, int]

NameKind = Literal["cells", "formula", "external"]


@dataclass(slots=True, frozen=True)
class NameDef:
    """One definition of a defined name.

    ``scope`` is the sheet that owns a sheet-scoped definition, None for the
    workbook's. ``kind`` is ``cells`` for a reference, ``external`` for cells in
    another workbook, and ``formula`` for a formula, a constant or a broken
    reference. ``areas`` is the IR's destination text, one entry per area;
    ``rects`` the areas this graph can read as cells, a whole column or row cut to
    its sheet's used range (``whole`` is then True). ``relative`` marks a reference
    written without ``$``: Excel reads it relative to the cell that uses the name.
    """

    name: str
    scope: str | None
    kind: NameKind
    areas: tuple[str, ...]
    rects: tuple[_Rect, ...]
    whole: bool = False
    relative: bool = False


def _unquote(sheet: str) -> str:
    if sheet.startswith("'") and sheet.endswith("'") and len(sheet) >= 2:
        return sheet[1:-1].replace("''", "'")
    return sheet


def _parse_rect(text: str, default_sheet: str | None = None) -> _Rect | None:
    """Parse ``Sheet!A1`` or ``Sheet!A1:B2`` (``$`` allowed); None if not a reference."""
    m = _RANGE_RE.match(text.strip())
    if m is None:
        return None
    sheet = _unquote(m.group("sheet")) if m.group("sheet") else default_sheet
    if sheet is None:
        return None
    r1, c1 = int(m.group("r1")), _col_to_number(m.group("c1"))
    r2 = int(m.group("r2")) if m.group("r2") else r1
    c2 = _col_to_number(m.group("c2")) if m.group("c2") else c1
    return (sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def _quoted(sheet: str) -> str:
    """A sheet name always quoted, as a reference may write it."""
    escaped = sheet.replace("'", "''")
    return f"'{escaped}'"


def _split_name(ref: str) -> tuple[str | None, str] | None:
    """``(sheet or None, name)`` for ``Name`` / ``Sheet!Name``; None for anything else."""
    m = _NAME_REF_RE.match(ref.strip())
    if m is None:
        return None
    sheet = m.group("sheet")
    return (_unquote(sheet) if sheet is not None else None), m.group("name")


def _parse_whole(text: str, default_sheet: str | None = None) -> _Rect | None:
    """Parse ``Sheet!B:C`` or ``Sheet!3:4`` to its full rectangle; None otherwise."""
    m = _WHOLE_RE.match(text.strip())
    if m is None:
        return None
    sheet = _unquote(m.group("sheet")) if m.group("sheet") else default_sheet
    if sheet is None:
        return None
    if m.group("c1") is not None:
        a, b = sorted((_col_to_number(m.group("c1")), _col_to_number(m.group("c2"))))
        return (sheet, 1, a, _MAX_ROW, b)
    a, b = sorted((int(m.group("r1")), int(m.group("r2"))))
    return (sheet, a, 1, b, _MAX_COL)


def _convert(raw: str, data_type: str | None) -> StoredValue:
    """Turn the IR's raw stored text into a typed value."""
    if data_type in _TEXT_TYPES:
        return raw
    if data_type == "b":
        return raw.strip().lower() in {"1", "true"}
    if data_type == "e":
        return raw
    try:
        return float(raw)
    except ValueError:
        return raw


def _text(value: StoredValue) -> str:
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


class WorkbookGraph:
    """Formulas, dependency edges and stored values of one extracted workbook."""

    def __init__(self) -> None:
        self.workbook_sha256: str = ""
        self.udf_cells: frozenset[CellRef] = frozenset()
        self.external_cells: dict[CellRef, list[str]] = {}
        self._formula: dict[CellRef, str] = {}
        self._family: dict[CellRef, str] = {}
        self._stored: dict[CellRef, StoredValue] = {}
        self._has_stored: set[CellRef] = set()
        self._cells_by_sheet: dict[str, list[tuple[int, int]]] = {}
        self._direct: dict[CellRef, set[CellRef]] = defaultdict(set)
        self._direct_rev: dict[CellRef, set[CellRef]] = defaultdict(set)
        self._ranges: dict[CellRef, list[_Rect]] = defaultdict(list)
        self._range_grid: dict[tuple[str, int], list[tuple[_Rect, CellRef]]] = defaultdict(list)
        # Keyed (owning sheet or None for the workbook, casefolded name): Excel
        # compares names without case.
        self._names: dict[tuple[str | None, str], NameDef] = {}
        self._sheet_key: dict[str, str] = {}
        self._extent: dict[str, tuple[int, int]] = {}
        self._array_cells: set[CellRef] = set()
        self._validations: list[tuple[list[_Rect], list[str] | None, str | None, str]] = []
        self._sheets: set[str] = set()

    @classmethod
    def from_ir(cls, path: Path) -> WorkbookGraph:
        """Load the graph from an extraction database, opened read-only."""
        g = cls()
        conn = connect_read_only(path)
        try:
            g._load(conn)
        finally:
            conn.close()
        return g

    # -- loading ---------------------------------------------------------

    def _load(self, conn: sqlite3.Connection) -> None:
        row = conn.execute("SELECT value FROM ir_metadata WHERE key='workbook_sha256'").fetchone()
        self.workbook_sha256 = str(row[0]) if row else ""
        sheet_names = dict(conn.execute("SELECT sheet_id, sheet_name FROM sheets"))
        self._sheets = set(sheet_names.values())
        self._sheet_key = {name.casefold(): name for name in self._sheets}

        by_id: dict[int, CellRef] = {}
        per_sheet: dict[str, list[tuple[int, int]]] = defaultdict(list)
        query = (
            "SELECT c.cell_id, c.sheet_id, c.row, c.col, c.formula_a1, c.data_type, "
            "f.formula_r1c1, b.json, c.is_array_formula FROM cells c "
            "LEFT JOIN formulas f ON f.formula_id = c.formula_id "
            "LEFT JOIN json_blobs b ON b.blob_id = c.value_blob_id"
        )
        for cell_id, sheet_id, r, c, fa1, dtype, r1c1, blob, is_array in conn.execute(query):
            ref = CellRef(sheet_names[sheet_id], r, c)
            by_id[cell_id] = ref
            per_sheet[ref.sheet].append((r, c))
            if fa1:
                self._formula[ref] = str(fa1)
                if is_array:
                    self._array_cells.add(ref)
                self._family[ref] = f"{ref.sheet}|{r1c1}"
            if blob is not None and dtype != "blank":
                raw = json.loads(blob)
                if raw is not None:
                    self._stored[ref] = _convert(str(raw), dtype)
                    self._has_stored.add(ref)
        self._cells_by_sheet = {s: sorted(v) for s, v in per_sheet.items()}
        # The used range: the last row and column holding a cell, per sheet.
        self._extent = {
            s: (max(r for r, _c in v), max(c for _r, c in v)) for s, v in per_sheet.items()
        }

        for frm, to in conn.execute("SELECT from_cell_id, to_cell_id FROM cell_edges_internal"):
            if frm in by_id and to in by_id:
                self._direct[by_id[frm]].add(by_id[to])
                self._direct_rev[by_id[to]].add(by_id[frm])

        # Whole columns/rows are stored to row 1048576 / column 16384; index only the
        # extent a sheet uses, or one entry per bucket would fill memory.
        max_row = {s: extent[0] for s, extent in self._extent.items()}
        range_q = (
            "SELECT from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2 FROM range_edges "
            "ORDER BY from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2"
        )
        for frm, sid, r1, c1, r2, c2 in conn.execute(range_q):
            if frm not in by_id or sid not in sheet_names:
                continue
            rect: _Rect = (sheet_names[sid], min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))
            src = by_id[frm]
            self._ranges[src].append(rect)
            last = min(rect[3], max_row.get(rect[0], 0))
            for bucket in range(rect[1] // _BUCKET, last // _BUCKET + 1):
                self._range_grid[(rect[0], bucket)].append((rect, src))

        ext: dict[CellRef, list[str]] = defaultdict(list)
        for frm, ref_text in conn.execute(
            "SELECT from_cell_id, external_ref FROM cell_edges_external ORDER BY 1, 2"
        ):
            if frm in by_id:
                ext[by_id[frm]].append(str(ref_text))
        self.external_cells = dict(ext)
        self.udf_cells = frozenset(
            by_id[cid]
            for (cid,) in conn.execute("SELECT cell_id FROM cell_udf_calls")
            if cid in by_id
        )

        for name, scope, dests, ext in conn.execute(
            "SELECT name, scope, destinations, is_external FROM defined_names ORDER BY 1, 2"
        ):
            raw = tuple(str(d) for d in json.loads(dests or "[]"))
            owner = None if scope == "workbook" else self.sheet(str(scope)) or str(scope)
            self._names.setdefault(
                (owner, str(name).casefold()), self._name_def(str(name), owner, raw, bool(ext))
            )

        for sid, sqref, vtype, f1, lit in conn.execute(
            "SELECT sheet_id, sqref, val_type, formula1, literal_values_json FROM data_validations"
        ):
            if vtype != "list" or sid not in sheet_names:
                continue
            sheet = sheet_names[sid]
            rects = [
                r for part in str(sqref).split() if (r := _parse_rect(part, sheet)) is not None
            ]
            literals = [str(v) for v in json.loads(lit)] if lit else None
            self._validations.append((rects, literals, f1, sheet))

    def _name_def(
        self, name: str, owner: str | None, raw: tuple[str, ...], external: bool
    ) -> NameDef:
        if external or any(d.lstrip("'").startswith("[") for d in raw):
            # Cells in another workbook: not cells of this graph (``[1]Src`` would
            # otherwise parse as a sheet of its own).
            return NameDef(name, owner, "external", raw, ())
        # The IR keeps only reference destinations: a formula name has none, a
        # constant name keeps its literal text. Either way it is not cells.
        if not raw or any(_parse_rect(d) is None and not _WHOLE_RE.match(d) for d in raw):
            return NameDef(name, owner, "formula", raw, ())
        rects: list[_Rect] = []
        whole = False
        for d in raw:
            if (r := _parse_rect(d)) is None and (r := _parse_whole(d)) is not None:
                whole = True
                r = self._clip(r)
            if r is not None:
                sheet, r1, c1, r2, c2 = r
                rects.append((self.sheet(sheet) or sheet, r1, c1, r2, c2))
        return NameDef(name, owner, "cells", raw, tuple(rects), whole)

    def _clip(self, rect: _Rect) -> _Rect | None:
        """A whole column or row cut to its sheet's used range; None when outside it."""
        sheet, r1, c1, r2, c2 = rect
        last_row, last_col = self._extent.get(self.sheet(sheet) or sheet, (0, 0))
        r2, c2 = min(r2, last_row), min(c2, last_col)
        if r1 > r2 or c1 > c2:
            return None
        return (sheet, r1, c1, r2, c2)

    def mark_relative(self, names: Iterable[tuple[str | None, str]]) -> None:
        """Flag the definitions whose reference is relative, given as (owning sheet or
        None, name); the IR stores destinations without ``$``, so the caller reads
        them from the workbook."""
        for owner, name in names:
            sheet = (self.sheet(owner) or owner) if owner is not None else None
            key = (sheet, name.casefold())
            found = self._names.get(key)
            if found is not None:
                self._names[key] = replace(found, relative=True)

    # -- basic lookups ---------------------------------------------------

    def sheets(self) -> list[str]:
        """Every sheet name, sorted."""
        return sorted(self._sheets)

    def sheet(self, text: str) -> str | None:
        """The workbook's spelling of a sheet name given in any case; None if unknown."""
        return self._sheet_key.get(text.casefold())

    def formula(self, cell: CellRef) -> str | None:
        """A1 formula text without ``=``, or None for a constant or unknown cell."""
        return self._formula.get(cell)

    def formula_cells(self) -> list[CellRef]:
        """Every formula cell, sorted."""
        return sorted(self._formula)

    def stored(self, cell: CellRef) -> StoredValue:
        """The cached value Excel stored (errors as ``#CODE`` text); None when absent."""
        return self._stored.get(cell)

    def has_stored(self, cell: CellRef) -> bool:
        """True when the IR holds a stored value for the cell."""
        return cell in self._has_stored

    def family(self, cell: CellRef) -> str:
        """Sheet plus relative (R1C1) formula: copied formulas share a family."""
        return self._family.get(cell, f"{cell.sheet}|")

    def lookup_name(self, ref: str, sheet: str | None) -> NameDef | None:
        """The definition ``Name`` or ``Sheet!Name`` reads in a formula on ``sheet``.

        As Excel resolves it: the sheet's own (sheet-scoped) definition first, then the
        workbook's; ``Sheet!Name`` looks on Sheet. Names and sheets ignore case. With
        ``sheet`` None an unqualified name reads the workbook's definition only. None
        when ``ref`` is not a defined name.
        """
        parts = _split_name(ref)
        if parts is None:
            return None
        qualifier, name = parts
        if qualifier is not None:
            sheet = self.sheet(qualifier)
            if sheet is None:
                return None
        key = name.casefold()
        if sheet is not None and (own := self._names.get((sheet, key))) is not None:
            return own
        return self._names.get((None, key))

    def is_array_formula(self, cell: CellRef) -> bool:
        """True when the cell holds an array (CSE or dynamic-array) formula."""
        return cell in self._array_cells

    def name_cells(self, ref: str, sheet: str | None = None) -> list[CellRef]:
        """Cells a defined name points at, as read from ``sheet`` (sorted; empty cells
        included); empty when it is not a name over cells."""
        found = self.lookup_name(ref, sheet)
        if found is None or found.kind != "cells":
            return []
        return self._full_rects(found.rects)

    # -- graph traversal -------------------------------------------------

    def _cells_in_rect(self, rect: _Rect) -> list[CellRef]:
        sheet, r1, c1, r2, c2 = rect
        cells = self._cells_by_sheet.get(sheet, [])
        out: list[CellRef] = []
        for i in range(bisect_left(cells, (r1, 0)), len(cells)):
            r, c = cells[i]
            if r > r2:
                break
            if c1 <= c <= c2:
                out.append(CellRef(sheet, r, c))
        return out

    @staticmethod
    def _full_rects(rects: Iterable[_Rect]) -> list[CellRef]:
        """Every position in each rectangle, row-major, empty cells included."""
        seen: set[CellRef] = set()
        out: list[CellRef] = []
        for sheet, r1, c1, r2, c2 in rects:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    ref = CellRef(sheet, r, c)
                    if ref not in seen:
                        seen.add(ref)
                        out.append(ref)
        return sorted(out)

    def precedents(self, cell: CellRef) -> frozenset[CellRef]:
        """Direct cell edges plus every IR cell inside each range the formula reads."""
        out = set(self._direct.get(cell, ()))
        for rect in self._ranges.get(cell, ()):
            out.update(self._cells_in_rect(rect))
        return frozenset(out)

    def _dependents(self, cell: CellRef) -> set[CellRef]:
        out = set(self._direct_rev.get(cell, ()))
        for rect, src in self._range_grid.get((cell.sheet, cell.row // _BUCKET), ()):
            if rect[1] <= cell.row <= rect[3] and rect[2] <= cell.col <= rect[4]:
                out.add(src)
        return out

    def dependents(self, cell: CellRef) -> frozenset[CellRef]:
        """Formula cells that read ``cell`` directly or through a range."""
        return frozenset(self._dependents(cell))

    def downstream(self, roots: Iterable[CellRef]) -> dict[CellRef, CellRef]:
        """Every cell reachable forward from the roots, mapped to the smallest root reaching it."""
        owner: dict[CellRef, CellRef] = {}
        for root in sorted(set(roots)):
            if root in owner:
                continue
            owner[root] = root
            queue = deque([root])
            while queue:
                cur = queue.popleft()
                for nxt in self._dependents(cur):
                    if nxt not in owner:
                        owner[nxt] = root
                        queue.append(nxt)
        return owner

    def upstream(self, cells: Iterable[CellRef]) -> set[CellRef]:
        """Backward closure through ``precedents`` (the cells themselves included)."""
        seen: set[CellRef] = set()
        stack = list(cells)
        while stack:
            cur = stack.pop()
            if cur in seen:
                continue
            seen.add(cur)
            stack.extend(self.precedents(cur) - seen)
        return seen

    # -- targets and validations ----------------------------------------

    def resolve_target(self, target: str) -> list[CellRef]:
        """Resolve an A1 cell or range, or a defined name, to cells.

        Sheets and names ignore case. A bare name is the workbook's definition;
        ``'Sheet'!Name`` is Sheet's own (sheet-scoped) definition. A bare name that a
        sheet defines too is ambiguous outside any sheet, so it is refused.

        Raises:
            CalcInputError: if the sheet or name is unknown, or the name is ambiguous.
        """
        text = target.strip()
        named = self._name_target(text, target)
        if named is not None:
            return named
        rect = _parse_rect(text)
        whole = rect is None and (rect := _parse_whole(text)) is not None
        if rect is None:
            raise CalcInputError(f"unknown target {target!r}: not a cell, range or defined name")
        written, r1, c1, r2, c2 = rect
        sheet = self.sheet(written)
        if sheet is None:
            raise CalcInputError(f"unknown sheet {written!r} in target {target!r}")
        if whole:
            # A whole column or row is the part of it the sheet uses (ruling D10).
            clipped = self._clip((sheet, r1, c1, r2, c2))
            if clipped is None:
                raise CalcInputError(
                    f"target {target!r} holds no cells: it is outside the range sheet "
                    f"{sheet!r} uses"
                )
            return self._full_rects([clipped])
        if (r1, c1) == (r2, c2):
            return [CellRef(sheet, r1, c1)]
        return self._full_rects([(sheet, r1, c1, r2, c2)])

    def _name_target(self, text: str, target: str) -> list[CellRef] | None:
        """The cells of a name target; None when ``text`` is not a defined name."""
        parts = _split_name(text)
        if parts is None:
            return None
        qualifier, name = parts
        key = name.casefold()
        owners = sorted(s for (s, k) in self._names if s is not None and k == key)
        book = self._names.get((None, key))
        if qualifier is not None:
            sheet = self.sheet(qualifier)
            if sheet is None:
                return None
            own = self._names.get((sheet, key))
            if own is not None:
                return self._name_cells_target(own, target)
            if book is None and not owners:
                return None
            where = f"the workbook's {book.name!r} is written without a sheet" if book else ""
            if owners:
                where = "; ".join(filter(None, [where, self._sheet_names_hint(owners, name)]))
            raise CalcInputError(
                f"target {target!r}: sheet {sheet!r} has no name {name!r} of its own ({where})"
            )
        if owners:
            hint = self._sheet_names_hint(owners, name)
            if book is None:
                raise CalcInputError(
                    f"target {target!r}: {name!r} is defined only for a sheet; {hint}"
                )
            address = ", ".join(book.areas) or "its address"
            raise CalcInputError(
                f"target {target!r} is ambiguous: the workbook and a sheet each define "
                f"{name!r}; {hint}, or {address} for the workbook's"
            )
        if book is None:
            return None
        return self._name_cells_target(book, target)

    def _name_cells_target(self, found: NameDef, target: str) -> list[CellRef]:
        """The cells a name target covers.

        Raises:
            CalcInputError: when the name holds no cells: a formula, a constant, a
                broken reference, another workbook's cells, or a whole column or row
                outside the range its sheet uses.
        """
        name = found.name
        if found.relative:
            raise CalcInputError(
                f"target {target!r}: name {name!r} has a relative reference (it names "
                "different cells from each cell that uses it); give the cells instead"
            )
        if found.kind == "external":
            raise CalcInputError(
                f"target {target!r}: name {name!r} points into another workbook; "
                "give the cells it is copied into instead"
            )
        if found.kind == "formula":
            if any("#REF!" in a for a in found.areas):
                raise CalcInputError(f"target {target!r}: name {name!r} is broken (#REF!)")
            raise CalcInputError(
                f"target {target!r}: name {name!r} holds a formula or a constant, not cells; "
                "give the cells as a range instead (for example Calc!C1:C40)"
            )
        cells = self._full_rects(found.rects)
        if not cells:
            raise CalcInputError(
                f"target {target!r}: name {name!r} holds no cells (a whole column or row "
                "outside the range its sheet uses)"
            )
        return cells

    def _sheet_names_hint(self, owners: list[str], name: str) -> str:
        spelled = [self._names[(s, name.casefold())].name for s in owners]
        options = " or ".join(f"{_quoted(s)}!{n}" for s, n in zip(owners, spelled, strict=True))
        return f"write {options}"

    def allowed_values(self, cell: CellRef) -> list[str] | None:
        """List-validation values for the cell, or None when it has no list validation
        that can be read.

        The list is the literal items, or the stored values of the source: a range, a
        whole column or row (the cells it holds), or a defined name over cells, read
        from the validation's sheet as a formula there would read it. A source that is
        a formula (``INDIRECT``, ``OFFSET``, a name holding one) is not read: its list
        can change with the inputs (ruling D8 leaves it unchecked in 0.8.1).
        """
        for rects, literals, f1, sheet in self._validations:
            if not any(
                r[0] == cell.sheet and r[1] <= cell.row <= r[3] and r[2] <= cell.col <= r[4]
                for r in rects
            ):
                continue
            if literals is not None:
                return literals
            source = self._list_source(str(f1 or ""), sheet)
            if source is None:
                continue
            cells = [c for rect in source for c in self._cells_in_rect(rect)]
            return [_text(self._stored[c]) for c in cells if c in self._stored]
        return None

    def _list_source(self, formula1: str, sheet: str) -> list[_Rect] | None:
        """The rectangles a list validation's source covers; None for a formula."""
        text = formula1.strip().removeprefix("=").strip()
        if (rect := _parse_rect(text, sheet)) is not None:
            written, r1, c1, r2, c2 = rect
            return [(self.sheet(written) or written, r1, c1, r2, c2)]
        if (whole := _parse_whole(text, sheet)) is not None:
            written, r1, c1, r2, c2 = whole
            clipped = self._clip((self.sheet(written) or written, r1, c1, r2, c2))
            return [clipped] if clipped is not None else []
        found = self.lookup_name(text, sheet)
        if found is None or found.kind != "cells" or found.relative:
            return None
        return list(found.rects)

    def list_source_name(self, cell: CellRef) -> NameDef | None:
        """The defined name a cell's list validation takes its items from, if any."""
        for rects, literals, f1, sheet in self._validations:
            if literals is not None or not any(
                r[0] == cell.sheet and r[1] <= cell.row <= r[3] and r[2] <= cell.col <= r[4]
                for r in rects
            ):
                continue
            text = str(f1 or "").strip().removeprefix("=").strip()
            found = self.lookup_name(text, sheet)
            if found is not None:
                return found
        return None


def open_graph(workbook: Path, ir: Path | None, tmp: Path) -> WorkbookGraph:
    """Load the graph for ``workbook``, reusing ``ir`` when given.

    With ``ir`` the database must have been extracted from this exact workbook
    (SHA-256 match). Without it, the workbook is extracted into ``tmp``; the caller
    owns ``tmp`` and deletes it. The extraction database is never written otherwise.

    Raises:
        CalcInputError: if ``ir`` was extracted from a different workbook.
    """
    if ir is not None:
        graph = WorkbookGraph.from_ir(ir)
        try:
            digest = hashlib.sha256(workbook.read_bytes()).hexdigest()
        except OSError as exc:
            raise CalcInputError(f"cannot read {workbook}: {exc}") from exc
        if digest != graph.workbook_sha256:
            raise CalcInputError(f"IR {ir} does not match workbook {workbook} (sha256 differs)")
    else:
        from xl_marinade.core.api import extract

        tmp.mkdir(parents=True, exist_ok=True)
        graph = WorkbookGraph.from_ir(extract(workbook, tmp / "ir.db"))
    graph.mark_relative(relative_names(workbook))
    return graph
