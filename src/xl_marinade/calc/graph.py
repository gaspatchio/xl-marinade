"""The formula graph and stored values, loaded from the extraction IR."""

from __future__ import annotations

import hashlib
import json
import re
import sqlite3
from bisect import bisect_left
from collections import defaultdict, deque
from collections.abc import Iterable
from pathlib import Path

from xl_marinade.calc.models import CellRef, _col_to_number
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
    r"^(?:(?:'(?:[^']|'')*'|[^'!]+)!)?"
    r"(?:\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[0-9]+:\$?[0-9]+)$"
)

# (sheet, r1, c1, r2, c2) rectangle
_Rect = tuple[str, int, int, int, int]


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
        self._names: dict[str, tuple[str, list[_Rect]]] = {}
        self._formula_names: set[str] = set()
        self._external_names: set[str] = set()
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

        for frm, to in conn.execute("SELECT from_cell_id, to_cell_id FROM cell_edges_internal"):
            if frm in by_id and to in by_id:
                self._direct[by_id[frm]].add(by_id[to])
                self._direct_rev[by_id[to]].add(by_id[frm])

        # Whole columns/rows are stored to row 1048576 / column 16384; index only the
        # extent a sheet uses, or one entry per bucket would fill memory.
        max_row = {s: max(r for r, _c in v) for s, v in per_sheet.items()}
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
            raw = [str(d) for d in json.loads(dests or "[]")]
            if ext or any(d.lstrip("'").startswith("[") for d in raw):
                # Cells in another workbook: not cells of this graph (``[1]Src`` would
                # otherwise parse as a sheet of its own).
                self._external_names.add(str(name))
                self._formula_names.add(str(name))
                if name not in self._names or scope == "workbook":
                    self._names[name] = (str(scope), [])
                continue
            rects = [r for d in raw if (r := _parse_rect(d)) is not None]
            # The IR keeps only reference destinations: a formula name has none, a
            # constant name keeps its literal text. Either way it is not cells.
            if not raw or any(_parse_rect(d) is None and not _WHOLE_RE.match(d) for d in raw):
                self._formula_names.add(str(name))
            if name not in self._names or scope == "workbook":
                self._names[name] = (str(scope), rects)

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

    # -- basic lookups ---------------------------------------------------

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

    def is_sheet_scoped_name(self, name: str) -> bool:
        """True when the only definition of ``name`` is sheet-scoped."""
        entry = self._names.get(name)
        return entry is not None and entry[0] != "workbook"

    def is_formula_name(self, name: str) -> bool:
        """True when a definition of ``name`` holds a formula or constant, not cells."""
        return name in self._formula_names

    def is_external_name(self, name: str) -> bool:
        """True when a definition of ``name`` points into another workbook."""
        return name in self._external_names

    def is_array_formula(self, cell: CellRef) -> bool:
        """True when the cell holds an array (CSE or dynamic-array) formula."""
        return cell in self._array_cells

    def name_cells(self, name: str) -> list[CellRef]:
        """Cells present in the IR that a defined name points at (sorted)."""
        entry = self._names.get(name)
        if entry is None:
            return []
        return self._full_rects(entry[1])

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
        """Resolve an A1 cell/range or a workbook-scoped defined name to cells.

        Raises:
            CalcInputError: if the sheet or name is unknown.
        """
        text = target.strip()
        entry = self._names.get(text)
        if entry is not None and entry[0] == "workbook":
            return self._full_rects(entry[1])
        rect = _parse_rect(text)
        if rect is None:
            raise CalcInputError(f"unknown target {target!r}: not a cell, range or defined name")
        if rect[0] not in self._sheets:
            raise CalcInputError(f"unknown sheet {rect[0]!r} in target {target!r}")
        if rect[1:3] == rect[3:5]:
            return [CellRef(rect[0], rect[1], rect[2])]
        return self._full_rects([rect])

    def allowed_values(self, cell: CellRef) -> list[str] | None:
        """List-validation values for the cell, or None when it has no list validation."""
        for rects, literals, f1, sheet in self._validations:
            if not any(
                r[0] == cell.sheet and r[1] <= cell.row <= r[3] and r[2] <= cell.col <= r[4]
                for r in rects
            ):
                continue
            if literals is not None:
                return literals
            source = _parse_rect(str(f1 or ""), sheet)
            if source is None:
                continue
            return [
                _text(self._stored[c]) for c in self._cells_in_rect(source) if c in self._stored
            ]
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
        return graph
    from xl_marinade.core.api import extract

    tmp.mkdir(parents=True, exist_ok=True)
    return WorkbookGraph.from_ir(extract(workbook, tmp / "ir.db"))
