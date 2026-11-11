"""`agent_dependencies` lists a formula's reads of cells that hold nothing.

Issue: the view joined the target of every cell edge to `cells`, and a target
with no row there (an empty input, a formatted-but-empty cell, a cell on an
empty sheet) dropped the edge, although `cell_edges_internal` held it. An
agent tracing what drives a cell then missed the read entirely.
"""

import sqlite3
from pathlib import Path

import openpyxl
import pytest
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from xl_marinade.core.api import extract
from xl_marinade.core.new_arch.cell_identity import pack


def _deps(db: Path) -> set[tuple[str, str, str]]:
    conn = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
    try:
        rows = conn.execute(
            "SELECT from_cell, to_cell, dependency_type FROM agent_dependencies"
        ).fetchall()
    finally:
        conn.close()
    return {(f, t, k) for f, t, k in rows}


@pytest.fixture(scope="module")
def blank_targets_db(tmp_path_factory: pytest.TempPathFactory) -> Path:
    tmp = tmp_path_factory.mktemp("blank_targets")
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Main"
    ws["A1"] = "=B20"  # never-touched cell
    ws["A2"] = "=C2"  # never-touched cell, same row
    ws["A3"] = "=D3"  # formatted but empty
    ws["D3"].fill = PatternFill("solid", fgColor="FFFF00")
    ws["A4"] = "=Empty!A1"  # cell on a sheet with no cells at all
    ws["A5"] = "=SUM(B50:B52)"  # range over absent cells (already worked)
    ws["A6"] = "=E6"  # a target that does hold a value
    ws["E6"] = 7
    wb.create_sheet("Empty")
    xlsx = tmp / "t.xlsx"
    wb.save(xlsx)
    return extract(xlsx, tmp / "t.db")


@pytest.mark.parametrize(
    ("from_cell", "to_cell"),
    [
        ("Main!A1", "Main!B20"),
        ("Main!A2", "Main!C2"),
        ("Main!A3", "Main!D3"),
        ("Main!A4", "Empty!A1"),
        ("Main!A6", "Main!E6"),
    ],
)
def test_cell_edge_listed_whether_or_not_target_is_stored(
    blank_targets_db: Path, from_cell: str, to_cell: str
) -> None:
    assert (from_cell, to_cell, "cell") in _deps(blank_targets_db)


def test_range_edge_over_absent_cells_still_listed(blank_targets_db: Path) -> None:
    assert ("Main!A5", "Main!B50:B52", "range") in _deps(blank_targets_db)


def test_view_lists_every_internal_edge(blank_targets_db: Path) -> None:
    conn = sqlite3.connect(f"file:{blank_targets_db}?mode=ro", uri=True)
    try:
        (raw,) = conn.execute("SELECT COUNT(*) FROM cell_edges_internal").fetchone()
        (viewed,) = conn.execute(
            "SELECT COUNT(*) FROM agent_dependencies WHERE dependency_type = 'cell'"
        ).fetchone()
    finally:
        conn.close()
    assert viewed == raw


def test_decoded_address_matches_excel_for_every_column(
    blank_targets_db: Path, tmp_path: Path
) -> None:
    """Edges to absent cells, across A..XFD and both row extremes, decode right."""
    db = tmp_path / "copy.db"
    db.write_bytes(blank_targets_db.read_bytes())
    conn = sqlite3.connect(db)
    try:
        (from_id,) = conn.execute(
            "SELECT cell_id FROM cells WHERE a1 = 'A1' AND sheet_id = 1"
        ).fetchone()
        conn.execute("DELETE FROM cell_edges_internal")
        conn.execute("DELETE FROM range_edges")
        conn.execute("DELETE FROM cell_edges_external")
        targets = [(2, row, col) for col in range(1, 16385) for row in (1, 1_048_576)]
        conn.executemany(
            "INSERT INTO cell_edges_internal (from_cell_id, to_cell_id) VALUES (?, ?)",
            [(from_id, pack(s, r, c)) for s, r, c in targets],
        )
        got = {t for (t,) in conn.execute("SELECT to_cell FROM agent_dependencies")}
    finally:
        conn.close()
    want = {f"Empty!{get_column_letter(c)}{r}" for _, r, c in targets}
    assert got == want
