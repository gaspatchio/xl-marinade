"""`cells.is_array_formula` is set for array formulas.

Issue: the extractor wrote 0 for every cell, so an array (CSE) formula such as
``{=G1:G3*2}`` over H1:H3 -- ``<f t="array" ref="H1:H3">`` in the sheet XML --
was indistinguishable from an ordinary one in `agent_cells`.
"""

import sqlite3
from pathlib import Path

import openpyxl
import pytest
from openpyxl.worksheet.formula import ArrayFormula

from xl_marinade.core.api import extract
from xl_marinade.core.new_arch.fast_parser import stream_worksheet_cells


@pytest.fixture(scope="module")
def workbook(tmp_path_factory: pytest.TempPathFactory) -> Path:
    tmp = tmp_path_factory.mktemp("array_flag")
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Main"
    for r in (1, 2, 3):
        ws.cell(row=r, column=7, value=r)  # G1:G3
    ws["H1"] = ArrayFormula("H1:H3", "=G1:G3*2")  # multi-cell array formula
    ws["J1"] = ArrayFormula("J1", "=SUM(G1:G3*G1:G3)")  # single-cell CSE
    ws["K1"] = "=G1*2"  # ordinary formula
    xlsx = tmp / "t.xlsx"
    wb.save(xlsx)
    return xlsx


@pytest.fixture(scope="module")
def flags(workbook: Path, tmp_path_factory: pytest.TempPathFactory) -> dict[str, int]:
    db = extract(workbook, tmp_path_factory.mktemp("array_flag_db") / "t.db")
    conn = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
    try:
        rows = conn.execute("SELECT cell_address, is_array_formula FROM agent_cells").fetchall()
    finally:
        conn.close()
    return {addr: int(flag) for addr, flag in rows}


@pytest.mark.parametrize("cell", ["Main!H1", "Main!J1"])
def test_array_formula_cell_is_flagged(flags: dict[str, int], cell: str) -> None:
    assert flags[cell] == 1


@pytest.mark.parametrize("cell", ["Main!K1", "Main!G1"])
def test_other_cells_are_not_flagged(flags: dict[str, int], cell: str) -> None:
    assert flags[cell] == 0


def test_parser_reports_array_cells_only_when_asked(workbook: Path) -> None:
    found: set[tuple[int, int]] = set()
    with_set = list(stream_worksheet_cells(str(workbook), "Main", found))
    assert found == {(1, 8), (1, 10)}
    # The yielded tuples are unchanged, and the argument stays optional.
    assert list(stream_worksheet_cells(str(workbook), "Main")) == with_set
