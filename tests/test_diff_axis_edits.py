"""Interior row/column insertions and deletions diff to what was edited (issue #53).

Deleting interior columns from a block of formula columns put the deletion at
the sheet's tail: every column to its right has relative references that now
span fewer columns (RC[-5] becomes RC[-3]), so no column matched itself, and
the identity fallback paired each with whatever now sat at its old address.
The binding map then reported renames that never happened and missed the
deleted columns entirely.

Each test builds both versions from a column spec and writes the formulas
Excel itself would hold after the edit, so the only difference between the
two workbooks is the edit under test.
"""

import json
from collections.abc import Callable, Sequence
from pathlib import Path

import openpyxl
import pytest
from openpyxl.utils import get_column_letter as col_letter
from openpyxl.workbook.defined_name import DefinedName

import xl_marinade
from xl_marinade.core.ir_diff.axis_detect import _loosen_formula
from xl_marinade.core.ir_diff.formula_rewrite import (
    AxisIndex,
    _tokenize_r1c1,
    rewrite_formula_r1c1,
)

ROWS = 12
VARS = [f"v{i}" for i in range(1, 11)]


def _projection(
    path: Path,
    names: Sequence[str],
    formula: Callable[[str, int, int], str] | None = None,
) -> Path:
    """Sheet `Proj`: `t` in column A, then one labelled formula column per name.

    Each column carries a defined name over its data rows. By default a
    column reads column A relatively (`=A2*k+k`), as a projection would.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Proj"
    ws["A1"] = "t"
    for r in range(2, ROWS + 2):
        ws[f"A{r}"] = r - 2
    for j, name in enumerate(names):
        col = col_letter(j + 2)
        k = int(name[1:])
        ws[f"{col}1"] = name
        for r in range(2, ROWS + 2):
            ws[f"{col}{r}"] = formula(name, j, r) if formula else f"=A{r}*{k}+{k}"
        wb.defined_names[f"Proj_{name}"] = DefinedName(
            f"Proj_{name}", attr_text=f"Proj!${col}$2:${col}${ROWS + 1}"
        )
    wb.save(path)
    return path


def _diff(tmp_path: Path, build_a: Callable[[Path], Path], build_b: Callable[[Path], Path]) -> dict:
    a = xl_marinade.extract(build_a(tmp_path / "a.xlsx"), tmp_path / "a.db")
    b = xl_marinade.extract(build_b(tmp_path / "b.xlsx"), tmp_path / "b.db")
    return xl_marinade.diff(a, b)


def _of(result: dict, *types: str) -> list[dict]:
    return [c for c in result["changes"] if c["type"] in types]


def _workbook_types(result: dict) -> set[str]:
    return {
        c["type"]
        for c in result["changes"]
        if c["layer"] == "workbook" and c["type"] != "IR_METADATA_CHANGED"
    }


# ---------------------------------------------------------------------------
# End to end
# ---------------------------------------------------------------------------


def test_interior_column_deletion(tmp_path: Path) -> None:
    """The #53 reproduction: delete v3 and v4 (columns D:E) from v1..v10."""
    result = _diff(
        tmp_path,
        lambda p: _projection(p, VARS),
        lambda p: _projection(p, [v for v in VARS if v not in ("v3", "v4")]),
    )

    deleted = _of(result, "COLS_DELETED")
    assert [(c["sheet"], c["at_col"], c["count"]) for c in deleted] == [("Proj", 4, 2)]
    assert result["summary"]["cols_deleted"] == 2  # columns, not events

    removed = _of(result, "BINDING_REMOVED")
    assert sorted((c["label"], c["address"]) for c in removed) == [
        ("v3", "D2:D13"),
        ("v4", "E2:E13"),
    ]
    assert all(c["with_axis_deletion"] for c in removed)

    # Nothing to the right of the gap was edited: no renames, no formula
    # edits, and the defined names Excel shifted read as unchanged.
    assert _workbook_types(result) == {
        "COLS_DELETED",
        "BINDING_REMOVED",
        "NAME_REMOVED",
        "BINDING_RESIZED",  # the header row A1:K1 is now A1:I1
    }
    summary = result["summary"]
    assert summary["bindings_removed"] == 2
    assert summary["bindings_label_changed"] == 0
    assert summary["bindings_formula_changed"] == 0
    assert summary["cells_formula_changed"] == 0

    by_label = {m["label_a"]: m for m in result["binding_map"] if m["label_a"]}
    assert by_label["v3"]["diff_state"] == "removed"
    assert by_label["v4"]["diff_state"] == "removed"
    assert not by_label["v3"]["shifted"]
    # Bindings left of the gap kept their address. (The extractor groups v1
    # with the t column, as A2:B13 labelled "t".)
    for label in ("t", "v2"):
        assert (by_label[label]["diff_state"], by_label[label]["shifted"]) == (
            "unchanged",
            False,
        )
    # Columns right of the gap are unchanged, but flagged as shifted.
    for k in range(5, 11):
        entry = by_label[f"v{k}"]
        assert entry["diff_state"] == "unchanged"
        assert entry["shifted"] is True
        assert entry["label_b"] == f"v{k}"
        assert entry["address_a"] == f"{col_letter(k + 1)}2:{col_letter(k + 1)}13"
        assert entry["address_b"] == f"{col_letter(k - 1)}2:{col_letter(k - 1)}13"
    # The header row only shrank: same top-left, so not shifted.
    header = next(m for m in result["binding_map"] if m["diff_state"] == "modified")
    assert header["address_a"].endswith("A1:K1")
    assert header["shifted"] is False
    # Every row carries the flag.
    assert all("shifted" in m for m in result["binding_map"])


def test_column_that_read_a_deleted_column_still_aligns_by_its_label(tmp_path: Path) -> None:
    """A column whose formula became #REF! keeps its header, and is matched by it.

    `v7` reads `v4`. Deleting `v4` turns v7's formula into `#REF!`, so no
    formula-based signature matches it; its header label still does.
    """

    def formula(deleted: bool) -> Callable[[str, int, int], str]:
        def f(name: str, j: int, r: int) -> str:
            k = int(name[1:])
            if name == "v7":
                return "=#REF!*2" if deleted else f"=E{r}*2"
            return f"=A{r}*{k}+{k}"

        return f

    result = _diff(
        tmp_path,
        lambda p: _projection(p, VARS, formula(False)),
        lambda p: _projection(p, [v for v in VARS if v != "v4"], formula(True)),
    )

    assert [(c["at_col"], c["count"]) for c in _of(result, "COLS_DELETED")] == [(5, 1)]
    assert [c["label"] for c in _of(result, "BINDING_REMOVED")] == ["v4"]
    assert result["summary"]["bindings_label_changed"] == 0
    # The one real consequence: v7's formula lost its reference.
    changed = _of(result, "BINDING_FORMULA_CHANGED")
    assert [c["address"] for c in changed] == ["G2:G13"]
    assert "#REF!" in changed[0]["new_formula"]


def test_interior_column_insertion(tmp_path: Path) -> None:
    result = _diff(
        tmp_path,
        lambda p: _projection(p, VARS),
        lambda p: _projection(p, [*VARS[:4], "v11", *VARS[4:]]),
    )
    assert [(c["at_col"], c["count"]) for c in _of(result, "COLS_INSERTED")] == [(6, 1)]
    assert [c["label"] for c in _of(result, "BINDING_ADDED")] == ["v11"]
    assert result["summary"]["bindings_label_changed"] == 0
    assert result["summary"]["cells_formula_changed"] == 0


def test_binding_that_lost_its_first_column_is_matched_not_dropped(tmp_path: Path) -> None:
    """A block C:E with one shared formula; deleting C leaves it two columns wide.

    Canonicalization used to drop any binding whose top-left column was
    deleted, so the surviving part came back as an unrelated added binding.
    Column B is left blank so the extractor keeps the block apart from `t`.
    """

    def build(path: Path, width: int) -> Path:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Proj"
        ws["A1"] = "t"
        # Distinct headers: identical columns would make "C deleted" and "E
        # deleted" the same edit.
        for c, header in enumerate(["x1", "x2", "x3"][3 - width :]):
            ws[f"{col_letter(c + 3)}1"] = header
        for r in range(2, ROWS + 2):
            ws[f"A{r}"] = r - 2
            for c in range(width):
                ws[f"{col_letter(c + 3)}{r}"] = f"=$A{r}*2"
        wb.save(path)
        return path

    result = _diff(tmp_path, lambda p: build(p, 3), lambda p: build(p, 2))
    assert [(c["at_col"], c["count"]) for c in _of(result, "COLS_DELETED")] == [(3, 1)]
    assert _of(result, "BINDING_REMOVED") == []
    assert _of(result, "BINDING_ADDED") == []
    resized = {c["address"]: c["new_shape"] for c in _of(result, "BINDING_RESIZED")}
    assert resized.get("C2:D13") == (ROWS, 2)
    assert result["summary"]["cells_formula_changed"] == 0


def test_interior_row_deletion_with_references_across_it(tmp_path: Path) -> None:
    """Rows are aligned the same way: a total row reading rows across a deletion."""

    def build(path: Path, values: Sequence[int]) -> Path:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        for i, v in enumerate(values, start=1):
            ws[f"A{i}"], ws[f"B{i}"] = f"item{v}", v
        total = len(values) + 2
        ws[f"A{total}"] = "total"
        ws[f"B{total}"] = f"=B1+B{len(values)}"  # relative refs across the whole block
        wb.save(path)
        return path

    result = _diff(
        tmp_path,
        lambda p: build(p, range(1, 9)),
        lambda p: build(p, [v for v in range(1, 9) if v != 4]),
    )
    assert [(c["at_row"], c["count"]) for c in _of(result, "ROWS_DELETED")] == [(4, 1)]
    assert _of(result, "FORMULA_CHANGED", "BINDING_FORMULA_CHANGED") == []


def test_axis_runs_are_contiguous_and_summary_counts_columns(tmp_path: Path) -> None:
    """Deleting non-adjacent columns gives one event per contiguous block."""
    result = _diff(
        tmp_path,
        lambda p: _projection(p, VARS),
        lambda p: _projection(p, [v for v in VARS if v not in ("v2", "v5", "v6")]),
    )
    assert [(c["at_col"], c["count"]) for c in _of(result, "COLS_DELETED")] == [(3, 1), (6, 2)]
    assert result["summary"]["cols_deleted"] == 3


def test_self_diff_is_empty(tmp_path: Path) -> None:
    result = _diff(tmp_path, lambda p: _projection(p, VARS), lambda p: _projection(p, VARS))
    # Only the provenance stamp (file name, workbook hash) differs.
    assert {c["type"] for c in result["changes"]} == {"IR_METADATA_CHANGED"}


# ---------------------------------------------------------------------------
# R1C1 rewriting
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("formula", "refs"),
    [
        ("RC[-1]*2", ["RC[-1]"]),
        ("R[-1]C+1", ["R[-1]C"]),
        ("RC5+R5C", ["RC5", "R5C"]),
        ("R1C1+Sheet2!R[2]C[-3]", ["R1C1", "Sheet2!R[2]C[-3]"]),
        ("ROUND(RC[-1],2)", ["RC[-1]"]),
        ('Proj_RC+"RC[-1]"+RCOUNT', []),
    ],
)
def test_tokenizer_recognises_every_reference_form(formula: str, refs: list[str]) -> None:
    assert [t for kind, t in _tokenize_r1c1(formula) if kind == "REF"] == refs


def test_relative_reference_across_a_deleted_column_is_rewritten() -> None:
    # Columns D:E (4, 5) deleted; columns 6..11 move to 4..9.
    cols = AxisIndex({1: 1, 2: 2, 3: 3, **{c: c - 2 for c in range(6, 12)}}, frozenset({4, 5}))
    rows = AxisIndex({r: r for r in range(1, 14)})

    def rewrite(formula: str, col: int) -> str | None:
        return rewrite_formula_r1c1(
            formula,
            {"Proj": "Proj"},
            {},
            {},
            position=(2, col),
            own_sheet="Proj",
            row_index=rows,
            col_index=cols,
        )

    assert rewrite("RC[-5]*5+5", 6) == "RC[-3]*5+5"  # reads column A across the gap
    assert rewrite("RC[-1]+R[-1]C", 7) == "RC[-1]+R[-1]C"  # neighbours on one side
    assert rewrite("RC[-2]", 6) is None  # reads deleted column D
    assert rewrite("Other!RC[-5]", 6) == "Other!RC[-5]"  # another sheet's axes


def test_axis_index_interpolates_unoccupied_indices() -> None:
    index = AxisIndex({2: 2, 6: 4, 7: 5}, frozenset({4}))
    assert index.map(4) is None  # deleted
    assert index.map(6) == 4  # matched
    assert index.map(8) == 6  # above every key: keeps its distance from 7
    assert index.map(1) == 1  # below every key: keeps its distance from 2
    assert index.edited
    assert not AxisIndex({1: 1, 2: 2}).edited


def test_loose_signature_blanks_only_the_edited_axis() -> None:
    assert _loosen_formula("RC[-5]*R[-1]C3+R2C", "col") == "RC**R[-1]C*+R2C*"
    assert _loosen_formula("RC[-5]*R[-1]C3+R2C", "row") == "R*C[-5]*R*C3+R*C"
    # Function names, defined names and string literals are left alone.
    assert _loosen_formula('ROUND(Proj_RC,2)&"RC[-1]"', "col") == 'ROUND(Proj_RC,2)&"RC[-1]"'


def test_name_destination_follows_its_shifted_columns(tmp_path: Path) -> None:
    result = _diff(
        tmp_path,
        lambda p: _projection(p, VARS),
        lambda p: _projection(p, [v for v in VARS if v != "v2"]),
    )
    assert _of(result, "NAME_DESTINATIONS_CHANGED") == []
    assert [c["name"] for c in _of(result, "NAME_REMOVED")] == ["Proj_v2"]
    json.dumps(result)  # still serialisable with the new detail field
