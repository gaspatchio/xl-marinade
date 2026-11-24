"""Tests for the formula graph loaded from the extraction IR."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import pytest

from tests.calc.conftest import F
from xl_marinade.calc.graph import ListItems, WorkbookGraph
from xl_marinade.calc.models import CellRef
from xl_marinade.errors import CalcInputError

Book = Callable[..., Path]
IrFor = Callable[[Path], Path]

A1, A2 = CellRef("S", 1, 1), CellRef("S", 2, 1)
B1, C1, D1 = CellRef("S", 1, 2), CellRef("S", 1, 3), CellRef("S", 1, 4)


@pytest.fixture
def graph(make_book: Book, ir_for: IrFor) -> WorkbookGraph:
    path = make_book(
        "g",
        {
            "S": {
                "A1": 2,
                "A2": 3,
                "B1": F("=SUM(A1:A2)", 5),
                "C1": F("=B1*2", 10),
                "D1": F("=C1+A1", 12),
                "E2": F("=NA()", "#N/A"),
                "E3": F("=1=1", True),
                "E4": F('="hi"', "hi"),
                "E5": F("=B1+1", 6),
                "F5": F("=C1+1", 11),
                "G1": "text",
            }
        },
        names={"rate": "S!$A$1"},
        validations=[("S!E1", ["M", "F"])],
    )
    return WorkbookGraph.from_ir(ir_for(path))


def test_metadata_and_formulas(graph: WorkbookGraph) -> None:
    assert len(graph.workbook_sha256) == 64
    assert graph.formula(B1) == "SUM(A1:A2)"
    assert graph.formula(A1) is None
    assert graph.formula(CellRef("S", 99, 99)) is None
    cells = graph.formula_cells()
    assert cells == sorted(cells)
    assert B1 in cells and A1 not in cells


def test_precedents_expand_ranges(graph: WorkbookGraph) -> None:
    assert graph.precedents(B1) == frozenset({A1, A2})
    assert graph.precedents(D1) == frozenset({C1, A1})
    assert graph.precedents(A1) == frozenset()


def test_downstream_maps_to_smallest_root(graph: WorkbookGraph) -> None:
    got = graph.downstream([B1])
    assert got[B1] == B1 and got[C1] == B1 and got[D1] == B1
    both = graph.downstream([C1, A2])
    # CellRef orders by row first, so C1 < A2: C1 owns what it reaches, A2 the rest.
    assert both[C1] == C1 and both[D1] == C1
    assert both[A2] == A2 and both[B1] == A2
    assert graph.downstream([C1, A2]) == graph.downstream([A2, C1])


def test_stored_values(graph: WorkbookGraph) -> None:
    assert graph.stored(C1) == 10.0
    assert graph.stored(A1) == 2.0
    assert graph.stored(CellRef("S", 2, 5)) == "#N/A"
    assert graph.stored(CellRef("S", 3, 5)) is True
    assert graph.stored(CellRef("S", 4, 5)) == "hi"
    assert graph.stored(CellRef("S", 1, 7)) == "text"
    assert graph.stored(CellRef("S", 1, 5)) is None
    assert graph.has_stored(C1) and not graph.has_stored(CellRef("S", 1, 5))
    assert not graph.has_stored(CellRef("S", 50, 50))


def test_family_shared_by_copied_formulas(graph: WorkbookGraph) -> None:
    assert graph.family(CellRef("S", 5, 5)) == graph.family(CellRef("S", 5, 6))
    assert graph.family(B1) != graph.family(C1)
    assert graph.family(B1).startswith("S|")


def test_upstream_closure(graph: WorkbookGraph) -> None:
    assert graph.upstream([D1]) == {D1, C1, B1, A1, A2}


def test_resolve_target(graph: WorkbookGraph) -> None:
    assert graph.resolve_target("rate") == [A1]
    assert graph.resolve_target("S!B1") == [B1]
    assert graph.resolve_target("S!A1:A2") == [A1, A2]
    with pytest.raises(CalcInputError):
        graph.resolve_target("nope")
    with pytest.raises(CalcInputError):
        graph.resolve_target("Missing!A1")
    assert graph.name_cells("rate") == [A1]
    found = graph.lookup_name("rate", "S")
    assert found is not None and found.scope is None


def test_allowed_values_and_flags(graph: WorkbookGraph) -> None:
    assert graph.allowed_values(CellRef("S", 1, 5)) == ListItems(["M", "F"], literal=True)
    assert graph.allowed_values(A1) is None
    assert graph.udf_cells == frozenset()
    assert graph.external_cells == {}


def test_allowed_values_from_range_source(make_book: Book, ir_for: IrFor) -> None:
    path = make_book("r", {"S": {"A1": "x", "A2": "y", "A3": 3, "B1": 1}})
    import openpyxl
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.load_workbook(path)
    dv = DataValidation(type="list", formula1="$A$1:$A$3")
    wb["S"].add_data_validation(dv)
    dv.add("B1")
    wb.save(path)
    g = WorkbookGraph.from_ir(ir_for(path))
    # Items keep their stored type: a number cell is a number item.
    assert g.allowed_values(CellRef("S", 1, 2)) == ListItems(["x", "y", 3.0], literal=False)


def test_range_target_includes_empty_cells(make_book: Book, ir_for: IrFor) -> None:
    path = make_book(
        "gap",
        {"S": {"C11": 1, "C13": 3}},
        names={"col": "S!$C$11:$C$13"},
    )
    g = WorkbookGraph.from_ir(ir_for(path))
    want = [CellRef("S", 11, 3), CellRef("S", 12, 3), CellRef("S", 13, 3)]
    assert g.resolve_target("S!C11:C13") == want
    assert g.resolve_target("col") == want
    assert g.name_cells("col") == want


def test_defined_name_edges(make_book: Book, ir_for: IrFor) -> None:
    path = make_book(
        "nm",
        {
            "S": {
                "A1": 2,
                "A2": 3,
                "A3": 4,
                "B1": F("=rate*2", 4),
                "B2": F("=SUM(xs)", 9),
            }
        },
        names={"rate": "S!$A$1", "xs": "S!$A$1:$A$3"},
    )
    g = WorkbookGraph.from_ir(ir_for(path))
    b1, b2 = CellRef("S", 1, 2), CellRef("S", 2, 2)
    assert A1 in g.precedents(b1)
    assert {A1, A2, CellRef("S", 3, 1)} <= g.precedents(b2)
    assert g.downstream([A1])[b1] == A1
    assert b2 in g.downstream([A2])


def test_formula_names_and_array_cells(make_book: Book, ir_for: IrFor) -> None:
    from openpyxl.worksheet.formula import ArrayFormula

    path = make_book(
        "n",
        {"S": {"A1": 1, "B1": F(ArrayFormula("B1", "=A1:A2*2"), 2), "C1": F("=A1", 1)}},
        names={"f": "TODAY()", "k": "0.05", "r": "S!$A$1:$A$2", "c": "S!$A:$A"},
    )
    g = WorkbookGraph.from_ir(ir_for(path))

    def kind(name: str) -> str | None:
        found = g.lookup_name(name, "S")
        return found.kind if found is not None else None

    assert kind("f") == "formula" and kind("k") == "formula"
    assert kind("r") == "cells" and kind("c") == "cells"
    assert kind("missing") is None
    assert g.is_array_formula(B1) and not g.is_array_formula(C1)


def test_whole_column_range_buckets_only_the_used_rows(make_book: Book, ir_for: IrFor) -> None:
    path = make_book(
        "wc",
        {"In": {"A1": 1, "A2": 2, "A3": 3}, "S": {"B1": F("=SUM(In!A:A)", 6)}},
    )
    g = WorkbookGraph.from_ir(ir_for(path))
    b1 = CellRef("S", 1, 2)
    assert g.precedents(b1) == {CellRef("In", r, 1) for r in (1, 2, 3)}
    assert g.dependents(CellRef("In", 2, 1)) == {b1}
    # One bucket for the three used rows, not one per 256 rows of the column.
    assert len(g._range_grid) == 1


def test_name_over_external_range_is_not_cells(
    make_book: Book, ir_for: IrFor, tmp_path: Path
) -> None:
    from tests.calc.conftest import add_external_link

    base = make_book(
        "xn",
        {"S": {"A1": 1, "B1": F("=SUM(Ext)", 6)}},
        names={"Ext": "[1]Src!$A$1:$A$3"},
    )
    book = add_external_link(base, tmp_path / "xn2.xlsx", "Src", {"A1": "1", "A2": "2"})
    g = WorkbookGraph.from_ir(ir_for(book))
    found = g.lookup_name("Ext", "S")
    assert found is not None and found.kind == "external"
    assert g.name_cells("Ext") == []
    assert g.lookup_name("missing", "S") is None
