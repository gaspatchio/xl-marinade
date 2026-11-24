"""Tests for probe cells: branch outcomes, lookup edges, near flips and dynamic refs."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import pytest
from openpyxl.worksheet.formula import ArrayFormula

pytest.importorskip("ironcalc")

from tests.calc.conftest import F  # noqa: E402
from xl_marinade.calc.engine import Engine  # noqa: E402
from xl_marinade.calc.graph import NameDef, WorkbookGraph  # noqa: E402
from xl_marinade.calc.models import CellRef  # noqa: E402
from xl_marinade.calc.probes import (  # noqa: E402
    PROBE_SHEET,
    ProbeKind,
    ProbePlan,
    ProbeReadings,
    anchor,
    near_flip,
    place_probes,
    plan_probes,
    read_probes,
)
from xl_marinade.core.parser import FormulaParser  # noqa: E402

Book = Callable[..., Path]
IrFor = Callable[[Path], Path]


def _c(a1: str, sheet: str = "S") -> CellRef:
    return CellRef(sheet, *_rc(a1))


def _rc(a1: str) -> tuple[int, int]:
    return CellRef.parse(f"X!{a1}")[1:]


class _Names:
    """Names the formula's sheet owns, each over two areas (which cannot be lifted)."""

    def __init__(self, scoped: set[str]) -> None:
        self.scoped = scoped

    def lookup_name(self, ref: str, sheet: str | None) -> NameDef | None:
        if ref not in self.scoped or sheet is None:
            return None
        rects = ((sheet, 1, 1, 1, 1), (sheet, 3, 1, 3, 1))
        return NameDef(ref, sheet, "cells", (f"{sheet}!A1", f"{sheet}!A3"), rects)


def _setup(
    make_book: Book, ir_for: IrFor, cells: dict[str, object], *, cap: int = 50_000
) -> tuple[WorkbookGraph, Engine, ProbePlan]:
    book = make_book("p", {"S": cells})
    graph = WorkbookGraph.from_ir(ir_for(book))
    engine = Engine.load(book)
    plan = plan_probes(graph, graph.formula_cells(), cap=cap)
    place_probes(engine, plan)
    return graph, engine, plan


def _read(engine: Engine, plan: ProbePlan) -> ProbeReadings:
    engine.evaluate()
    return read_probes(engine, plan)


# -- near_flip ---------------------------------------------------------------


def test_near_flip() -> None:
    assert near_flip("ROUND", 1.005, 2) is True  # 100.49999999999999 is 1.4e-14 from the tie
    assert near_flip("ROUND", 0.125, 2) is True
    assert near_flip("INT", 3.0000000000001, None) is True
    assert near_flip("ROUND", 0.123, 2) is False


def test_near_flip_other_functions() -> None:
    assert near_flip("ROUNDUP", 1.2300000000000002, 2) is True
    assert near_flip("ROUNDDOWN", 1.235, 2) is False
    # Exactly on an integer / multiple is not near (ruling); a hair off it is.
    assert near_flip("TRUNC", 7.0, None) is False
    assert near_flip("TRUNC", 7.000000000000001, None) is True
    assert near_flip("INT", 3.0, None) is False
    assert near_flip("ROUNDUP", 1.5, 1) is False
    assert near_flip("CEILING", 6.0, 3) is False
    assert near_flip("FLOOR", 6.0, None) is False
    assert near_flip("TRUNC", 7.5, None) is False
    assert near_flip("INT", 3.5, None) is False
    assert near_flip("MROUND", 5.0, 2) is True  # 2.5 is a half-integer
    assert near_flip("MROUND", 5.6, 2) is False
    assert near_flip("MROUND", 5.0, 0) is False
    assert near_flip("CEILING", 6.000000000000001, 3) is True
    assert near_flip("FLOOR", 6.5, 3) is False
    assert near_flip("CEILING", 2.5, None) is False
    assert near_flip("ROUND", 125.0, -1) is True  # 12.5
    assert near_flip("ROUND", float("nan"), 2) is False


def test_near_flip_comparisons() -> None:
    assert near_flip(">=", 0.30000000000000004, 0.3) is True
    assert near_flip("<", 1e15, 1e15 + 0.125) is True  # relative scale
    assert near_flip("=", 0.3, 0.31) is False
    # Exact equality is not near: whole-number periods sit on ``t=0`` exactly.
    assert near_flip("lookup", 2.0, 2.0) is False
    assert near_flip("=", 0.0, 0.0) is False
    assert near_flip(">=", 0.3, 0.3) is False
    assert near_flip("lookup", 2.0, 2.0000000000000004) is True
    with pytest.raises(ValueError, match="FOO"):
        near_flip("FOO", 1.0, None)


# -- anchor ------------------------------------------------------------------


def _anchor(text: str, sheet: str = "Life Table ", **kw: object) -> str:
    return anchor(FormulaParser().parse(text), sheet, **kw)  # type: ignore[arg-type]


def test_anchor_quotes_sheet_and_absolutises() -> None:
    # Ruling: already-qualified references keep their sheet and become absolute too.
    assert _anchor("L45>='Proj Sheet'!A1") == "'Life Table '!$L$45>='Proj Sheet'!$A$1"


def test_anchor_ranges_names_and_structure() -> None:
    assert _anchor("SUM(A1:$B2)", "S") == "SUM(S!$A$1:$B$2)"
    assert _anchor("SUM(A:B)+SUM(2:$3)", "S") == "SUM(S!$A:$B)+SUM(S!$2:$3)"
    assert _anchor("Other!A1:B2", "S") == "Other!$A$1:$B$2"
    assert _anchor("Rate*2", "S") == "Rate*2"
    assert _anchor("Other!Rate", "S") == "Other!Rate"
    assert _anchor('IF(A1="x",-(B1+1),2.5%)', "S") == 'IF(S!$A$1="x",-(S!$B$1+1),2.5%)'
    assert _anchor("(A1-B1)-(C1-D1)", "S") == "(S!$A$1-S!$B$1)-(S!$C$1-S!$D$1)"
    assert _anchor("_xlfn.IFS(A1>1,1,TRUE,0)", "S") == "_xlfn.IFS(S!$A$1>1,1,TRUE,0)"


def test_anchor_rewrites_indirect_to_the_source_sheet() -> None:
    got = _anchor('INDIRECT("A"&C1)', "It's")
    assert got == (
        "INDIRECT(IF(ISERROR(FIND(\"!\",\"A\"&'It''s'!$C$1)),"
        "\"'It''s'!\"&(\"A\"&'It''s'!$C$1),\"A\"&'It''s'!$C$1))"
    )
    assert _anchor("INDIRECT(B1,TRUE)", "S").startswith("INDIRECT(IF(")


@pytest.mark.parametrize(
    "text",
    [
        "LET(x,A1,x+1)",
        "_xlfn.LAMBDA(x,x+1)(A1)",
        "SUM({1,2,3})",
        "ROW()+A1",
        "COLUMN()",
        'INDIRECT("R1C1",FALSE)',
        "Table1[Amount]",
        "Local*2",
        "[1]Other!B2+1",
        "'[1]My Other'!$B$2+1",
    ],
)
def test_anchor_refuses_what_cannot_be_lifted(text: str) -> None:
    with pytest.raises(ValueError):
        _anchor(text, "S", names=_Names({"Local"}))


# -- plan / place / read -------------------------------------------------------


def test_branch_probe_reads_outcome(make_book: Book, ir_for: IrFor) -> None:
    graph, engine, plan = _setup(
        make_book, ir_for, {"A1": 2, "B1": F('=IF(A1>1,"big","small")', "big")}
    )
    (branch,) = [p for p in plan.probes if p.kind is ProbeKind.branch]
    assert branch.cell == _c("B1") and branch.ordinal == 0
    assert branch.family == graph.family(_c("B1"))
    assert branch.meta["function"] == "IF"
    assert _read(engine, plan).branch == {(_c("B1"), 0): "TRUE"}
    engine.set_value(_c("A1"), 0.0)
    assert _read(engine, plan).branch == {(_c("B1"), 0): "FALSE"}
    engine.set_value(_c("A1"), "text")  # text sorts above every number
    assert _read(engine, plan).branch == {(_c("B1"), 0): "TRUE"}
    engine.set_formula(_c("A1"), "1/0")
    assert _read(engine, plan).branch == {(_c("B1"), 0): "error"}


def test_branch_kinds(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "A1": 4,
        "A2": 0,
        "B1": F('=_xlfn.IFS(A1>5,"a",A1>3,"b")', "b"),
        "B2": F('=_xlfn.IFS(A1>5,"a",A1>9,"b")', "#N/A"),
        "B3": F('=_xlfn.SWITCH(A1,1,"x",4,"y","z")', "y"),
        "B4": F('=_xlfn.SWITCH(A1,1,"x",2,"y")', "#N/A"),
        "B5": F('=CHOOSE(A1/2,"p","q")', "q"),
        "B6": F('=CHOOSE(A1,"p","q")', "#VALUE!"),
        "B7": F("=IFERROR(1/A2,0)", 0),
        "B8": F("=IFERROR(1/A1,0)", 0.25),
        "B9": F("=_xlfn.IFNA(MATCH(A1,A1:A2,0),0)", 1),
        "B10": F('=IF(A1>1,IF(A2>1,"a","b"),"c")', "b"),
    }
    _graph, engine, plan = _setup(make_book, ir_for, cells)
    got = _read(engine, plan).branch
    assert got == {
        (_c("B1"), 0): "2",
        (_c("B2"), 0): "none",
        (_c("B3"), 0): "2",
        (_c("B4"), 0): "error",  # no case, no default: the cell is #N/A
        (_c("B5"), 0): "2",
        (_c("B6"), 0): "error",
        (_c("B7"), 0): "error",
        (_c("B8"), 0): "ok",
        (_c("B9"), 0): "ok",
        (_c("B10"), 0): "TRUE",
        (_c("B10"), 1): "FALSE",
    }


def test_lookup_edge_classes(make_book: Book, ir_for: IrFor) -> None:
    table: dict[str, object] = {
        "A1": 10,
        "A2": 20,
        "A3": 30,
        "A4": 40,
        "B1": 1,
        "B2": 2,
        "B3": 3,
        "B4": 4,
        "D1": 10,
        "E1": 20,
        "F1": 30,
    }
    cells: dict[str, object] = {
        **table,
        "C1": F("=VLOOKUP(10,A1:B4,2)", 1),  # first
        "C2": F("=VLOOKUP(45,A1:B4,2,TRUE)", 4),  # beyond end: last
        "C3": F("=VLOOKUP(25,$A$1:$B$4,2)", 2),  # interior
        "C4": F("=VLOOKUP(99,A1:B4,2,FALSE)", "#N/A"),  # exact miss
        "C5": F("=VLOOKUP(5,A1:B4,2)", "#N/A"),  # before first
        "C6": F("=INDEX(B1:B4,MATCH(40,A1:A4,0))", 4),
        "C7": F("=HLOOKUP(30,D1:F1,1)", 30),
        "C8": F("=LOOKUP(20,A1:B4)", 2),
        "C9": F("=LOOKUP(25,D1:F1)", 20),
    }
    _graph, engine, plan = _setup(make_book, ir_for, cells)
    got = _read(engine, plan).lookup_edge
    assert got == {
        (_c("C1"), 0): "first",
        (_c("C2"), 0): "last",
        (_c("C3"), 0): "interior",
        (_c("C4"), 0): "not_found",
        (_c("C5"), 0): "not_found",
        (_c("C6"), 0): "last",
        (_c("C7"), 0): "last",
        (_c("C8"), 0): "interior",
        (_c("C9"), 0): "interior",
    }


def test_near_rounding_and_comparison_readings(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "A1": 1.005,
        "A2": 0.1,
        "A3": 0.2,
        "A4": 20,
        "A5": 0.123,
        "K1": 10,
        "K2": 20,
        "K3": 30,
        "B1": F("=ROUND(A1,2)", 1.0),
        "B2": F("=ROUND(A5,2)", 0.12),
        "B3": F('=IF(A2+A3>=0.3,"y","n")', "y"),
        "B4": F('=IF(A2+A3>=0.5,"y","n")', "n"),
        "B5": F("=VLOOKUP(A4,K1:K3,1)", 20),
        "B6": F("=VLOOKUP(A4+1,K1:K3,1)", 20),
        "B7": F("=MATCH(A4,K1:K3,0)", 2),  # exact: no knife edge
        "B8": F("=INT(A4/10)", 2),
    }
    _graph, engine, plan = _setup(make_book, ir_for, cells)
    rounding = [p for p in plan.probes if p.kind is ProbeKind.rounding]
    assert [p.meta["function"] for p in rounding] == ["ROUND", "ROUND", "INT"]
    assert not any(p.cell == _c("B7") for p in plan.probes if p.kind is ProbeKind.comparison)
    got = _read(engine, plan)
    assert got.near_rounding == {_c("B1")}  # INT(2.0) sits exactly on the boundary
    # B5 matches its key exactly: decided, not near (only 0 < distance <= tol is).
    assert got.near_comparison == {_c("B3")}


def test_dynamic_ref_probe(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "C1": 4,
        "D1": 3,
        "B1": F("=SUM(OFFSET(A1,C1,0,2,1))", 0),
        "B2": F('=INDIRECT("A"&D1)', 0),
        "B3": F('=SUM(INDIRECT("Other!B"&D1&":C"&D1))', 0),
        "B4": F("=OFFSET($A$1,1,1)", 0),
    }
    book = make_book("p", {"S": cells, "Other": {"A1": 1}})
    graph = WorkbookGraph.from_ir(ir_for(book))
    engine = Engine.load(book)
    plan = plan_probes(graph, graph.formula_cells(), cap=50_000)
    place_probes(engine, plan)
    got = _read(engine, plan).dynamic_refs
    assert got[_c("B1")] == [_c("A5"), _c("A6")]
    assert got[_c("B2")] == [_c("A3")]
    assert got[_c("B3")] == [_c("B3", "Other"), _c("C3", "Other")]
    assert got[_c("B4")] == [_c("B2")]
    assert _read(engine, plan).dynamic_unread == set()
    engine.set_value(_c("C1"), 0.0)
    assert _read(engine, plan).dynamic_refs[_c("B1")] == [_c("A1"), _c("A2")]


def test_unreadable_dynamic_refs_are_reported(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "C1": 4,
        "D1": "A:A",
        "B1": F("=SUM(OFFSET(A1,C1,0,2,1))", 0),
        "B2": F("=SUM(INDIRECT(D1))", 0),  # whole column: text not a plain range
        "B3": F("=SUM(OFFSET(A1,C1,0),INDIRECT(D1))", 0),
    }
    _graph, engine, plan = _setup(make_book, ir_for, cells)
    got = _read(engine, plan)
    assert got.dynamic_unread == {_c("B2"), _c("B3")}
    assert got.dynamic_refs[_c("B2")] == []
    assert got.dynamic_refs[_c("B3")] == [_c("A5")]  # the readable one is kept
    engine.set_value(_c("C1"), -5.0)  # OFFSET off the sheet: #REF!
    got = _read(engine, plan)
    assert _c("B1") in got.dynamic_unread and got.dynamic_refs[_c("B1")] == []


def test_branch_probe_lifts_only_the_deciding_arguments(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "A1": 2,
        "B1": F("=IF(A1>1,1,ROW())", 1),
        "B2": F("=IF(A1>1,IF(A1>5,ROW(),2),3)", 2),
        "B3": F("=_xlfn.SWITCH(A1,2,COLUMN(),0)", 2),
        "B4": F("=IFERROR(1/A1,ROW())", 0.5),
    }
    _graph, engine, plan = _setup(make_book, ir_for, cells)
    assert plan.unplaced == {}
    assert _read(engine, plan).branch == {
        (_c("B1"), 0): "TRUE",
        (_c("B2"), 0): "TRUE",
        (_c("B2"), 1): "FALSE",
        (_c("B3"), 0): "1",
        (_c("B4"), 0): "ok",
    }


def test_error_code_text_is_not_lifted() -> None:
    with pytest.raises(ValueError, match="error-code"):
        _anchor('IF(A1="#REF!",1,0)', "S")
    assert _anchor('IF(A1="#REFS",1,0)', "S") == 'IF(S!$A$1="#REFS",1,0)'


def test_error_code_text_goes_unplaced(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {"A1": "#REF!", "B1": F('=IF(A1="#REF!",1,0)', 1)}
    _graph, _engine, plan = _setup(make_book, ir_for, cells)
    assert plan.probes == []
    assert plan.unplaced[_c("B1")] == [ProbeKind.branch, ProbeKind.comparison]


def test_array_contexts_go_unplaced(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "A1": 2,
        "A2": 3,
        "B1": F(ArrayFormula("B1", "=SUM(IF(A1:A2>1,1,0))"), 2),
        "B2": F("=SUMPRODUCT(ROUND(A1:A2,0))+ROUND(A1,0)", 7),
        "B3": F('=IF(A1>1,"a","b")', "a"),
    }
    _graph, _engine, plan = _setup(make_book, ir_for, cells)
    assert plan.unplaced == {
        _c("B1"): [ProbeKind.branch, ProbeKind.comparison],
        _c("B2"): [ProbeKind.rounding],
    }
    assert [(p.cell, p.kind, p.ordinal) for p in plan.probes] == [
        (_c("B2"), ProbeKind.rounding, 1),
        (_c("B3"), ProbeKind.branch, 0),
        (_c("B3"), ProbeKind.comparison, 0),
    ]


def test_unliftable_goes_unplaced(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "A1": 2,
        "B1": F('=_xlfn.LET(_xlpm.x,A1,IF(_xlpm.x>1,"a","b"))', "a"),
        "B2": F('=IF(A1>1,"a","b")', "a"),
    }
    _graph, _engine, plan = _setup(make_book, ir_for, cells)
    assert plan.unplaced[_c("B1")] == [ProbeKind.branch, ProbeKind.comparison]
    assert {p.cell for p in plan.probes} == {_c("B2")}

    _graph, _engine, capped = _setup(make_book, ir_for, cells, cap=0)
    assert capped.probes == []
    assert capped.unplaced[_c("B2")] == [ProbeKind.branch, ProbeKind.comparison]


def test_cap_keeps_plan_order(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "A1": 2,
        "B1": F('=IF(A1>1,"a","b")', "a"),
        "B2": F("=ROUND(A1,0)", 2),
    }
    _graph, engine, plan = _setup(make_book, ir_for, cells, cap=1)
    assert [(p.cell, p.kind) for p in plan.probes] == [(_c("B1"), ProbeKind.branch)]
    assert plan.unplaced == {
        _c("B1"): [ProbeKind.comparison],
        _c("B2"): [ProbeKind.rounding],
    }
    engine.evaluate()
    assert engine.value(CellRef(PROBE_SHEET, 1, 1)) == "TRUE"
    props = {p["name"]: p["state"] for p in engine._model.get_worksheets_properties()}
    assert props[PROBE_SHEET] == "hidden"


def test_ordinals_count_per_kind_in_pre_order(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "A1": 2,
        "B1": F("=IF(A1>1,ROUND(A1,0),IF(A1<0,1,ROUND(A1,1)))", 2),
    }
    _graph, _engine, plan = _setup(make_book, ir_for, cells)
    got = [(p.kind.value, p.ordinal) for p in plan.probes]
    assert got == [
        ("branch", 0),
        ("comparison", 0),
        ("rounding", 0),
        ("branch", 1),
        ("comparison", 1),
        ("rounding", 1),
    ]


def test_probes_record_their_enclosing_arms(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "A1": 2,
        "B1": F("=IF(A1>1,ROUND(A1,0),IF(A1<0,1,ROUND(A1,1)))", 2),
        "B2": F("=_xlfn.IFS(A1>5,ROUND(A1,0),A1>1,INT(A1))", 2),
        "B3": F("=CHOOSE(A1,INT(A1),ROUND(A1,0))", 2),
        "B4": F("=IFERROR(INT(A1),ROUND(A1,0))", 2),
        "B5": F("=_xlfn.SWITCH(A1,1,INT(A1),ROUND(A1,0))", 2),
    }
    _graph, _engine, plan = _setup(make_book, ir_for, cells)
    guards = {(p.cell, p.kind.value, p.ordinal): p.guards for p in plan.probes}
    t, f = frozenset({"TRUE"}), frozenset({"FALSE"})
    assert guards[(_c("B1"), "branch", 0)] == ()
    assert guards[(_c("B1"), "rounding", 0)] == ((0, t, frozenset()),)
    assert guards[(_c("B1"), "branch", 1)] == ((0, f, frozenset()),)
    assert guards[(_c("B1"), "rounding", 1)] == ((0, f, frozenset()), (1, f, frozenset()))
    # IFS: a later condition runs unless an earlier one matched; a value only on its own.
    assert guards[(_c("B2"), "comparison", 0)] == ()
    assert guards[(_c("B2"), "comparison", 1)] == (
        (0, frozenset({"2", "none"}), frozenset({"error"})),
    )
    assert guards[(_c("B2"), "rounding", 0)] == ((0, frozenset({"1"}), frozenset()),)
    assert guards[(_c("B2"), "rounding", 1)] == ((0, frozenset({"2"}), frozenset()),)
    assert guards[(_c("B3"), "rounding", 0)] == ((0, frozenset({"1"}), frozenset()),)
    assert guards[(_c("B3"), "rounding", 1)] == ((0, frozenset({"2"}), frozenset()),)
    assert guards[(_c("B4"), "rounding", 0)] == ()
    assert guards[(_c("B4"), "rounding", 1)] == ((0, frozenset({"error"}), frozenset()),)
    assert guards[(_c("B5"), "rounding", 0)] == ((0, frozenset({"1"}), frozenset()),)
    assert guards[(_c("B5"), "rounding", 1)] == ((0, frozenset({"default"}), frozenset()),)


def test_readings_skip_probes_in_untaken_arms(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "A1": 0,
        "A2": 35,
        "T1": 10,
        "T2": 20,
        "T3": 30,
        "B1": F("=IF(A1>0.5,MATCH(A2,$T$1:$T$3,1),0)", 0),
    }
    _graph, engine, plan = _setup(make_book, ir_for, cells)
    off = _read(engine, plan)
    assert off.branch == {(_c("B1"), 0): "FALSE"}
    assert off.lookup_edge == {}
    engine.set_value(_c("A1"), 1.0)
    on = _read(engine, plan)
    assert on.lookup_edge == {(_c("B1"), 0): "last"}
    assert on.unsure == set()


def test_unplaced_enclosing_branch_makes_readings_unsure(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {
        "A1": 1,
        "A2": 35,
        "T1": 10,
        "T2": 20,
        "T3": 30,
        "B1": F("=IF(A1>0.5,MATCH(A2,$T$1:$T$3,1),0)", 3),
    }
    _graph, engine, plan = _setup(make_book, ir_for, cells, cap=0)
    assert plan.probes == []
    book_plan = ProbePlan(
        [p for p in plan_probes(_graph, [_c("B1")], cap=50).probes if p.kind is ProbeKind.lookup],
        {},
    )
    place_probes(engine, book_plan)
    got = _read(engine, book_plan)
    assert got.lookup_edge == {(_c("B1"), 0): "last"}
    assert got.unsure == {(ProbeKind.lookup, _c("B1"), 0)}
