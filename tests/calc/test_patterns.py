"""Tests for the static scan of formulas that reproduce only by luck."""

from __future__ import annotations

from collections.abc import Callable
from datetime import date
from pathlib import Path

from openpyxl.worksheet.formula import ArrayFormula

from tests.calc.conftest import F
from xl_marinade.calc.graph import WorkbookGraph
from xl_marinade.calc.models import CellRef, Reason, ReasonKind
from xl_marinade.calc.patterns import scan_patterns

Book = Callable[..., Path]
IrFor = Callable[[Path], Path]

UNSORTED = "approximate lookup on unsorted column"
MIXED = "approximate lookup on mixed text and numbers"
NON_RANGE = "approximate lookup on a non-range table"
IMPLICIT = "implicit intersection on multi-cell name"
IMPLICIT_RANGE = "implicit intersection on multi-cell range"
NAME_FORMULA = "name holds a formula"


def _c(a1: str) -> CellRef:
    return CellRef.parse(f"S!{a1}")


def _scan(
    make_book: Book,
    ir_for: IrFor,
    cells: dict[str, object],
    *,
    names: dict[str, str] | None = None,
    as_of: date | None = None,
    data_table_cells: list[CellRef] | None = None,
) -> dict[CellRef, list[Reason]]:
    book = make_book("p", {"S": cells}, names=names)
    graph = WorkbookGraph.from_ir(ir_for(book))
    return scan_patterns(graph, as_of=as_of, data_table_cells=data_table_cells or [])


def _details(found: dict[CellRef, list[Reason]], a1: str) -> list[str]:
    reasons = found.get(_c(a1), [])
    for r in reasons:
        assert r.kind is ReasonKind.refused_pattern
        assert r.root == _c(a1)
    return [r.detail for r in reasons]


# Keys in A1:A4 with values in B1:B4.
UNSORTED_KEYS: dict[str, object] = {"A1": 1, "A2": 5, "A3": 3, "A4": 7}
SORTED_KEYS: dict[str, object] = {"A1": 1, "A2": 3, "A4": 7}  # A3 blank: ignored
VALUES: dict[str, object] = {"B1": 10, "B2": 20, "B3": 30, "B4": 40}


def test_approximate_vlookup_on_unsorted_column(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            **UNSORTED_KEYS,
            **VALUES,
            "D1": F("=VLOOKUP(4,A1:B4,2)", 30),
            "D2": F("=VLOOKUP(4,A1:B4,2,TRUE)", 30),
            "D3": F("=VLOOKUP(4,$A$1:$B$4,2,1)", 30),
            "D4": F("=VLOOKUP(4,A1:B4,2,FALSE)", "#N/A"),
            "D5": F("=VLOOKUP(4,A1:B4,2,0)", "#N/A"),
            "D6": F("=vlookup(4,A1:B4,2,E1)", 30),
        },
    )
    for a1 in ("D1", "D2", "D3", "D6"):
        assert _details(found, a1) == [UNSORTED], a1
    assert _c("D4") not in found and _c("D5") not in found


def test_approximate_lookup_on_sorted_column_is_not_refused(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            **SORTED_KEYS,
            **VALUES,
            "A5": 7,  # ties are fine
            "D1": F("=VLOOKUP(4,A1:B5,2)", 20),
            "D2": F("=MATCH(4,A1:A5,1)", 2),
            "D3": F("=LOOKUP(4,A1:A5,B1:B5)", 20),
            "D4": F('=_xlfn.XLOOKUP(4,A1:A5,B1:B5,"none",0,2)', "#N/A"),
            "D5": F("=VLOOKUP(4,A:B,2)", 20),
        },
    )
    assert found == {}


def test_approximate_lookup_on_mixed_keys(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            "A1": 1,
            "A2": "b",
            "A3": 3,
            "C1": 1,
            "C2": True,
            "D1": F("=VLOOKUP(2,A1:B3,2)", 0),
            "D2": F("=MATCH(2,C1:C2)", 1),
        },
    )
    assert _details(found, "D1") == [MIXED]
    assert _details(found, "D2") == [MIXED]


def test_approximate_lookup_on_non_range_table(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            **SORTED_KEYS,
            **VALUES,
            "D1": F("=VLOOKUP(4,OFFSET(A1,0,0,4,2),2)", 20),
            "D2": F("=MATCH(4,INDEX(A1:B4,0,1))", 2),
            "D3": F("=VLOOKUP(4,OFFSET(A1,0,0,4,2),2,FALSE)", "#N/A"),
        },
    )
    assert _details(found, "D1") == [NON_RANGE]
    assert _details(found, "D2") == [NON_RANGE]
    assert _c("D3") not in found


def test_match_types_and_descending(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            "A1": 1,
            "A2": 3,
            "A3": 7,
            "C1": 9,
            "C2": 4,
            "C3": 2,
            "D1": F("=MATCH(4,A1:A3,-1)", "#N/A"),
            "D2": F("=MATCH(4,C1:C3,-1)", 2),
            "D3": F("=MATCH(4,C1:C3,0)", 2),
            "D4": F("=MATCH(4,C1:C3)", "#N/A"),
            "D5": F("=MATCH(4,S!$A$1:S!$A3,1)", 2),
            "D6": F("=MATCH(4,A1:A3,E1)", 2),
        },
    )
    assert _details(found, "D1") == [UNSORTED]
    assert _c("D2") not in found and _c("D3") not in found
    assert _details(found, "D4") == [UNSORTED]
    assert _c("D5") not in found
    # A non-literal match type could be 1 or -1: ascending keys are not safe for -1.
    assert _details(found, "D6") == [UNSORTED]


def test_hlookup_uses_first_row_and_lookup_is_always_approximate(
    make_book: Book, ir_for: IrFor
) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            "A1": 1,
            "B1": 5,
            "C1": 3,
            "A2": 1,
            "B2": 2,
            "C2": 3,  # second row sorted, first is not
            "A5": 2,
            "A6": 1,
            "D1": F("=HLOOKUP(4,A1:C2,2)", 3),
            "D2": F("=HLOOKUP(4,A2:C3,1)", 3),
            "D3": F("=LOOKUP(2,A5:A6)", 1),
            "D4": F("=LOOKUP(2,A5:A6,B5:B6)", 0),
        },
    )
    assert _details(found, "D1") == [UNSORTED]
    assert _c("D2") not in found
    assert _details(found, "D3") == [UNSORTED]
    assert _details(found, "D4") == [UNSORTED]


def test_approximate_lookup_through_a_defined_name(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {**UNSORTED_KEYS, **VALUES, "D1": F("=VLOOKUP(4,tbl,2)", 30)},
        names={"tbl": "S!$A$1:$B$4"},
    )
    assert _details(found, "D1") == [UNSORTED]


def test_implicit_intersection_on_multi_cell_name(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            "A1": 1,
            "A2": 2,
            "D1": F("=rate*2", 2),
            "D2": F("=rate", 1),
            "D3": F("=SUM(rate)", 3),
            "D4": F("=one*2", 2),
            "D5": F("=SUM(rate*2)", 2),
            "D6": F("=IF(rate>1,1,0)", 0),
        },
        names={"rate": "S!$A$1:$A$2", "one": "S!$A$1"},
    )
    assert _details(found, "D1") == [IMPLICIT]
    assert _details(found, "D2") == [IMPLICIT]
    assert _c("D3") not in found
    assert _c("D4") not in found
    assert _details(found, "D5") == [IMPLICIT]
    assert _details(found, "D6") == [IMPLICIT]


def test_random_functions(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {"D1": F("=RAND()", 0.5), "D2": F("=RANDBETWEEN(1,6)", 3), "D3": F("=rand()*2", 1.0)},
        as_of=date(2026, 1, 1),
    )
    assert _details(found, "D1") == ["RAND"]
    assert _details(found, "D2") == ["RANDBETWEEN"]
    assert _details(found, "D3") == ["RAND"]


def test_today_and_now_only_without_as_of(make_book: Book, ir_for: IrFor) -> None:
    cells: dict[str, object] = {"D1": F("=TODAY()", 46000), "D2": F("=NOW()+1", 46001.5)}
    found = _scan(make_book, ir_for, cells)
    assert _details(found, "D1") == ["TODAY"]
    assert _details(found, "D2") == ["NOW"]
    assert _scan(make_book, ir_for, cells, as_of=date(2026, 1, 1)) == {}


def test_circular_reference(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            "A1": F("=B1+1", 0),
            "B1": F("=A1+1", 0),
            "C1": F("=C1+1", 0),
            "D1": F("=A1*2", 0),  # downstream of the cycle, not in it
        },
    )
    assert _details(found, "A1") == ["circular reference"]
    assert _details(found, "B1") == ["circular reference"]
    assert _details(found, "C1") == ["circular reference"]
    assert _c("D1") not in found


def test_data_table_cells(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(make_book, ir_for, {"A1": 1}, data_table_cells=[_c("E5"), _c("E6")])
    assert _details(found, "E5") == ["data table"]
    assert _details(found, "E6") == ["data table"]


def test_unparseable_formula(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            "A1": 1,
            "B1": 1,
            "D1": F("=IF(A1,,1)", 0),
            "D2": F("=SUM({1,2})", 3),  # array constant: the parser mangles it
            "D3": F("=A1:A2 A1:B1", 1),  # the parser drops the intersection
            "D4": F("=A1+1", 2),
        },
    )
    for a1 in ("D1", "D2", "D3"):
        assert _details(found, a1) == ["unparseable formula"], a1
    assert _c("D4") not in found


def test_several_reasons_sorted_without_duplicates(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {"D1": F("=RAND()+NOW()+RAND()", 1.0)},
        data_table_cells=[_c("D1")],
    )
    assert _details(found, "D1") == ["NOW", "RAND", "data table"]


def test_binary_search_xlookup_and_whole_column_on_unsorted_keys(
    make_book: Book, ir_for: IrFor
) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            **UNSORTED_KEYS,
            **VALUES,
            "D1": F('=_xlfn.XLOOKUP(4,A1:A4,B1:B4,"none",0,2)', 30),
            "D2": F('=_xlfn.XLOOKUP(4,A1:A4,B1:B4,"none",0,1)', "none"),
            "D3": F("=VLOOKUP(4,A:B,2)", 30),
            "D4": F("=_xlfn.XMATCH(4,A1:A4,0,-2)", 2),
        },
    )
    assert _details(found, "D1") == [UNSORTED]
    assert _c("D2") not in found
    assert _details(found, "D3") == [UNSORTED]
    assert _details(found, "D4") == [UNSORTED]


def test_self_reference_through_a_range_is_circular(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(make_book, ir_for, {"A1": 1, "A2": 2, "A3": F("=SUM(A1:A3)", 3)})
    assert _details(found, "A3") == ["circular reference"]


def test_text_keys_outside_plain_collation_are_unsorted(make_book: Book, ir_for: IrFor) -> None:
    # Python orders these, Excel's collation does not (``_`` sorts before digits,
    # accented letters sit beside their base letter).
    found = _scan(
        make_book,
        ir_for,
        {
            "A1": "1",
            "A2": "_",
            "C1": "f",
            "C2": "é",
            "E1": "apple",
            "E2": "Banana",
            "E3": "cherry 2",
            "D1": F('=MATCH("x",A1:A2)', 2),
            "D2": F('=MATCH("x",C1:C2)', 2),
            "D3": F('=MATCH("x",E1:E3)', 3),
        },
    )
    assert _details(found, "D1") == [UNSORTED]
    assert _details(found, "D2") == [UNSORTED]
    assert _c("D3") not in found


def test_name_holding_a_formula_or_constant(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            "A1": 1,
            "A2": 2,
            "B1": 10,
            "B2": 20,
            "D1": F("=q*2", 20),
            "D2": F("=SUM(rate)", 0.05),
            "D3": F("=asof+1", 46001),
            "D4": F("=SUM(col)", 3),
            "D5": F("=SUM(r2)", 3),
        },
        names={
            "q": "VLOOKUP(S!$A$1,S!$A$1:$B$2,2,TRUE)",
            "asof": "TODAY()",
            "rate": "0.05",
            "col": "S!$A:$A",
            "r2": "S!$A$1:$A$2",
        },
        as_of=date(2026, 1, 1),
    )
    assert _details(found, "D1") == [NAME_FORMULA]
    assert _details(found, "D2") == [NAME_FORMULA]
    assert _details(found, "D3") == [NAME_FORMULA]
    assert _c("D4") not in found and _c("D5") not in found


def test_implicit_intersection_on_multi_cell_range(make_book: Book, ir_for: IrFor) -> None:
    found = _scan(
        make_book,
        ir_for,
        {
            "A1": 1,
            "A2": 2,
            "A3": 3,
            "D1": F("=A1:A3*2", 2),
            "D2": F("=A1:A3", 1),
            "D3": F("=SUM(A1:A3*2)", 2),
            "D4": F("=SUM(A1:A3)", 6),
            "D5": F("=A1*2", 2),
            "D6": F(ArrayFormula("D6", "=A1:A3*2"), 2),
            "D7": F("=$A$1:$A$1*2", 2),
        },
    )
    assert _details(found, "D1") == [IMPLICIT_RANGE]
    assert _details(found, "D2") == [IMPLICIT_RANGE]
    assert _details(found, "D3") == [IMPLICIT_RANGE]
    for a1 in ("D4", "D5", "D6", "D7"):
        assert _c(a1) not in found, a1
