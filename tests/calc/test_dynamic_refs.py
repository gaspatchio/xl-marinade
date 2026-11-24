"""Dynamic references (``OFFSET``, ``INDIRECT``, ranges built by a call) are followed
through their arguments, their built range ends and the cells they land on.

The device for a wrong value is ``DATE(1900,1,1)``: Excel gives 1 (its 1900 date
system starts there), IronCalc 0.9.1 gives 2. Every other value is worked by hand from
Excel's documented semantics: ``OFFSET(base,r,c)`` is the cell ``r`` rows below and
``c`` columns right of ``base``; ``INDIRECT(text)`` is the cell the address text names
(on the formula's own sheet when the text names none); ``B1:OFFSET(...)`` is the
smallest range holding ``B1`` and the cell ``OFFSET`` returns; ``ROUND`` rounds an exact
half away from zero.
"""

from __future__ import annotations

import csv
import json
from collections.abc import Callable
from pathlib import Path

import pytest

pytest.importorskip("ironcalc")

from tests.calc.conftest import F  # noqa: E402
from xl_marinade.calc import CalcRequest, CalcResult, calc  # noqa: E402
from xl_marinade.calc.graph import open_graph  # noqa: E402
from xl_marinade.calc.models import CellRef  # noqa: E402
from xl_marinade.calc.patterns import INPUT_KEYS, input_dependent_lookups  # noqa: E402
from xl_marinade.calc.probes import ProbeKind, plan_probes  # noqa: E402

Book = Callable[..., Path]
_RANK = {"ok": 0, "unverified_path": 1, "uncertain": 2, "refused": 3}


def _run(
    tmp: Path,
    book: Path,
    policies: str,
    fields: dict[str, str],
    outputs: list[str],
    *,
    check: dict[str, object] | None = None,
    tag: str = "out",
) -> CalcResult:
    (tmp / "p.csv").write_text(policies, encoding="utf-8")
    (tmp / "m.json").write_text(json.dumps({"id_field": "id", "fields": fields}), encoding="utf-8")
    check_path = None
    if check is not None:
        check_path = tmp / "c.json"
        check_path.write_text(json.dumps(check), encoding="utf-8")
    return calc(
        CalcRequest(
            workbook=book,
            policies=tmp / "p.csv",
            input_map=tmp / "m.json",
            out_dir=tmp / tag,
            outputs=outputs,
            check_against=check_path,
        )
    )


def _rows(res: CalcResult) -> dict[tuple[str, str], dict[str, str]]:
    with res.results_csv.open(encoding="utf-8", newline="") as fh:
        return {(r["policy_id"], r["output"]): r for r in csv.DictReader(fh)}


# -- (a) arguments --------------------------------------------------------------

_BAND = [100, 200, 300, 300, 500, 600, 700, 800]  # Tab!A1:A8, flat at rows 3-4


def _argument_book(make_book: Book) -> Path:
    # Stored x = 3. C1 = x - DATE(1900,1,1) + 1: Excel 3 (IronCalc 2). C2 = IF(x>4,
    # DATE(1900,1,1)+2, 1): Excel 1 (the FALSE arm). D1 = OFFSET(Tab!A1, C1, 0) = A4 =
    # 300, which IronCalc reproduces by landing on A3 = 300. D2 the same through
    # INDIRECT's text. D8 = OFFSET(Tab!A1, C2, 0) = A2 = 200.
    tab = {f"A{i}": v for i, v in enumerate(_BAND, 1)}
    return make_book(
        "args",
        {
            "Inputs": {"B1": 3},
            "Calc": {
                "C1": F("=Inputs!$B$1-DATE(1900,1,1)+1", 3),
                "C2": F("=IF(Inputs!$B$1>4,DATE(1900,1,1)+2,1)", 1),
                "D1": F("=OFFSET(Tab!$A$1,Calc!C1,0)", 300),
                "D2": F('=INDIRECT("Tab!A"&(Calc!C1+1))', 300),
                "D8": F("=OFFSET(Tab!$A$1,Calc!C2,0)", 200),
            },
            "Tab": tab,
        },
    )


def test_an_argument_that_fails_the_trust_run_refuses_the_reference(
    make_book: Book, tmp_path: Path
) -> None:
    # x = 5: Excel C1 = 5, D1 = D2 = A6 = 600; IronCalc lands on A5 = 500.
    book = _argument_book(make_book)
    res = _run(tmp_path, book, "id,x\nP,5\n", {"x": "Inputs!B1"}, ["Calc!D1", "Calc!D2"])
    rows = _rows(res)
    for out in ("Calc!D1", "Calc!D2"):
        row = rows[("P", out)]
        assert row["status"] == "refused", (out, row)
        assert "silently_wrong@Calc!C1" in row["reasons"], (out, row)


def test_an_argument_on_an_unexercised_arm_is_unverified(make_book: Book, tmp_path: Path) -> None:
    # x = 5: Excel C2 = DATE(1900,1,1)+2 = 3, D8 = A4 = 300 (IronCalc 4 -> 500). The
    # trust run never took C2's TRUE arm.
    book = _argument_book(make_book)
    res = _run(tmp_path, book, "id,x\nP,5\nS,3\n", {"x": "Inputs!B1"}, ["Calc!D8"])
    rows = _rows(res)
    row = rows[("P", "Calc!D8")]
    assert row["status"] == "unverified_path", row
    assert "unexercised_branch@Calc!C2:IF #0 took 'TRUE'" in row["reasons"]
    assert rows[("S", "Calc!D8")]["status"] == "ok"  # the stored path: A2 = 200


def test_argument_cells_are_precedents(make_book: Book, tmp_path: Path) -> None:
    book = make_book(
        "edges",
        {
            "Inputs": {"B1": 1, "B2": 2, "B3": 1, "B4": "A"},
            "Calc": {
                "A1": 10,
                "A2": 20,
                "A3": 30,
                "B1": F("=SUM(OFFSET(A1,Inputs!B1,0,Inputs!B2,Inputs!B3))", 50),
                "B2": F('=INDIRECT(Inputs!B4&"2")', 20),
                "B3": F("=SUM(A1:OFFSET(A1,Inputs!B1,0))", 30),
                "B4": F("=OFFSET(A1,INDEX(Inputs!B1:B2,2),0)", 30),
            },
        },
    )
    graph = open_graph(book, None, tmp_path / "ir")
    assert {CellRef("Inputs", r, 2) for r in (1, 2, 3)} <= graph.precedents(CellRef("Calc", 1, 2))
    assert CellRef("Inputs", 4, 2) in graph.precedents(CellRef("Calc", 2, 2))
    assert CellRef("Inputs", 1, 2) in graph.precedents(CellRef("Calc", 3, 2))
    # An INDEX whose value is an argument reads its whole range.
    assert {CellRef("Inputs", 1, 2), CellRef("Inputs", 2, 2)} <= graph.precedents(
        CellRef("Calc", 4, 2)
    )


def test_keys_moved_by_an_offset_argument_are_input_dependent(
    make_book: Book, tmp_path: Path
) -> None:
    # K1:K3 = OFFSET(Base!A1..A3, In!B1, 0): the keys move with In!B1 only through
    # OFFSET's argument; L1 looks them up approximately.
    book = make_book(
        "keys",
        {
            "In": {"B1": 0},
            "Base": {f"A{i}": 10 * i for i in range(1, 7)},
            "S": {
                "K1": F("=OFFSET(Base!A1,In!B1,0)", 10),
                "K2": F("=OFFSET(Base!A2,In!B1,0)", 20),
                "K3": F("=OFFSET(Base!A3,In!B1,0)", 30),
                "V1": 1,
                "V2": 2,
                "V3": 3,
                "L1": F("=VLOOKUP(25,K1:V3,2,TRUE)", 2),
            },
        },
    )
    graph = open_graph(book, None, tmp_path / "ir")
    found = input_dependent_lookups(graph, {CellRef("In", 1, 2)})
    assert [r.detail for r in found.get(CellRef("S", 1, 12), [])] == [INPUT_KEYS]


# -- (b) built range ends -------------------------------------------------------


def test_range_ends_built_by_offset_and_indirect_are_followed(
    make_book: Book, tmp_path: Path
) -> None:
    # G1:G3 = 10, 20, DATE(1900,1,1)*3 (Excel 3, IronCalc 6). Stored n = 0: E1 = E2 =
    # SUM(G1:G1) = 10, which IronCalc reproduces. n = 2: Excel 33 (IronCalc 36).
    book = make_book(
        "ends",
        {
            "Inputs": {"B2": 0},
            "Tab": {
                "G1": 10,
                "G2": 20,
                "G3": F("=DATE(1900,1,1)*3", 3),
                "E1": F("=SUM(G1:OFFSET(G1,Inputs!$B$2,0))", 10),
                "E2": F('=SUM(G1:INDIRECT("G"&(Inputs!$B$2+1)))', 10),
            },
        },
    )
    res = _run(tmp_path, book, "id,n\nP,2\n", {"n": "Inputs!B2"}, ["Tab!E1", "Tab!E2"])
    rows = _rows(res)
    for out in ("Tab!E1", "Tab!E2"):
        row = rows[("P", out)]
        assert row["value"] == "36.0", (out, row)
        assert row["status"] == "refused", (out, row)
        assert "silently_wrong@Tab!G3" in row["reasons"], (out, row)


def test_a_range_end_that_cannot_be_read_is_refused(make_book: Book, tmp_path: Path) -> None:
    # A range from S!A1 to a cell on another sheet is not a range Excel can build (it
    # gives #VALUE!): calc cannot say which cells it covers, so its probe is unplaced
    # and the outputs behind it are refused. B1 and B2 build such a range through an
    # OFFSET starting on another sheet and through an INDIRECT naming one; B3 builds
    # an ordinary one, which gets its probe.
    book = make_book(
        "unread",
        {
            "S": {
                "A1": 1,
                "A2": 2,
                "B1": F("=SUM(A1:OFFSET(T!A1,1,0))", "#VALUE!"),
                "B2": F('=SUM(A1:INDIRECT("T!A2"))', "#VALUE!"),
                "B3": F("=SUM(A1:OFFSET(A1,1,0))", 3),
            },
            "T": {"A1": 1, "A2": 2},
        },
    )
    graph = open_graph(book, None, tmp_path / "ir")
    for row in (1, 2):
        cell = CellRef("S", row, 2)
        plan = plan_probes(graph, [cell], cap=100)
        assert plan.unplaced.get(cell) == [ProbeKind.dynamic_ref], cell
    cell = CellRef("S", 3, 2)
    plan = plan_probes(graph, [cell], cap=100)
    assert [p.kind for p in plan.probes] == [ProbeKind.dynamic_ref]
    assert plan.probes[0].meta["function"] == "range"


def test_sweep_cells_pinned(make_book: Book, tmp_path: Path) -> None:
    # Cells of the sweep verified by script, pinned here. Stored k = 1, x = 5, n = 1,
    # t = "RngA". Policy k = 2, x = 500, n = 2, t = "RngB".
    #   Data2!A2 = DATE(1900,1,1): Excel 1, IronCalc 2, so it fails the trust run.
    #   Lvl!A2 = OFFSET(Data2!A1, k-1, 0); Out!B4 = OFFSET(Lvl!A1, IF(x>1,1,0), 0):
    #     OFFSET to OFFSET to the failed cell. Excel 1.
    #   Out!I1 = SUM(D1:INDEX(D1:D5, n)) over 10..50: Excel 30 (IronCalc #VALUE!).
    #   Out!C1 = SUM(CHOOSE(k, Calc!A1:A2, Calc!B1:B2)), B2 = IF(x>100, DATE(1900,1,1), 7):
    #     Excel 1 + 1 = 2; B2's TRUE arm is untried.
    #   Out!N1 = SUM(INDIRECT(t)): names RngA = Calc!A1:A2, RngB = Calc!B1:B2. Excel 2.
    book = make_book(
        "pins",
        {
            "Inputs": {"B1": 1, "B2": 5, "B3": 1, "B4": "RngA"},
            "Data2": {"A1": 5, "A2": F("=DATE(1900,1,1)", 1)},
            "Lvl": {"A1": 3, "A2": F("=OFFSET(Data2!$A$1,Inputs!$B$1-1,0)", 5)},
            "Calc": {
                "A1": 1,
                "A2": 2,
                "B1": 1,
                "B2": F("=IF(Inputs!$B$2>100,DATE(1900,1,1),7)", 7),
                "D1": 10,
                "D2": 20,
                "D3": 30,
                "D4": 40,
                "D5": 50,
            },
            "Out": {
                "B4": F("=OFFSET(Lvl!$A$1,IF(Inputs!$B$2>1,1,0),0)", 5),
                "I1": F("=SUM(Calc!D1:INDEX(Calc!D1:D5,Inputs!$B$3))", 10),
                "C1": F("=SUM(CHOOSE(Inputs!$B$1,Calc!A1:A2,Calc!B1:B2))", 3),
                "N1": F("=SUM(INDIRECT(Inputs!$B$4))", 3),
            },
        },
        names={"RngA": "Calc!$A$1:$A$2", "RngB": "Calc!$B$1:$B$2"},
    )
    fields = {"k": "Inputs!B1", "x": "Inputs!B2", "n": "Inputs!B3", "t": "Inputs!B4"}
    outputs = ["Out!B4", "Out!I1", "Out!C1", "Out!N1"]
    rows = _rows(_run(tmp_path, book, "id,k,x,n,t\nP,2,500,2,RngB\n", fields, outputs))
    assert rows[("P", "Out!B4")]["status"] == "refused"
    assert "silently_wrong@Data2!A2" in rows[("P", "Out!B4")]["reasons"]
    assert rows[("P", "Out!I1")]["status"] == "refused"
    assert rows[("P", "Out!C1")]["status"] == "unverified_path"
    assert "unexercised_branch@Calc!B2" in rows[("P", "Out!C1")]["reasons"]
    assert rows[("P", "Out!N1")]["status"] != "ok"


def test_a_landing_region_past_the_cap_is_never_ok(make_book: Book, tmp_path: Path) -> None:
    # Out!B1 = OFFSET(Calc!A1, k-1, 0); Out!B2 = IF(x>1, 1, 2) takes the cap first.
    # Policies land on Calc!A2 (P: its untried arm, Excel 1; Q: the stored arm, 40). For
    # every cap the region's probes may not fit, and a landing they miss is not ok.
    book = make_book(
        "capped",
        {
            "Inputs": {"B1": 1, "B2": 5},
            "Calc": {
                "A1": F("=IF(Inputs!$B$2>1,10,20)", 10),
                "A2": F("=IF(Inputs!$B$2>100,DATE(1900,1,1),40)", 40),
            },
            "Out": {
                "B1": F("=OFFSET(Calc!$A$1,Inputs!$B$1-1,0)", 10),
                "B2": F("=IF(Inputs!$B$2>1,1,2)", 1),
            },
        },
    )
    fields = {"k": "Inputs!B1", "x": "Inputs!B2"}
    for cap in (1, 2, 3, 4, 5):
        (tmp_path / "p.csv").write_text("id,k,x\nP,2,500\nQ,2,5\n", encoding="utf-8")
        (tmp_path / "m.json").write_text(json.dumps({"fields": fields}), encoding="utf-8")
        res = calc(
            CalcRequest(
                workbook=book,
                policies=tmp_path / "p.csv",
                input_map=tmp_path / "m.json",
                out_dir=tmp_path / f"c{cap}",
                outputs=["Out!B1", "Out!B2"],
                probe_cap=cap,
            )
        )
        rows = _rows(res)
        assert rows[("P", "Out!B1")]["status"] != "ok", cap
        report = json.loads(res.trust_report.read_text(encoding="utf-8"))
        if report["probes_over_cap"] == 0:
            continue
        # Something did not fit: Q's landing is not vouched for either.
        assert rows[("Q", "Out!B1")]["status"] != "ok", (cap, rows[("Q", "Out!B1")])


# -- (c) landing cells ----------------------------------------------------------


def _landing_book(make_book: Book) -> Path:
    # Stored k = 1, x = 5, r = 1.2.
    #   Calc!A1 = IF(x>1, 10, 20) = 10; A2 = IF(x>100, DATE(1900,1,1), 40) = 40;
    #   A3 = ROUND(r, 0) = 1; C2 = IF(x>100, DATE(1900,1,1), 7) = 7; D1 = 3;
    #   D2 = C2*2 = 14.
    #   Out!B1 = OFFSET(Calc!A1, k-1, 0)        -> A1 = 10
    #   Out!B2 = INDIRECT("Calc!A"&k)            -> A1 = 10 (a literal sheet)
    #   Out!B3 = OFFSET(Calc!A2, k-1, 0)        -> A2 = 40
    #   Out!B4 = OFFSET(Calc!D1, k-1, 0)        -> D1 = 3
    #   Calc!E1 = INDIRECT("A"&k)                -> A1 = 10 (no sheet in the text)
    #   Calc!E3 = INDIRECT("A"&(k+2))            -> A3 = 1
    #   Calc!F1 = INDIRECT("D"&k)                -> D1 = 3
    return make_book(
        "landing",
        {
            "Inputs": {"B1": 1, "B2": 5, "B3": 1.2},
            "Calc": {
                "A1": F("=IF(Inputs!$B$2>1,10,20)", 10),
                "A2": F("=IF(Inputs!$B$2>100,DATE(1900,1,1),40)", 40),
                "A3": F("=ROUND(Inputs!$B$3,0)", 1),
                "C2": F("=IF(Inputs!$B$2>100,DATE(1900,1,1),7)", 7),
                "D1": 3,
                "D2": F("=Calc!C2*2", 14),
                "E1": F('=INDIRECT("A"&Inputs!$B$1)', 10),
                "E3": F('=INDIRECT("A"&(Inputs!$B$1+2))', 1),
                "F1": F('=INDIRECT("D"&Inputs!$B$1)', 3),
            },
            "Out": {
                "B1": F("=OFFSET(Calc!$A$1,Inputs!$B$1-1,0)", 10),
                "B2": F('=INDIRECT("Calc!A"&Inputs!$B$1)', 10),
                "B3": F("=OFFSET(Calc!$A$2,Inputs!$B$1-1,0)", 40),
                "B4": F("=OFFSET(Calc!$D$1,Inputs!$B$1-1,0)", 3),
            },
        },
    )


_FIELDS = {"k": "Inputs!B1", "x": "Inputs!B2", "r": "Inputs!B3"}


def test_landing_cells_are_probed(make_book: Book, tmp_path: Path) -> None:
    # Policy k = 2, x = 500, r = 2.5. Excel: B1 = B2 = A2 = DATE(1900,1,1) = 1, the TRUE
    # arm the trust run never took; B3 = A3 = ROUND(2.5,0) = 3, an exact half; B4 = D2 =
    # C2*2 = 2, whose precedent took the untried arm.
    book = _landing_book(make_book)
    outputs = ["Out!B1", "Out!B2", "Out!B3", "Out!B4"]
    res = _run(tmp_path, book, "id,k,x,r\nP,2,500,2.5\n", _FIELDS, outputs)
    rows = _rows(res)
    want = {
        "Out!B1": ("unverified_path", "unexercised_branch@Calc!A2:IF #0 took 'TRUE'"),
        "Out!B2": ("unverified_path", "unexercised_branch@Calc!A2:IF #0 took 'TRUE'"),
        "Out!B3": ("uncertain", "near_rounding_flip@Calc!A3"),
        "Out!B4": ("unverified_path", "unexercised_branch@Calc!C2:IF #0 took 'TRUE'"),
    }
    for out, (status, reason) in want.items():
        row = rows[("P", out)]
        assert row["status"] == status, (out, row)
        assert reason in row["reasons"], (out, row)


def test_an_indirect_naming_no_sheet_lands_unchecked(make_book: Book, tmp_path: Path) -> None:
    # INDIRECT text that names no sheet can land anywhere: the cells it lands on are
    # not probed (ruling D1), so a landing on formulas is never ok. Each such cell is
    # not_checked_path, and its probes count as unplaced: the IF's comparison gives
    # not_checked_value (uncertain), as an unplaced comparison probe does. Policy k =
    # 2, x = 500: E1 = A2 = 1, F1 = D2 = C2*2 = 2. k = 1, x = 5 lands where the stored
    # inputs do: A1 (an IF the plan did not probe) and D1 (a constant).
    book = _landing_book(make_book)
    outputs = ["Calc!E1", "Calc!F1"]
    res = _run(tmp_path, book, "id,k,x,r\nP,2,500,2.5\nQ,1,5,1.2\n", _FIELDS, outputs)
    rows = _rows(res)
    p_e1 = rows[("P", "Calc!E1")]
    assert p_e1["value"] == "2.0"  # IronCalc's DATE(1900,1,1); Excel 1
    assert p_e1["status"] == "uncertain", p_e1
    assert "not_checked_path@Calc!A2:reached through a dynamic reference" in p_e1["reasons"]
    assert "not_checked_value@Calc!A2:comparison not checked" in p_e1["reasons"]
    p_f1 = rows[("P", "Calc!F1")]
    assert p_f1["status"] == "uncertain", p_f1
    assert "not_checked_path@Calc!C2:reached through a dynamic reference" in p_f1["reasons"]
    assert "not_checked_path@Calc!D2:reached through a dynamic reference" in p_f1["reasons"]
    assert rows[("Q", "Calc!E1")]["status"] == "uncertain"
    assert rows[("Q", "Calc!F1")]["status"] == "ok"  # a constant: nothing to check


def test_an_unprobed_landing_on_rounding_is_uncertain(make_book: Book, tmp_path: Path) -> None:
    # Calc!E3 = INDIRECT("A"&(k+2)): k = 1 lands on A3 = ROUND(r,0). The rounding is not
    # probed, so the value is uncertain, as an unplaced rounding probe is.
    book = _landing_book(make_book)
    res = _run(tmp_path, book, "id,k,x,r\nP,1,5,2.5\n", _FIELDS, ["Calc!E3"])
    row = _rows(res)[("P", "Calc!E3")]
    assert row["value"] == "3.0"
    assert row["status"] == "uncertain", row
    assert "not_checked_value@Calc!A3:rounding not checked" in row["reasons"]


def test_fewer_outputs_are_never_less_conservative(make_book: Book, tmp_path: Path) -> None:
    # Each output alone, then all of them with every cell they can land on: asking for
    # fewer outputs may make a status more conservative, never less (guide).
    book = _landing_book(make_book)
    policies = "id,k,x,r\nP,2,500,2.5\nQ,1,5,2.5\nS,1,5,1.2\nT,3,500,1.2\n"
    alone = ["Out!B1", "Out!B2", "Out!B3", "Out!B4", "Calc!E1", "Calc!E3", "Calc!F1"]
    every = [*alone, "Calc!A1:A3", "Calc!C2", "Calc!D2"]
    full = _rows(_run(tmp_path, book, policies, _FIELDS, every, tag="all"))
    for out in alone:
        rows = _rows(_run(tmp_path, book, policies, _FIELDS, [out], tag=out.replace("!", "_")))
        for pid in ("P", "Q", "S", "T"):
            few, many = rows[(pid, out)], full[(pid, out)]
            assert few["value"] == many["value"], (pid, out)
            assert _RANK[few["status"]] >= _RANK[many["status"]], (pid, out, few, many)


def test_a_check_set_exercises_the_cells_it_landed_on(make_book: Book, tmp_path: Path) -> None:
    # Calc!A2 = IF(x>100, 77, 40). Out!B1 = OFFSET(Calc!A1, k-1, 0). A results sheet
    # stores Out!B1 for input set C (k = 2, x = 500): 77, which the engine reproduces.
    # Policy P (k = 2, x = 700) takes the same arm: 77, checked by set C.
    book = make_book(
        "check",
        {
            "Inputs": {"B1": 1, "B2": 5},
            "Calc": {
                "A1": F("=IF(Inputs!$B$2>1,10,20)", 10),
                "A2": F("=IF(Inputs!$B$2>100,77,40)", 40),
            },
            "Out": {"B1": F("=OFFSET(Calc!$A$1,Inputs!$B$1-1,0)", 10)},
            "Sets": {"A2": "C", "B2": 2, "C2": 500},
            "Results": {"A2": "C", "B2": 77},
        },
    )
    fields = {"k": "Inputs!B1", "x": "Inputs!B2"}
    check = {
        "results": {
            "sheet": "Results",
            "id_column": "A",
            "first_row": 2,
            "outputs": {"Out!B1": "B"},
        },
        "inputs": {
            "table": {
                "sheet": "Sets",
                "id_column": "A",
                "first_row": 2,
                "fields": {"k": "B", "x": "C"},
            }
        },
    }
    without = _rows(_run(tmp_path, book, "id,k,x\nP,2,700\n", fields, ["Out!B1"], tag="a"))
    assert without[("P", "Out!B1")]["status"] == "unverified_path"
    checked = _rows(
        _run(tmp_path, book, "id,k,x\nP,2,700\n", fields, ["Out!B1"], check=check, tag="b")
    )
    row = checked[("P", "Out!B1")]
    assert row["value"] == "77.0"
    assert row["status"] == "ok", row


def _base_credit_book(make_book: Book, stored_k: int) -> Path:
    # Calc!A1 = IF(x>100, DATE(1900,1,1), 40); A2 = 50. Out!B1 = OFFSET(Calc!A1, k-1, 0).
    # Stored x = 5. Results hold Out!B1 for input set C (k = 2, x = 500): A2 = 50, which
    # the engine reproduces; A1 takes its TRUE arm in that set but does not feed Out!B1.
    stored = {1: 40, 2: 50}[stored_k]
    return make_book(
        f"credit{stored_k}",
        {
            "Inputs": {"B1": stored_k, "B2": 5},
            "Calc": {"A1": F("=IF(Inputs!$B$2>100,DATE(1900,1,1),40)", 40), "A2": 50},
            "Out": {"B1": F("=OFFSET(Calc!$A$1,Inputs!$B$1-1,0)", stored)},
            "Sets": {"A2": "C", "B2": 2, "C2": 500},
            "Results": {"A2": "C", "B2": 50},
        },
    )


_SET_C = {
    "results": {"sheet": "Results", "id_column": "A", "first_row": 2, "outputs": {"Out!B1": "B"}},
    "inputs": {
        "table": {"sheet": "Sets", "id_column": "A", "first_row": 2, "fields": {"k": "B", "x": "C"}}
    },
}


@pytest.mark.parametrize("stored_k", [1, 2], ids=["stored-landing", "base-only"])
def test_a_check_set_vouches_only_for_where_it_landed(
    make_book: Book, tmp_path: Path, stored_k: int
) -> None:
    # Policy P (k = 1, x = 500) lands on A1 and takes its TRUE arm: Excel 1 (IronCalc 2).
    # Neither the trust run (x = 5, FALSE) nor set C (which landed on A2) checked that
    # arm feeding Out!B1, whether A1 is the stored landing cell (k = 1) or only the cell
    # the OFFSET starts from (k = 2).
    book = _base_credit_book(make_book, stored_k)
    fields = {"k": "Inputs!B1", "x": "Inputs!B2"}
    res = _run(tmp_path, book, "id,k,x\nP,1,500\n", fields, ["Out!B1"], check=_SET_C)
    row = _rows(res)[("P", "Out!B1")]
    assert row["value"] == "2.0"
    assert row["status"] == "unverified_path", row
    assert "unexercised_branch@Calc!A1:IF #0 took 'TRUE'" in row["reasons"]


def test_paths_list_only_where_the_policy_landed(make_book: Book, tmp_path: Path) -> None:
    # Stored k = 1 (the IR's landing is A1). Policy Q lands on A2, a constant: no branch
    # fed its output, so paths.csv has no row for A1.
    book = _base_credit_book(make_book, 1)
    fields = {"k": "Inputs!B1", "x": "Inputs!B2"}
    res = _run(tmp_path, book, "id,k,x\nQ,2,5\n", fields, ["Out!B1"])
    with res.paths_csv.open(encoding="utf-8", newline="") as fh:
        cells = {r["cell"] for r in csv.DictReader(fh)}
    assert "Calc!A1" not in cells


def test_unprobed_landings_are_walked_once_per_cell(make_book: Book, tmp_path: Path) -> None:
    # A recursive projection P!B2:D41 (each row reads the row above), summed through an
    # INDIRECT naming no sheet, so none of it is probed. Landing on every row of D must
    # look at each upstream cell once, and plan each formula's probes once per run, not
    # once per landed cell (which grows as rows squared).
    from xl_marinade.calc.probes import probe_kinds
    from xl_marinade.calc.status import LandingCheck

    rows = 40
    sheet: dict[str, object] = {}
    for r in range(2, rows + 2):
        for col, left in (("B", None), ("C", "B"), ("D", "C")):
            prev = f"{col}{r - 1}" if r > 2 else "1"
            extra = f"+{left}{r}" if left else ""
            sheet[f"{col}{r}"] = F(f"={prev}*2+IF(A1>0,1,0){extra}", 0)
    sheet["A1"] = 1
    book = make_book("chain", {"P": sheet})
    graph = open_graph(book, None, tmp_path / "ir")
    planned: list[CellRef] = []

    def kinds(cell: CellRef) -> list[ProbeKind]:
        planned.append(cell)
        return probe_kinds(graph, cell)

    check = LandingCheck(graph, frozenset(), kinds)
    landed = {CellRef("P", 1, 9): [CellRef("P", r, 4) for r in range(2, rows + 2)]}
    first = check.roots(landed)
    assert len(planned) == len(set(planned)) == 3 * rows
    assert check.visits <= 3 * rows + 1  # each formula cell once, plus A1
    again = check.roots(landed)
    assert again == first
    assert len(planned) == 3 * rows  # memoised across policies
    for r in range(2, rows + 2):
        reasons = first[CellRef("P", r, 2)]
        assert {x.kind.value for x in reasons} == {"not_checked_path", "not_checked_value"}


def test_probes_past_the_cap_are_reported_with_the_cap_needed(
    make_book: Book, tmp_path: Path
) -> None:
    # The OFFSET in Out!B1 brings every formula of Calc into the plan; a cap of 2 cannot
    # hold them. The run says how many probes went past the cap and which cap holds them
    # all, in trust_report.json and on stderr; rerunning with that cap places them all.
    import re

    from typer.testing import CliRunner

    from xl_marinade.cli.main import app

    book = _landing_book(make_book)
    (tmp_path / "p.csv").write_text("id,k,x,r\nP,2,500,2.5\n", encoding="utf-8")
    (tmp_path / "m.json").write_text(json.dumps({"fields": _FIELDS}), encoding="utf-8")

    def cli(cap: int, out: str) -> tuple[str, dict[str, object]]:
        args = ["calc", str(book), "--policies", str(tmp_path / "p.csv"), "--map"]
        args += [str(tmp_path / "m.json"), "--outputs", "Out!B1", "-o", str(tmp_path / out)]
        result = CliRunner().invoke(app, [*args, "--probe-cap", str(cap)])
        assert result.exit_code == 0, result.output
        report = json.loads((tmp_path / out / "trust_report.json").read_text(encoding="utf-8"))
        return result.stderr, report

    stderr, report = cli(2, "a")
    over = report["probes_over_cap"]
    assert isinstance(over, int) and over > 0
    assert isinstance(report["probes_unplaced"], int) and report["probes_unplaced"] >= over
    m = re.search(r"rerun with --probe-cap (\d+)", stderr)
    assert m is not None, stderr
    assert int(m.group(1)) == 2 + over
    stderr, report = cli(2 + over, "b")
    assert report["probes_over_cap"] == 0
    assert "--probe-cap" not in stderr


def test_sheet_scoped_and_relative_names_stay_out_of_dynamic_paths(
    make_book: Book, tmp_path: Path
) -> None:
    # A sheet's own name and a name with a relative reference cannot be read off the
    # formula's own sheet (the probe sheet would read another definition, or other
    # cells), so no dynamic-reference path may use them as fixed cells:
    #   Calc!B1 = OFFSET(Local, k, 0), Local = Calc's own name for Data!$A$1: the OFFSET
    #     reaches no landing region (its probe is unplaced, so it is refused);
    #   Calc!B2 = OFFSET(Data!$A$1, Rel, 0), Rel = Calc!XFD2 written relative (one
    #     column left of the cell using it): the stored text's cell is not among what
    #     the OFFSET reads (the formula itself is refused for the relative name);
    #   Calc!B3 = TEXTJOIN(",",TRUE,Local2), Local2 = Calc's own name for Data!$A$1:$A$2:
    #     the join is not read cell by cell (refused per policy as not checked).
    from xl_marinade.calc.probes import reach_sheets

    book = make_book(
        "names_dyn",
        {
            "Data": {"A1": 1, "A2": 2, "A3": 3},
            "Calc": {
                "A2": 1,
                "B1": F("=OFFSET(Local,Inputs!B1,0)", 2),
                "B2": F("=OFFSET(Data!$A$1,Rel,0)", 2),
                "B3": F('=_xlfn.TEXTJOIN(",",TRUE,Local2)', "1,2"),
                "XFD2": 7,
            },
            "Inputs": {"B1": 1},
        },
        names={"Rel": "Calc!XFD2"},
        local_names={("Calc", "Local"): "Data!$A$1", ("Calc", "Local2"): "Data!$A$1:$A$2"},
    )
    graph = open_graph(book, None, tmp_path / "ir")
    assert reach_sheets(graph, CellRef("Calc", 1, 2)) == set()
    assert CellRef("Calc", 2, 16384) not in graph.read_precedents(CellRef("Calc", 2, 2))
    plan = plan_probes(graph, [CellRef("Calc", 3, 2)], cap=100)
    assert [(p.kind, p.meta.get("check")) for p in plan.probes] == [
        (ProbeKind.number_text, "always")
    ]
