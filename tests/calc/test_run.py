"""Tests for the ``calc()`` run: settings, trust run, probes, policy loop, output files."""

from __future__ import annotations

import csv
import gc
import json
import tempfile
from collections.abc import Callable
from pathlib import Path

import pytest

pytest.importorskip("ironcalc")

from tests.calc.conftest import F, add_external_link  # noqa: E402
from xl_marinade.calc import CalcRequest, CalcResult, calc  # noqa: E402
from xl_marinade.calc import run as run_module  # noqa: E402
from xl_marinade.calc.output import OutputValue, PathRecord  # noqa: E402
from xl_marinade.calc.probes import ProbeReadings  # noqa: E402
from xl_marinade.calc.run import coerce_input  # noqa: E402
from xl_marinade.errors import CalcInputError, WorkbookRefused  # noqa: E402

Book = Callable[..., Path]

SHEETS: dict[str, dict[str, object]] = {
    "Inputs": {"A1": "age", "B1": 40, "A2": "sex", "B2": "M"},
    "Calc": {
        "B1": F("=IF(Inputs!B1>60,Inputs!B1*2,Inputs!B1)", 40),
        "B2": F("=ROUND(Calc!B1/8,2)", 5),
        "B3": F("=Inputs!B1+RAND()*0", 40),
        "B4": F('=IF(Inputs!B2="M",1,2)', 1),
    },
}
MAP = {"id_field": "id", "fields": {"age": "Inputs!B1", "sex": "Inputs!B2"}}


def _book(make_book: Book, name: str = "book", **kw: object) -> Path:
    return make_book(name, SHEETS, validations=[("Inputs!B2", ["M", "F"])], **kw)


def _write(tmp: Path, name: str, text: str) -> Path:
    path = tmp / name
    path.write_text(text, encoding="utf-8")
    return path


def _req(
    tmp: Path,
    book: Path,
    policies: str,
    *,
    outputs: list[str] | None = None,
    imap: dict[str, object] | None = None,
    **kw: object,
) -> CalcRequest:
    return CalcRequest(
        workbook=book,
        policies=_write(tmp, "policies.csv", policies),
        input_map=_write(tmp, "map.json", json.dumps(imap or MAP)),
        out_dir=tmp / "out",
        outputs=outputs if outputs is not None else ["Calc!B1:B4"],
        **kw,  # type: ignore[arg-type]
    )


def _results(res: CalcResult) -> dict[tuple[str, str], dict[str, str]]:
    with res.results_csv.open(encoding="utf-8", newline="") as fh:
        return {(r["policy_id"], r["output"]): r for r in csv.DictReader(fh)}


def test_branch_rounding_and_refused_pattern(make_book: Book, tmp_path: Path) -> None:
    book = _book(make_book)
    res = calc(_req(tmp_path, book, "id,age,sex\n1,50,M\n2,70,M\n3,41,M\n"))
    got = _results(res)

    ok = got[("1", "Calc!B1")]
    assert (ok["value"], ok["status"], ok["reasons"]) == ("50.0", "ok", "")

    unverified = got[("2", "Calc!B1")]
    assert unverified["value"] == "140.0"
    assert unverified["status"] == "unverified_path"
    assert unverified["reasons"].startswith("unexercised_branch@Calc!B1:")

    near = got[("3", "Calc!B2")]
    assert near["value"] == "5.13"
    assert near["status"] == "uncertain"
    assert "near_rounding_flip@Calc!B2" in near["reasons"]

    for pid in ("1", "2", "3"):
        rand = got[(pid, "Calc!B3")]
        assert rand["status"] == "refused"
        assert "refused_pattern@Calc!B3" in rand["reasons"]

    # policies in CSV order, outputs in request order
    with res.results_csv.open(encoding="utf-8", newline="") as fh:
        order = [(r["policy_id"], r["output"]) for r in csv.DictReader(fh)]
    assert order[:4] == [("1", "Calc!B1"), ("1", "Calc!B2"), ("1", "Calc!B3"), ("1", "Calc!B4")]
    assert order[4][0] == "2"


def test_trust_report_and_run_json(make_book: Book, tmp_path: Path) -> None:
    book = _book(make_book)
    res = calc(_req(tmp_path, book, "id,age,sex\n1,50,M\n"))
    report = json.loads(res.trust_report.read_text(encoding="utf-8"))
    for key in (
        "formula_cells",
        "reproduced",
        "share",
        "failures",
        "refused_patterns",
        "settings",
        "held_external_links",
    ):
        assert key in report
    assert report["formula_cells"] == 4
    assert report["reproduced"] == 3 and report["share"] == 1.0
    assert report["failures"] == []
    assert [p["cell"] for p in report["refused_patterns"]] == ["Calc!B3"]
    run = json.loads(res.run_json.read_text(encoding="utf-8"))
    for key in (
        "ironcalc_version",
        "marinade_version",
        "workbook_sha256",
        "tolerance",
        "as_of",
        "probe_cap",
        "probes_unplaced",
        "variables",
    ):
        assert key in run
    assert run["ironcalc_version"] == "0.9.1"
    assert run["tolerance"] == 1e-9 and run["as_of"] is None and run["probe_cap"] == 50_000
    assert sorted(p.name for p in (tmp_path / "out").iterdir()) == [
        "paths.csv",
        "results.csv",
        "run.json",
        "trust_report.json",
    ]


def test_trust_failure_is_a_root_silently_wrong_first(make_book: Book, tmp_path: Path) -> None:
    sheets = {
        "Inputs": {"B1": 40},
        "Calc": {"B1": F("=Inputs!B1*2", 81), "B2": F("=Calc!B1+1", 82), "B3": F("=1", 1)},
    }
    book = make_book("wrong", sheets)
    imap = {"id_field": "id", "fields": {"age": "Inputs!B1"}}
    res = calc(_req(tmp_path, book, "id,age\n1,50\n", outputs=["Calc!B2", "Calc!B3"], imap=imap))
    got = _results(res)
    assert got[("1", "Calc!B2")]["status"] == "refused"
    assert "silently_wrong@Calc!B1" in got[("1", "Calc!B2")]["reasons"]
    assert got[("1", "Calc!B3")]["status"] == "ok"
    report = json.loads(res.trust_report.read_text(encoding="utf-8"))
    assert report["failures"][0]["kind"] == "silently_wrong"
    assert report["failures"][0]["cell"] == "Calc!B1"


def test_date1904_refuses_whole_workbook(make_book: Book, tmp_path: Path) -> None:
    book = _book(make_book, "b1904", date1904=True)
    with pytest.raises(WorkbookRefused) as info:
        calc(_req(tmp_path, book, "id,age,sex\n1,50,M\n"))
    assert info.value.reasons == ["1904 date system"]
    out = tmp_path / "out"
    assert sorted(p.name for p in out.iterdir()) == ["run.json", "trust_report.json"]
    report = json.loads((out / "trust_report.json").read_text(encoding="utf-8"))
    assert report["settings"]["date1904"] is True


def test_mapping_a_formula_cell_is_an_input_error(make_book: Book, tmp_path: Path) -> None:
    book = _book(make_book)
    imap = {"id_field": "id", "fields": {"age": "Calc!B1"}}
    with pytest.raises(CalcInputError, match="formula"):
        calc(_req(tmp_path, book, "id,age\n1,50\n", imap=imap))


def test_value_outside_dropdown_refuses_the_policy(make_book: Book, tmp_path: Path) -> None:
    book = _book(make_book)
    res = calc(_req(tmp_path, book, "id,age,sex\n1,50,X\n2,50,F\n"))
    got = _results(res)
    for out in ("Calc!B1", "Calc!B2", "Calc!B3", "Calc!B4"):
        row = got[("1", out)]
        assert row["status"] == "refused"
        assert "input_outside_allowed_values@Inputs!B2" in row["reasons"]
    assert got[("2", "Calc!B1")]["status"] == "ok"
    assert "input_outside_allowed_values" not in got[("2", "Calc!B4")]["reasons"]


def test_selector_and_inputs_are_reset_before_each_policy(make_book: Book, tmp_path: Path) -> None:
    sheets: dict[str, dict[str, object]] = {
        "Sel": {"A1": 1},
        "T": {"A1": 40, "A2": 70},
        "Inputs": {"B1": 5, "B2": 0},
        "Calc": {"B1": F("=INDEX(T!A1:A2,Sel!A1)+Inputs!B1+Inputs!B2", 45)},
        "Results": {"A2": 1, "B2": 45, "A3": 2, "B3": 75},
    }
    book = make_book("reset", sheets)
    check = {
        "results": {
            "sheet": "Results",
            "id_column": "A",
            "first_row": 2,
            "outputs": {"Calc!B1": "B"},
        },
        "inputs": {"selector": "Sel!A1"},
    }
    imap = {"id_field": "id", "fields": {"x": "Inputs!B1", "y": "Inputs!B2"}}
    res = calc(
        _req(
            tmp_path,
            book,
            "id,x,y\n1,10,100\n2,10,0\n",
            imap=imap,
            outputs=["Calc!B1"],
            check_against=_write(tmp_path, "check.json", json.dumps(check)),
        )
    )
    got = _results(res)
    # the check left the selector on input set 2 (70); each policy starts from stored (40)
    assert got[("1", "Calc!B1")]["value"] == "150.0"
    assert got[("2", "Calc!B1")]["value"] == "50.0"


def test_uncoercible_value_outside_dropdown_refuses_only_that_policy(
    make_book: Book, tmp_path: Path
) -> None:
    book = make_book(
        "numdd",
        {"Inputs": {"B1": 40}, "Calc": {"B1": F("=Inputs!B1*2", 80)}},
        validations=[("Inputs!B1", ["40", "50"])],
    )
    imap = {"id_field": "id", "fields": {"age": "Inputs!B1"}}
    res = calc(_req(tmp_path, book, "id,age\n1,X\n2,50\n", imap=imap, outputs=["Calc!B1"]))
    got = _results(res)
    assert got[("1", "Calc!B1")]["status"] == "refused"
    assert "input_outside_allowed_values@Inputs!B1" in got[("1", "Calc!B1")]["reasons"]
    assert (got[("2", "Calc!B1")]["value"], got[("2", "Calc!B1")]["status"]) == ("100.0", "ok")


def test_uncoercible_allowed_value_is_an_input_error(make_book: Book, tmp_path: Path) -> None:
    book = make_book("numdd2", {"Inputs": {"B1": 40}, "Calc": {"B1": F("=Inputs!B1*2", 80)}})
    imap = {"id_field": "id", "fields": {"age": "Inputs!B1"}}
    with pytest.raises(CalcInputError, match="not a number"):
        calc(_req(tmp_path, book, "id,age\n1,X\n", imap=imap, outputs=["Calc!B1"]))


def _check_book(make_book: Book, name: str, stored_140: float) -> Path:
    sheets = dict(SHEETS)
    sheets["Results"] = {"A1": "id", "B1": "b1", "A2": 1, "B2": 40, "A3": 2, "B3": stored_140}
    sheets["Table"] = {"A1": "id", "B1": "age", "A2": 1, "B2": 40, "A3": 2, "B3": 70}
    return make_book(name, sheets, validations=[("Inputs!B2", ["M", "F"])])


CHECK = {
    "results": {"sheet": "Results", "id_column": "A", "first_row": 2, "outputs": {"Calc!B1": "B"}},
    "inputs": {
        "table": {"sheet": "Table", "id_column": "A", "first_row": 2, "fields": {"age": "B"}}
    },
}


def test_check_against_exercises_the_other_branch(make_book: Book, tmp_path: Path) -> None:
    book = _check_book(make_book, "chk", 140)
    check = _write(tmp_path, "check.json", json.dumps(CHECK))
    res = calc(_req(tmp_path, book, "id,age,sex\n9,75,M\n", check_against=check))
    row = _results(res)[("9", "Calc!B1")]
    assert (row["value"], row["status"]) == ("150.0", "ok")


def test_check_against_mismatch_refuses_the_output(make_book: Book, tmp_path: Path) -> None:
    book = _check_book(make_book, "chkbad", 141)
    check = _write(tmp_path, "check.json", json.dumps(CHECK))
    # Age 75 takes the branch the mismatching input set (age 70) took.
    res = calc(_req(tmp_path, book, "id,age,sex\n9,75,M\n", check_against=check))
    got = _results(res)
    assert got[("9", "Calc!B1")]["status"] == "refused"
    assert "check_against_mismatch@Calc!B1" in got[("9", "Calc!B1")]["reasons"]
    assert got[("9", "Calc!B2")]["status"] == "refused"


def test_check_against_selector(make_book: Book, tmp_path: Path) -> None:
    sheets: dict[str, dict[str, object]] = {
        "Sel": {"A1": 1},
        "T": {"A1": 40, "A2": 70},
        "Calc": {
            "B1": F("=INDEX(T!A1:A2,Sel!A1)", 40),
            "B2": F("=IF(Calc!B1>60,1,2)", 2),
        },
        "Results": {"A2": 1, "B2": 2, "A3": 2, "B3": 1},
    }
    book = make_book("sel", sheets)
    check = {
        "results": {
            "sheet": "Results",
            "id_column": "A",
            "first_row": 2,
            "outputs": {"Calc!B2": "B"},
        },
        "inputs": {"selector": "Sel!A1"},
    }
    imap = {"id_field": "id", "fields": {"sel": "Sel!A1"}}
    res = calc(
        _req(
            tmp_path,
            book,
            "id,sel\n5,2\n",
            imap=imap,
            outputs=["Calc!B2"],
            check_against=_write(tmp_path, "check.json", json.dumps(check)),
        )
    )
    row = _results(res)[("5", "Calc!B2")]
    assert (row["value"], row["status"]) == ("1.0", "ok")


def test_variables_write_rows_with_statuses(make_book: Book, tmp_path: Path) -> None:
    book = _book(make_book)
    variables = _write(
        tmp_path,
        "vars.json",
        json.dumps({"calc": {"range": "Calc!B1:B2", "period_start": 0}, "r": {"cell": "Calc!B3"}}),
    )
    res = calc(_req(tmp_path, book, "id,age,sex\n1,70,M\n", outputs=[], variables=variables))
    assert res.rows_csv is not None
    with res.rows_csv.open(encoding="utf-8", newline="") as fh:
        rows = list(csv.DictReader(fh))
    assert [(r["variable"], r["period"], r["value"], r["status"]) for r in rows] == [
        ("calc", "0", "140.0", "unverified_path"),
        ("calc", "1", "17.5", "unverified_path"),
        ("r", "", "70.0", "refused"),
    ]
    run = json.loads(res.run_json.read_text(encoding="utf-8"))
    assert run["variables"]["calc"]["range"] == "Calc!B1:B2"
    assert sorted(_results(res)) == [("1", "Calc!B1"), ("1", "Calc!B2"), ("1", "Calc!B3")]


def test_paths_csv_records_each_branch(make_book: Book, tmp_path: Path) -> None:
    book = _book(make_book)
    res = calc(_req(tmp_path, book, "id,age,sex\n1,50,M\n2,70,M\n", outputs=["Calc!B1"]))
    with res.paths_csv.open(encoding="utf-8", newline="") as fh:
        rows = list(csv.DictReader(fh))
    assert [(r["policy_id"], r["cell"], r["kind"], r["exercised"]) for r in rows] == [
        ("1", "Calc!B1", "branch", "true"),
        ("2", "Calc!B1", "branch", "false"),
    ]
    assert rows[0]["outcome"] != rows[1]["outcome"]


def test_reruns_are_byte_identical_and_quiet(
    make_book: Book, tmp_path: Path, capfd: pytest.CaptureFixture[str]
) -> None:
    book = _book(make_book)
    policies = "id,age,sex\n1,50,M\n2,70,F\n"
    (tmp_path / "a").mkdir()
    (tmp_path / "b").mkdir()
    first = calc(_req(tmp_path / "a", book, policies))
    second = calc(_req(tmp_path / "b", book, policies))
    captured = capfd.readouterr()
    assert captured.out == ""
    assert "Unexpected type" not in captured.err
    for name in ("results.csv", "paths.csv", "trust_report.json", "run.json"):
        a = (first.results_csv.parent / name).read_bytes()
        b = (second.results_csv.parent / name).read_bytes()
        assert a == b, name


def test_temp_dir_deleted_and_nothing_outside_out_dir(
    make_book: Book, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    book = _book(make_book)
    temp = tmp_path / "tmp"
    temp.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(temp))
    work = tmp_path / "w"
    work.mkdir()
    before = set(tmp_path.rglob("*"))
    calc(_req(work, book, "id,age,sex\n1,50,M\n"))
    assert list(temp.iterdir()) == []
    new = set(tmp_path.rglob("*")) - before
    inputs = {work / "policies.csv", work / "map.json"}
    assert {p for p in new if p not in inputs and work / "out" not in p.parents} == {work / "out"}


def test_held_external_link_stays_ok_with_reason(make_book: Book, tmp_path: Path) -> None:
    base = make_book(
        "extbase", {"Inputs": {"B1": 1}, "Calc": {"B1": F("=[1]Other!B2+Inputs!B1", 43)}}
    )
    book = add_external_link(base, tmp_path / "ext.xlsx", "Other", {"B2": "42"})
    imap = {"id_field": "id", "fields": {"x": "Inputs!B1"}}
    res = calc(_req(tmp_path, book, "id,x\n1,2\n", outputs=["Calc!B1"], imap=imap))
    row = _results(res)[("1", "Calc!B1")]
    assert (row["value"], row["status"]) == ("44.0", "ok")
    assert row["reasons"] == "external_link_held@Calc!B1:[1]Other!B2"
    report = json.loads(res.trust_report.read_text(encoding="utf-8"))
    assert report["held_external_links"] == [{"cell": "Calc!B1", "refs": ["[1]Other!B2"]}]


def test_mismatching_check_set_does_not_count_as_exercised(make_book: Book, tmp_path: Path) -> None:
    branch = "=IF(Inputs!$B$1>60,Inputs!$B$1*2,Inputs!$B$1)"
    sheets: dict[str, dict[str, object]] = {
        "Inputs": {"B1": 40},
        "Calc": {"B1": F(branch, 40), "C1": F(branch, 40)},
        "Results": {"A2": 1, "B2": 40, "A3": 2, "B3": 141},
        "Table": {"A2": 1, "B2": 40, "A3": 2, "B3": 70},
    }
    book = make_book("chkex", sheets)
    check = {
        "results": {
            "sheet": "Results",
            "id_column": "A",
            "first_row": 2,
            "outputs": {"Calc!B1": "B"},
        },
        "inputs": {
            "table": {"sheet": "Table", "id_column": "A", "first_row": 2, "fields": {"age": "B"}}
        },
    }
    imap = {"id_field": "id", "fields": {"age": "Inputs!B1"}}
    res = calc(
        _req(
            tmp_path,
            book,
            "id,age\n9,75\n",
            imap=imap,
            outputs=["Calc!C1"],
            check_against=_write(tmp_path, "check.json", json.dumps(check)),
        )
    )
    row = _results(res)[("9", "Calc!C1")]
    assert row["status"] == "unverified_path"
    assert "unexercised_branch@Calc!C1" in row["reasons"]


@pytest.mark.parametrize(
    ("formula", "cache"),
    [
        ("=IF(Inputs!B1>60,IFERROR(SUM([1]Other!A1:A3),0),Inputs!B1)", {"A1": "1"}),
        ("=IF(Inputs!B1>60,IFERROR([1]Other!Z9,0),Inputs!B1)", {"A1": "1"}),
    ],
)
def test_external_reference_not_held_is_refused(
    make_book: Book, tmp_path: Path, formula: str, cache: dict[str, str]
) -> None:
    base = make_book("extx", {"Inputs": {"B1": 40}, "Calc": {"B1": F(formula, 40)}})
    book = add_external_link(base, tmp_path / "extx2.xlsx", "Other", cache)
    imap = {"id_field": "id", "fields": {"age": "Inputs!B1"}}
    res = calc(_req(tmp_path, book, "id,age\n1,70\n", imap=imap, outputs=["Calc!B1"]))
    row = _results(res)[("1", "Calc!B1")]
    assert row["status"] == "refused"
    assert "refused_pattern@Calc!B1:external reference not held" in row["reasons"]
    report = json.loads(res.trust_report.read_text(encoding="utf-8"))
    assert {"cell": "Calc!B1", "detail": "external reference not held"} in report[
        "refused_patterns"
    ]


def test_stale_outputs_are_removed(make_book: Book, tmp_path: Path) -> None:
    book = _book(make_book)
    variables = _write(tmp_path, "vars.json", json.dumps({"r": {"cell": "Calc!B1"}}))
    out = tmp_path / "out"
    out.mkdir()
    (out / "keep.txt").write_text("mine", encoding="utf-8")
    calc(_req(tmp_path, book, "id,age,sex\n1,50,M\n", variables=variables))
    assert (out / "rows.csv").exists()
    calc(_req(tmp_path, book, "id,age,sex\n1,50,M\n"))
    assert not (out / "rows.csv").exists()
    with pytest.raises(WorkbookRefused):
        calc(_req(tmp_path, _book(make_book, "stale1904", date1904=True), "id,age,sex\n1,50,M\n"))
    assert sorted(f.name for f in out.iterdir()) == ["keep.txt", "run.json", "trust_report.json"]


def test_coerce_input() -> None:
    assert coerce_input("2020-01-01", 1.0) == 43831.0
    assert coerce_input("12.5", 3.0) == 12.5
    assert coerce_input("true", False) is True
    assert coerce_input("FALSE", True) is False
    assert coerce_input("007", "x") == "007"
    with pytest.raises(CalcInputError):
        coerce_input("abc", 1.0)
    with pytest.raises(CalcInputError):
        coerce_input("yes", True)


# -- a --check-against mismatch refuses only the path that input set took ---------

_NESTED = "=IF(Inputs!B1>60,Inputs!B1*2,IF(Inputs!B1>30,Inputs!B1,0))"


def _scoped_book(make_book: Book, name: str) -> Path:
    """Set 1 (age 40) matches; set 2 (age 70, outer branch true) mismatches B1 and B5."""
    sheets: dict[str, dict[str, object]] = {
        "Inputs": {"B1": 40},
        "Calc": {"B1": F(_NESTED, 40), "B5": F("=Inputs!B1*3", 120)},
        "Results": {"A2": 1, "B2": 40, "C2": 120, "A3": 2, "B3": 141, "C3": 211},
        "Table": {"A2": 1, "B2": 40, "A3": 2, "B3": 70},
    }
    return make_book(name, sheets)


def _scoped_check(outputs: dict[str, str]) -> dict[str, object]:
    return {
        "results": {"sheet": "Results", "id_column": "A", "first_row": 2, "outputs": outputs},
        "inputs": {
            "table": {"sheet": "Table", "id_column": "A", "first_row": 2, "fields": {"age": "B"}}
        },
    }


def _scoped(make_book: Book, tmp_path: Path, outputs: dict[str, str], target: str) -> CalcResult:
    imap = {"id_field": "id", "fields": {"age": "Inputs!B1"}}
    return calc(
        _req(
            tmp_path,
            _scoped_book(make_book, "scoped"),
            "id,age\nsame,75\nchecked,50\nunchecked,10\n",
            imap=imap,
            outputs=[target],
            check_against=_write(tmp_path, "check.json", json.dumps(_scoped_check(outputs))),
        )
    )


def test_mismatch_refuses_a_policy_on_the_same_path(make_book: Book, tmp_path: Path) -> None:
    got = _results(_scoped(make_book, tmp_path, {"Calc!B1": "B"}, "Calc!B1"))
    row = got[("same", "Calc!B1")]
    assert (row["value"], row["status"]) == ("150.0", "refused")
    assert "check_against_mismatch@Calc!B1:input set '2'" in row["reasons"]


def test_mismatch_spares_a_policy_on_a_path_another_set_checked(
    make_book: Book, tmp_path: Path
) -> None:
    got = _results(_scoped(make_book, tmp_path, {"Calc!B1": "B"}, "Calc!B1"))
    row = got[("checked", "Calc!B1")]
    assert (row["value"], row["status"], row["reasons"]) == ("50.0", "ok", "")


def test_mismatch_leaves_an_unchecked_path_unverified(make_book: Book, tmp_path: Path) -> None:
    got = _results(_scoped(make_book, tmp_path, {"Calc!B1": "B"}, "Calc!B1"))
    row = got[("unchecked", "Calc!B1")]
    assert row["status"] == "unverified_path"
    assert "check_against_mismatch" not in row["reasons"]


def test_mismatch_on_a_probe_free_output_refuses_every_policy(
    make_book: Book, tmp_path: Path
) -> None:
    got = _results(_scoped(make_book, tmp_path, {"Calc!B5": "C"}, "Calc!B5"))
    for pid in ("same", "checked", "unchecked"):
        row = got[(pid, "Calc!B5")]
        assert row["status"] == "refused", pid
        assert "check_against_mismatch@Calc!B5" in row["reasons"]


def test_trust_report_lists_the_mismatch_with_its_difference(
    make_book: Book, tmp_path: Path
) -> None:
    res = _scoped(make_book, tmp_path, {"Calc!B1": "B"}, "Calc!B1")
    report = json.loads(res.trust_report.read_text(encoding="utf-8"))
    [mismatch] = report["check_against"]["mismatches"]
    assert mismatch["input_set"] == "2"
    assert mismatch["cell"] == "Calc!B1"
    assert (mismatch["stored"], mismatch["computed"]) == (141.0, 140.0)
    assert mismatch["relative_difference"] == pytest.approx(1 / 141)


# -- results stream to disk: memory does not grow with policies or check sets ---------


def _live(cls: type) -> int:
    return sum(1 for o in gc.get_objects() if type(o) is cls)


def test_results_stream_to_disk_one_policy_at_a_time(
    make_book: Book, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    sheets: dict[str, dict[str, object]] = {
        "Inputs": {"B1": 40},
        "Calc": {"B1": F(_NESTED, 40), "B2": F("=Calc!B1+1", 41)},
        "Results": {"A2": 1, "B2": 40, "A3": 2, "B3": 140, "A4": 3, "B4": 50, "A5": 4, "B5": 0},
        "Table": {"A2": 1, "B2": 40, "A3": 2, "B3": 70, "A4": 3, "B4": 50, "A5": 4, "B5": 10},
    }
    book = make_book("stream", sheets)
    out = tmp_path / "out"
    seen: list[tuple[int, int, int, int, int]] = []
    real = run_module._Session.recalc

    def spy(self: object) -> ProbeReadings:
        results = out / "results.csv"
        lines = len(results.read_text(encoding="utf-8").splitlines()) if results.exists() else -1
        paths = out / "paths.csv"
        plines = len(paths.read_text(encoding="utf-8").splitlines()) if paths.exists() else -1
        seen.append((lines, plines, _live(OutputValue), _live(PathRecord), _live(ProbeReadings)))
        return real(self)  # type: ignore[arg-type]

    monkeypatch.setattr(run_module._Session, "recalc", spy)
    n = 8
    policies = "id,age\n" + "".join(f"{i},{(10, 50, 75)[i % 3]}\n" for i in range(n))
    imap = {"id_field": "id", "fields": {"age": "Inputs!B1"}}
    check = _scoped_check({"Calc!B1": "B"})
    res = calc(
        _req(
            tmp_path,
            book,
            policies,
            imap=imap,
            outputs=["Calc!B1:B2"],
            check_against=_write(tmp_path, "check.json", json.dumps(check)),
        )
    )
    # Four check sets, then one recalculation per policy.
    assert len(seen) == 4 + n
    loop = seen[4:]
    # Before policy k recalculates, the k policies before it are already on disk.
    assert [s[0] for s in loop] == [1 + 2 * k for k in range(n)]
    assert all(b[1] > a[1] for a, b in zip(loop[1:], loop[2:], strict=False))
    # Nothing per policy or per check set stays in memory.
    assert max(s[2] for s in loop) <= 2
    assert max(s[3] for s in loop) <= 4
    assert max(s[4] for s in loop) <= 1
    assert len(res.results_csv.read_text(encoding="utf-8").splitlines()) == 1 + 2 * n


def test_failed_policy_leaves_no_partial_results(make_book: Book, tmp_path: Path) -> None:
    book = make_book("partial", {"Inputs": {"B1": 40}, "Calc": {"B1": F("=Inputs!B1*2", 80)}})
    imap = {"id_field": "id", "fields": {"age": "Inputs!B1"}}
    with pytest.raises(CalcInputError, match="not a number"):
        calc(_req(tmp_path, book, "id,age\n1,50\n2,X\n", imap=imap, outputs=["Calc!B1"]))
    out = tmp_path / "out"
    assert not any((out / name).exists() for name in ("results.csv", "rows.csv", "paths.csv"))
