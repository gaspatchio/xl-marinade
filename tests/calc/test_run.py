"""Tests for the ``calc()`` run: settings, trust run, probes, policy loop, output files."""

from __future__ import annotations

import csv
import json
import tempfile
from collections.abc import Callable
from pathlib import Path

import pytest

pytest.importorskip("ironcalc")

from tests.calc.conftest import F, add_external_link  # noqa: E402
from xl_marinade.calc import CalcRequest, CalcResult, calc  # noqa: E402
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


def test_inputs_are_reset_between_policies(make_book: Book, tmp_path: Path) -> None:
    book = _book(make_book)
    imap = {"id_field": "id", "fields": {"age": "Inputs!B1"}}
    res = calc(_req(tmp_path, book, "id,age\n1,70\n2,50\n", imap=imap, outputs=["Calc!B4"]))
    assert _results(res)[("2", "Calc!B4")]["value"] == "1.0"


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
    res = calc(_req(tmp_path, book, "id,age,sex\n9,50,M\n", check_against=check))
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
