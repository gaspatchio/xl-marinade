"""Data-validation lists checked the way Excel checks them.

Hand values follow Excel's list validation: the list comes from literal items, a range,
or a defined name over a range (a sheet's own name first, as in a formula); a whole
column means the cells it holds; the entry matches an item ignoring case, and a date or
number matches by value (an ISO date typed into a date cell is that date's serial).
Lists fed by a formula (``INDIRECT``, ``OFFSET``) are not checked (ruling D8).
"""

from __future__ import annotations

import csv
import json
from collections.abc import Callable
from pathlib import Path

import pytest

pytest.importorskip("ironcalc")

from tests.calc.conftest import F  # noqa: E402
from xl_marinade.calc import CalcRequest, calc  # noqa: E402

Book = Callable[..., Path]

# One input per list style; Out!An reads Inputs!Bn so every input feeds an output.
INPUTS: dict[str, object] = {
    "B1": "Gold",
    "B2": "A",
    "B3": "P",
    "B4": "M",
    "B5": 45292,  # 2024-01-01
    "B6": "Yes",
    "B7": "k1",
    "B8": "Gold",
    "B9": "Gold",
    "B10": "A",
    "H1": "P",
    "H2": "Q",
    "J1": "A",
    "J2": "B",
    "K1": "k1",
    "K2": "k2",
}
LISTS: dict[str, object] = {
    "A1": "Gold",
    "A2": "Silver",
    "B1": 45292,  # 2024-01-01
    "B2": 45658,  # 2025-01-01
    "C1": "w1",
    "C2": "w2",
}
FIELDS = {
    "plan": "Inputs!B1",
    "code": "Inputs!B2",
    "col": "Inputs!B3",
    "sex": "Inputs!B4",
    "dt": "Inputs!B5",
    "yn": "Inputs!B6",
    "opt": "Inputs!B7",
    "ind": "Inputs!B8",
    "dyn": "Inputs!B9",
    "eq": "Inputs!B10",
}
SOURCES = [
    ("Inputs!B1", "PlanList"),  # a workbook name
    ("Inputs!B2", "Codes"),  # a name over the same sheet
    ("Inputs!B3", "$H:$H"),  # a whole column
    ("Inputs!B4", '"M,F"'),  # literal items
    ("Inputs!B5", "DateList"),  # a name over date serials
    ("Inputs!B6", '"Yes,No"'),  # literal items, compared ignoring case
    ("Inputs!B7", "Opts"),  # Inputs' own Opts, not the workbook's
    ("Inputs!B8", 'INDIRECT("Lists!A1:A2")'),  # a formula: not checked (D8)
    ("Inputs!B9", "Dyn"),  # an OFFSET name: a formula, not checked (D8)
    ("Inputs!B10", "=Codes"),  # a name written with "="
]
BASE = {
    "plan": "Silver",
    "code": "B",
    "col": "Q",
    "sex": "F",
    "dt": "45658",
    "yn": "No",
    "opt": "k2",
    "ind": "Silver",
    "dyn": "Silver",
    "eq": "B",
}
# policy id -> (field, value, Excel accepts it)
CASES = {
    "base": (None, None, True),
    "plan_out": ("plan", "Platinum", False),
    "code_out": ("code", "Z", False),
    "col_out": ("col", "ZZ", False),
    "col_case": ("col", "q", True),
    "sex_out": ("sex", "X", False),
    "date_iso_in": ("dt", "2025-01-01", True),
    "date_iso_out": ("dt", "2026-01-01", False),
    "date_serial_in": ("dt", "45292", True),
    "yn_case": ("yn", "yes", True),
    "plan_case": ("plan", "GOLD", True),
    "opt_own": ("opt", "k1", True),
    "opt_workbook": ("opt", "w1", False),
    "ind_any": ("ind", "Anything", True),
    "dyn_any": ("dyn", "Anything", True),
    "eq_out": ("eq", "Z", False),
}


def test_list_validations(make_book: Book, tmp_path: Path) -> None:
    out = {f"A{i}": F(f'=Inputs!B{i}&""', str(INPUTS[f"B{i}"])) for i in range(1, 11)}
    book = make_book(
        "dv",
        {"Inputs": INPUTS, "Lists": LISTS, "Out": out},
        names={
            "PlanList": "Lists!$A$1:$A$2",
            "Codes": "Inputs!$J$1:$J$2",
            "DateList": "Lists!$B$1:$B$2",
            "Opts": "Lists!$C$1:$C$2",
            "Dyn": "OFFSET(Lists!$A$1,0,0,2,1)",
        },
        local_names={("Inputs", "Opts"): "Inputs!$K$1:$K$2"},
        list_sources=SOURCES,
    )
    header = ["id", *FIELDS]
    lines = [",".join(header)]
    for pid, (field, value, _ok) in CASES.items():
        row = dict(BASE)
        if field is not None:
            row[field] = str(value)
        lines.append(",".join([pid, *(row[f] for f in FIELDS)]))
    policies = tmp_path / "policies.csv"
    policies.write_text("\n".join(lines) + "\n", encoding="utf-8")
    imap = tmp_path / "map.json"
    imap.write_text(json.dumps({"fields": FIELDS}), encoding="utf-8")
    res = calc(CalcRequest(book, policies, imap, tmp_path / "out", outputs=["Out!A1:A10"]))
    with res.results_csv.open(encoding="utf-8", newline="") as fh:
        rows = list(csv.DictReader(fh))
    refused = {
        r["policy_id"]
        for r in rows
        if "input_outside_allowed_values" in r["reasons"] and r["output"] == "Out!A1"
    }
    want = {pid for pid, (_f, _v, ok) in CASES.items() if not ok}
    assert refused == want


def test_numbers_match_number_items_only(make_book: Book, tmp_path: Path) -> None:
    # Excel compares the number 1 with the list's items: a cell holding the number 1
    # matches, the text "001" does not. A literal list "1,2" holds numbers.
    book = make_book(
        "numtext",
        {
            "In": {"B1": 1, "B2": 1, "B3": 1},
            "Lists": {"A1": "001", "A2": "002", "B1": 1, "B2": 2},
            "Out": {
                "A1": F("=In!B1*10", 10),
                "A2": F("=In!B2*10", 10),
                "A3": F("=In!B3*10", 10),
            },
        },
        list_sources=[
            ("In!B1", "Lists!$A$1:$A$2"),
            ("In!B2", "Lists!$B$1:$B$2"),
            ("In!B3", '"1,2"'),
        ],
    )
    policies = tmp_path / "policies.csv"
    policies.write_text("id,t,n,l\ntext,1,1,1\nnum,001,2,2\n", encoding="utf-8")
    imap = tmp_path / "map.json"
    imap.write_text(
        json.dumps({"fields": {"t": "In!B1", "n": "In!B2", "l": "In!B3"}}), encoding="utf-8"
    )
    res = calc(CalcRequest(book, policies, imap, tmp_path / "out", outputs=["Out!A1"]))
    with res.results_csv.open(encoding="utf-8", newline="") as fh:
        reasons = {r["policy_id"]: r["reasons"] for r in csv.DictReader(fh)}
    # Policy "text": 1 against "001","002" -> refused; 1 against 1,2 and "1,2" -> fine.
    assert "input_outside_allowed_values@In!B1" in reasons["text"]
    assert "@In!B2" not in reasons["text"] and "@In!B3" not in reasons["text"]
    # Policy "num": "001" typed into a number cell is the number 1 -> not "001" either.
    assert "input_outside_allowed_values@In!B1" in reasons["num"]
