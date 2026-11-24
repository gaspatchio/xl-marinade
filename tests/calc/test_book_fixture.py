import openpyxl

from tests.calc.conftest import F


def test_make_book_writes_cached_values(make_book):
    p = make_book(
        "b",
        {
            "S": {
                "A1": 2,
                "B1": F("=A1*3", 6),
                "C1": F('=IF(A1>1,"hi","lo")', "hi"),
                "D1": F("=NA()", "#N/A"),
                "E1": F("=A1>1", True),
            }
        },
    )
    wv = openpyxl.load_workbook(p, data_only=True)["S"]
    assert (wv["B1"].value, wv["C1"].value, wv["D1"].value, wv["E1"].value) == (
        6,
        "hi",
        "#N/A",
        True,
    )
    wf = openpyxl.load_workbook(p)["S"]
    assert wf["B1"].value == "=A1*3"


def test_make_book_settings_and_names(make_book):
    p = make_book(
        "n",
        {"S": {"A1": 1}},
        names={"Rate": "S!$A$1"},
        date1904=True,
        manual=True,
        iterate=True,
        full_precision=False,
        validations=[("S!A1", ["1", "2"])],
    )
    wb = openpyxl.load_workbook(p)
    assert wb.epoch.year == 1904
    assert "Rate" in wb.defined_names
    assert wb.calculation.calcMode == "manual"
    assert wb.calculation.iterate is True
    assert wb.calculation.fullPrecision is False
    assert len(wb["S"].data_validations.dataValidation) == 1


def test_ir_for_extracts(make_book, ir_for):
    p = make_book("e", {"S": {"A1": 2, "B1": F("=A1*3", 6)}})
    assert ir_for(p).exists()
