# ABOUTME: Tests for sheet-name quoting normalization in formula comparison —
# 'Calculations'! vs Calculations! is the same reference, not a formula change.

from xl_marinade.core.ir_diff.diff_emit import normalize_sheet_quoting, same_formula
from xl_marinade.core.ir_diff.pipeline import _classify_formula_change


def test_normalize_drops_redundant_quotes():
    assert (
        normalize_sheet_quoting("SUM('Calculations'!R[43]C6:'Calculations'!R[43]C[-1])")
        == "SUM(Calculations!R[43]C6:Calculations!R[43]C[-1])"
    )


def test_normalize_keeps_quotes_that_are_required():
    # spaces / punctuation force quoting — leave untouched
    assert normalize_sheet_quoting("'My Sheet'!R1C1") == "'My Sheet'!R1C1"
    assert normalize_sheet_quoting("'Ver 1.2'!RC") == "'Ver 1.2'!RC"
    # embedded (escaped) quote never matches the identifier pattern
    assert normalize_sheet_quoting("'Bob''s'!RC") == "'Bob''s'!RC"


def test_normalize_leaves_string_literals_alone():
    f = "IF(RC1=\"'Calculations'!\",'Calculations'!RC2,0)"
    assert normalize_sheet_quoting(f) == "IF(RC1=\"'Calculations'!\",Calculations!RC2,0)"


def test_normalize_passthrough():
    assert normalize_sheet_quoting(None) is None
    assert normalize_sheet_quoting("") == ""
    assert normalize_sheet_quoting("R[1]C+1") == "R[1]C+1"


def test_same_formula_quoting_only_is_equal():
    assert same_formula("Calculations!R[43]C6", "'Calculations'!R[43]C6")
    assert same_formula(None, None)


def test_same_formula_real_change_still_differs():
    assert not same_formula("'Calculations'!R[43]C6", "'Calculations'!R[44]C6")
    assert not same_formula("RC1", None)


def test_classify_ignores_quoting_in_shift_detection():
    # absolute row shifted AND quoting changed -> still a reference_shift
    assert (
        _classify_formula_change("SUM('Calculations'!R10C6)", "SUM(Calculations!R11C6)")
        == "reference_shift"
    )
    # genuine logic difference stays logic_change regardless of quoting
    assert (
        _classify_formula_change("SUM('Calculations'!R10C6)", "MAX(Calculations!R10C6)")
        == "logic_change"
    )
