# ABOUTME: Defined names that lex like cell refs (T0, ZZZ1, AB1x) must tokenize as
# ABOUTME: identifiers, not crash parse_a1_ref with "row out of bounds: 0".

"""
Excel forbids defined names that collide with real cell references, but it
ALLOWS ref-lookalikes whose reference would be invalid: T0 (row 0 does not
exist), ZZZ1 (column ZZZ is past XFD), AB1x (continues with identifier
characters). Such names are common in real actuarial/finance workbooks
(T0 for valuation time, Q1/FY2024-style names are banned but T0 slips the
ban precisely because it is not a valid ref).

Before this fix, the tokenizer classified them as cell_ref and
normalize_formula died with ValueError("row out of bounds: 0") — an entire
sheet extraction aborted. Found by extracting a Solvency II workbook whose
valuation-time control cell is named T0.
"""

import pytest

from xl_marinade.core.new_arch.formula_normalizer import (
    FormulaContext,
    FormulaTokenizer,
    normalize_formula,
)


@pytest.fixture
def tok():
    return FormulaTokenizer()


def _refs_and_names(tokens):
    return [(t.type, t.value) for t in tokens if t.type in ("cell_ref", "range_ref", "identifier")]


class TestNameLookalikeTokens:
    def test_row_zero_name_is_identifier(self, tok):
        assert _refs_and_names(tok.tokenize("IF(A7>T0,1,0)")) == [
            ("cell_ref", "A7"),
            ("identifier", "T0"),
        ]

    def test_column_past_xfd_is_identifier(self, tok):
        assert _refs_and_names(tok.tokenize("T0+ZZZ1")) == [
            ("identifier", "T0"),
            ("identifier", "ZZZ1"),
        ]

    def test_ref_followed_by_identifier_chars_is_identifier(self, tok):
        # AB1 is a valid ref, but AB1x can only be a name — the token must
        # not be split into cell_ref AB1 + identifier x.
        assert _refs_and_names(tok.tokenize("AB1x&T0")) == [
            ("identifier", "AB1x"),
            ("identifier", "T0"),
        ]

    def test_valid_refs_still_tokenize_as_refs(self, tok):
        assert _refs_and_names(tok.tokenize("SUM(Q1:Q4)+MAX(B3,XFD1048576)")) == [
            ("range_ref", "Q1:Q4"),
            ("cell_ref", "B3"),
            ("cell_ref", "XFD1048576"),
        ]

    def test_range_with_invalid_endpoint_not_a_range(self, tok):
        # T0:T5 cannot be a range (T0 is not a cell); T5 alone still is one.
        types = _refs_and_names(tok.tokenize("A1:B2+T0:T5"))
        assert types == [("range_ref", "A1:B2"), ("identifier", "T0"), ("cell_ref", "T5")]


class TestNormalizeWithNameLookalikes:
    def test_normalize_keeps_name_verbatim_no_crash(self):
        ctx = FormulaContext(sheet_id=1, row=7, col=1, sheet_name="Proj")
        out = normalize_formula("IF(A7>T0,0,B6*(1+InflRate))", ctx, {})
        # T0 survives as a name; the real refs are converted to R1C1.
        assert "T0" in out
        assert "R[" in out or "RC" in out

    def test_normalize_row_zero_name_alone(self):
        ctx = FormulaContext(sheet_id=1, row=3, col=2, sheet_name="Ctl")
        out = normalize_formula("T0+1", ctx, {})
        assert "T0" in out
