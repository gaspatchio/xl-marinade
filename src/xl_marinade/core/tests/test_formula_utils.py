# ABOUTME: Tests for formula extraction utilities including ArrayFormula handling.
# ABOUTME: Ensures robust handling of openpyxl cell.value including ArrayFormula objects.

from unittest.mock import Mock

import pytest

from xl_marinade.core.formula_utils import extract_formula_string


class TestExtractFormulaString:
    """Test formula extraction from openpyxl cells."""

    def test_extract_regular_formula(self):
        """Test extraction of regular formula string."""
        cell = Mock()
        cell.value = "=A1+B1"

        result = extract_formula_string(cell)

        assert result == "=A1+B1"

    def test_extract_complex_formula(self):
        """Test extraction of complex formula."""
        cell = Mock()
        cell.value = "=SUM(A1:A10)*INDEX(B1:B10,MATCH(C1,D1:D10,0))"

        result = extract_formula_string(cell)

        assert result == "=SUM(A1:A10)*INDEX(B1:B10,MATCH(C1,D1:D10,0))"

    def test_extract_array_formula_with_text(self):
        """Test extraction of ArrayFormula object with .text attribute."""
        # Mock ArrayFormula object
        array_formula = Mock()
        array_formula.text = (
            "=INDEX('Mortality table'!$B$2:$F$102, MATCH(C8, 'Mortality table'!$A$2:$A$102, 0))"
        )

        # Mock cell with ArrayFormula
        cell = Mock()
        cell.value = array_formula

        # Make isinstance(cell.value, ArrayFormula) work
        # We need to actually import and use the real ArrayFormula class for isinstance
        from openpyxl.worksheet.formula import ArrayFormula

        cell.value = Mock(spec=ArrayFormula)
        cell.value.text = (
            "=INDEX('Mortality table'!$B$2:$F$102, MATCH(C8, 'Mortality table'!$A$2:$A$102, 0))"
        )

        result = extract_formula_string(cell)

        assert (
            result
            == "=INDEX('Mortality table'!$B$2:$F$102, MATCH(C8, 'Mortality table'!$A$2:$A$102, 0))"
        )

    def test_extract_array_formula_without_text(self):
        """Test extraction of ArrayFormula object without .text attribute (edge case)."""
        from openpyxl.worksheet.formula import ArrayFormula

        cell = Mock()
        cell.value = Mock(spec=ArrayFormula)
        # Simulate missing .text attribute
        del cell.value.text

        result = extract_formula_string(cell)

        # Should return empty string when .text is missing
        assert result == ""

    def test_extract_array_formula_with_none_text(self):
        """Test extraction of ArrayFormula object with None .text."""
        from openpyxl.worksheet.formula import ArrayFormula

        cell = Mock()
        cell.value = Mock(spec=ArrayFormula)
        cell.value.text = None

        result = extract_formula_string(cell)

        # Should return empty string when .text is None
        assert result == ""

    def test_extract_from_value_only_cell_number(self):
        """Test extraction from cell with number value (no formula)."""
        cell = Mock()
        cell.value = 123.45

        result = extract_formula_string(cell)

        assert result == ""

    def test_extract_from_value_only_cell_string(self):
        """Test extraction from cell with string value (not a formula)."""
        cell = Mock()
        cell.value = "Hello World"

        result = extract_formula_string(cell)

        # Non-formula strings should be returned as-is
        # This matches the current behavior where cell.value could be a plain string
        assert result == "Hello World"

    def test_extract_from_empty_cell(self):
        """Test extraction from empty cell (None value)."""
        cell = Mock()
        cell.value = None

        result = extract_formula_string(cell)

        assert result == ""

    def test_extract_from_boolean_cell(self):
        """Test extraction from cell with boolean value."""
        cell = Mock()
        cell.value = True

        result = extract_formula_string(cell)

        assert result == ""

    def test_extract_preserves_formula_exactly(self):
        """Test that formula is preserved exactly (no normalization)."""
        # Per IR spec, formulas should be hashed exactly as read
        cell = Mock()
        cell.value = "=sum( A1:A2 )"  # Note: lowercase, spaces

        result = extract_formula_string(cell)

        # Should preserve exact casing and whitespace
        assert result == "=sum( A1:A2 )"

    def test_extract_cross_sheet_formula(self):
        """Test extraction of cross-sheet reference formula."""
        cell = Mock()
        cell.value = "='Sheet with spaces'!A1+Sheet2!B2"

        result = extract_formula_string(cell)

        assert result == "='Sheet with spaces'!A1+Sheet2!B2"

    def test_extract_from_error_cell(self):
        """Test extraction from cell with error value."""
        cell = Mock()
        cell.value = "#DIV/0!"

        result = extract_formula_string(cell)

        # Error strings should be returned as-is (they're strings in cell.value)
        assert result == "#DIV/0!"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
