# ABOUTME: Unit tests for formula parser - validates AST generation for all node types
# ABOUTME: Tests constants, references, functions, and operators with comprehensive coverage

import pytest

from xl_marinade.core.parser import parse_formula


class TestFormulaParser:
    """Test formula parser AST generation"""

    def test_parse_const_number_integer(self):
        """Test parsing integer constant"""
        result = parse_formula("42")
        assert result == {"type": "Const", "value": 42}

    def test_parse_const_number_float(self):
        """Test parsing float constant"""
        result = parse_formula("3.14")
        assert result == {"type": "Const", "value": 3.14}

    def test_parse_const_string(self):
        """Test parsing string constant"""
        result = parse_formula('"Hello"')
        assert result == {"type": "Const", "value": "Hello"}

    def test_parse_const_string_with_escaped_quote(self):
        """Test parsing string with escaped quote"""
        result = parse_formula('"Hello ""World"""')
        assert result == {"type": "Const", "value": 'Hello "World"'}

    def test_parse_const_boolean_true(self):
        """Test parsing TRUE constant"""
        result = parse_formula("TRUE")
        assert result == {"type": "Const", "value": True}

    def test_parse_const_boolean_false(self):
        """Test parsing FALSE constant"""
        result = parse_formula("FALSE")
        assert result == {"type": "Const", "value": False}

    def test_parse_const_error(self):
        """Test parsing error constant"""
        result = parse_formula("#REF!")
        assert result == {"type": "Const", "value": "#REF!"}

    def test_parse_ref_simple(self):
        """Test parsing simple cell reference"""
        result = parse_formula("A1")
        assert result == {"type": "Ref", "ref": "A1"}

    def test_parse_ref_absolute(self):
        """Test parsing absolute reference"""
        result = parse_formula("$A$1")
        assert result == {"type": "Ref", "ref": "$A$1"}

    def test_parse_ref_mixed(self):
        """Test parsing mixed reference"""
        result = parse_formula("$A1")
        assert result == {"type": "Ref", "ref": "$A1"}

        result = parse_formula("A$1")
        assert result == {"type": "Ref", "ref": "A$1"}

    def test_parse_ref_sheet_qualified(self):
        """Test parsing sheet-qualified reference"""
        result = parse_formula("Sheet1!A1")
        assert result == {"type": "Ref", "ref": "Sheet1!A1"}

    def test_parse_ref_range(self):
        """Test parsing range reference"""
        result = parse_formula("A1:B10")
        assert result == {"type": "Ref", "ref": "A1:B10"}

    def test_parse_ref_structured(self):
        """Test parsing structured table reference"""
        result = parse_formula("Table1[Revenue]")
        assert result == {"type": "Ref", "ref": "Table1[Revenue]"}

    def test_parse_ref_defined_name(self):
        """Test parsing defined name"""
        result = parse_formula("MyVariable")
        assert result == {"type": "Ref", "ref": "MyVariable"}

    def test_parse_function_no_args(self):
        """Test parsing function with no arguments"""
        result = parse_formula("NOW()")
        assert result == {"type": "Function", "name": "NOW", "args": []}

    def test_parse_function_one_arg(self):
        """Test parsing function with one argument"""
        result = parse_formula("SUM(A1)")
        assert result == {"type": "Function", "name": "SUM", "args": [{"type": "Ref", "ref": "A1"}]}

    def test_parse_function_multiple_args(self):
        """Test parsing function with multiple arguments"""
        result = parse_formula("SUM(A1,B1,C1)")
        assert result == {
            "type": "Function",
            "name": "SUM",
            "args": [
                {"type": "Ref", "ref": "A1"},
                {"type": "Ref", "ref": "B1"},
                {"type": "Ref", "ref": "C1"},
            ],
        }

    def test_parse_function_range_arg(self):
        """Test parsing function with range argument"""
        result = parse_formula("SUM(A1:A10)")
        assert result == {
            "type": "Function",
            "name": "SUM",
            "args": [{"type": "Ref", "ref": "A1:A10"}],
        }

    def test_parse_function_nested(self):
        """Test parsing nested functions"""
        result = parse_formula("SUM(AVERAGE(A1:A10),B1)")
        assert result["type"] == "Function"
        assert result["name"] == "SUM"
        assert len(result["args"]) == 2
        assert result["args"][0]["type"] == "Function"
        assert result["args"][0]["name"] == "AVERAGE"

    def test_parse_unary_minus(self):
        """Test parsing unary minus"""
        result = parse_formula("-A1")
        assert result == {"type": "Unary", "operator": "-", "operand": {"type": "Ref", "ref": "A1"}}

    def test_parse_unary_plus(self):
        """Test parsing unary plus"""
        result = parse_formula("+A1")
        assert result == {"type": "Unary", "operator": "+", "operand": {"type": "Ref", "ref": "A1"}}

    def test_parse_binary_addition(self):
        """Test parsing addition operator"""
        result = parse_formula("A1+B1")
        assert result == {
            "type": "Binary",
            "operator": "+",
            "left": {"type": "Ref", "ref": "A1"},
            "right": {"type": "Ref", "ref": "B1"},
        }

    def test_parse_binary_subtraction(self):
        """Test parsing subtraction operator"""
        result = parse_formula("A1-B1")
        assert result == {
            "type": "Binary",
            "operator": "-",
            "left": {"type": "Ref", "ref": "A1"},
            "right": {"type": "Ref", "ref": "B1"},
        }

    def test_parse_binary_multiplication(self):
        """Test parsing multiplication operator"""
        result = parse_formula("A1*B1")
        assert result == {
            "type": "Binary",
            "operator": "*",
            "left": {"type": "Ref", "ref": "A1"},
            "right": {"type": "Ref", "ref": "B1"},
        }

    def test_parse_binary_division(self):
        """Test parsing division operator"""
        result = parse_formula("A1/B1")
        assert result == {
            "type": "Binary",
            "operator": "/",
            "left": {"type": "Ref", "ref": "A1"},
            "right": {"type": "Ref", "ref": "B1"},
        }

    def test_parse_binary_exponentiation(self):
        """Test parsing exponentiation operator"""
        result = parse_formula("A1^2")
        assert result == {
            "type": "Binary",
            "operator": "^",
            "left": {"type": "Ref", "ref": "A1"},
            "right": {"type": "Const", "value": 2},
        }

    def test_parse_binary_concatenation(self):
        """Test parsing string concatenation operator"""
        result = parse_formula('A1&"text"')
        assert result == {
            "type": "Binary",
            "operator": "&",
            "left": {"type": "Ref", "ref": "A1"},
            "right": {"type": "Const", "value": "text"},
        }

    def test_parse_binary_comparison_equal(self):
        """Test parsing equality comparison"""
        result = parse_formula("A1=B1")
        assert result["type"] == "Binary"
        assert result["operator"] == "="

    def test_parse_binary_comparison_not_equal(self):
        """Test parsing not-equal comparison"""
        result = parse_formula("A1<>B1")
        assert result["type"] == "Binary"
        assert result["operator"] == "<>"

    def test_parse_binary_comparison_less_than(self):
        """Test parsing less-than comparison"""
        result = parse_formula("A1<B1")
        assert result["type"] == "Binary"
        assert result["operator"] == "<"

    def test_parse_binary_comparison_greater_than(self):
        """Test parsing greater-than comparison"""
        result = parse_formula("A1>B1")
        assert result["type"] == "Binary"
        assert result["operator"] == ">"

    def test_parse_binary_comparison_less_equal(self):
        """Test parsing less-than-or-equal comparison"""
        result = parse_formula("A1<=B1")
        assert result["type"] == "Binary"
        assert result["operator"] == "<="

    def test_parse_binary_comparison_greater_equal(self):
        """Test parsing greater-than-or-equal comparison"""
        result = parse_formula("A1>=B1")
        assert result["type"] == "Binary"
        assert result["operator"] == ">="

    def test_parse_operator_precedence_addition_multiplication(self):
        """Test operator precedence: multiplication before addition"""
        result = parse_formula("A1+B1*C1")
        assert result["type"] == "Binary"
        assert result["operator"] == "+"
        assert result["left"] == {"type": "Ref", "ref": "A1"}
        assert result["right"]["type"] == "Binary"
        assert result["right"]["operator"] == "*"

    def test_parse_operator_precedence_exponentiation(self):
        """Test operator precedence: exponentiation before multiplication"""
        result = parse_formula("A1*B1^2")
        assert result["type"] == "Binary"
        assert result["operator"] == "*"
        assert result["right"]["type"] == "Binary"
        assert result["right"]["operator"] == "^"

    def test_parse_parentheses(self):
        """Test parentheses override precedence"""
        result = parse_formula("(A1+B1)*C1")
        assert result["type"] == "Binary"
        assert result["operator"] == "*"
        assert result["left"]["type"] == "Binary"
        assert result["left"]["operator"] == "+"

    def test_parse_complex_formula(self):
        """Test complex formula with multiple operators"""
        result = parse_formula("SUM(A1:A10)+AVERAGE(B1:B10)*2")
        assert result["type"] == "Binary"
        assert result["operator"] == "+"
        assert result["left"]["type"] == "Function"
        assert result["right"]["type"] == "Binary"
        assert result["right"]["operator"] == "*"

    def test_parse_with_leading_equals(self):
        """Test parsing formula with leading = sign"""
        result = parse_formula("=A1+B1")
        assert result["type"] == "Binary"
        assert result["operator"] == "+"

    def test_parse_with_whitespace(self):
        """Test parsing formula with whitespace"""
        result = parse_formula(" A1 + B1 ")
        assert result["type"] == "Binary"
        assert result["operator"] == "+"

    def test_parse_empty_formula_raises(self):
        """Test that empty formula raises ValueError"""
        with pytest.raises(ValueError, match="Empty formula"):
            parse_formula("")

    def test_parse_unterminated_string_raises(self):
        """Test that unterminated string raises ValueError"""
        with pytest.raises(ValueError, match="Unterminated string"):
            parse_formula('"Hello')

    def test_parse_missing_closing_paren_raises(self):
        """Test that missing closing parenthesis raises ValueError"""
        with pytest.raises(ValueError, match="Expected '\\)'"):
            parse_formula("(A1+B1")

    def test_parse_vlookup_example(self):
        """Test parsing VLOOKUP formula from IR spec"""
        result = parse_formula("VLOOKUP(A1,$D$1:$F$10,3,FALSE)")
        assert result["type"] == "Function"
        assert result["name"] == "VLOOKUP"
        assert len(result["args"]) == 4
        assert result["args"][0] == {"type": "Ref", "ref": "A1"}
        assert result["args"][1] == {"type": "Ref", "ref": "$D$1:$F$10"}
        assert result["args"][2] == {"type": "Const", "value": 3}
        assert result["args"][3] == {"type": "Const", "value": False}

    def test_parse_index_match_example(self):
        """Test parsing INDEX/MATCH formula"""
        result = parse_formula("INDEX($B$1:$B$10,MATCH(A1,$A$1:$A$10,0))")
        assert result["type"] == "Function"
        assert result["name"] == "INDEX"
        assert len(result["args"]) == 2
        assert result["args"][1]["type"] == "Function"
        assert result["args"][1]["name"] == "MATCH"

    def test_parse_offset_example(self):
        """Test parsing OFFSET formula"""
        result = parse_formula("OFFSET(A1,1,1)")
        assert result["type"] == "Function"
        assert result["name"] == "OFFSET"
        assert len(result["args"]) == 3

    def test_parse_indirect_example(self):
        """Test parsing INDIRECT formula"""
        result = parse_formula('INDIRECT("A"&B1)')
        assert result["type"] == "Function"
        assert result["name"] == "INDIRECT"
        assert len(result["args"]) == 1
        assert result["args"][0]["type"] == "Binary"
        assert result["args"][0]["operator"] == "&"

    def test_parse_sheet_name_with_spaces(self):
        """
        Test parsing sheet references with spaces in sheet names.

        Regression test for bug where parser truncated at first space.
        Excel requires single quotes around sheet names with spaces: 'My Sheet'!A1
        """
        # Single space in sheet name
        result = parse_formula("='Interest rate'!B1")
        assert result["type"] == "Ref"
        assert result["ref"] == "'Interest rate'!B1"

        # Multiple spaces
        result = parse_formula("='My Sheet'!C5")
        assert result["type"] == "Ref"
        assert result["ref"] == "'My Sheet'!C5"

        # Multiple words
        result = parse_formula("='A B C'!D10")
        assert result["type"] == "Ref"
        assert result["ref"] == "'A B C'!D10"

        # In complex formula
        result = parse_formula("=SUM('Monthly Data'!A1:A10)")
        assert result["type"] == "Function"
        assert result["name"] == "SUM"
        assert result["args"][0]["type"] == "Ref"
        assert result["args"][0]["ref"] == "'Monthly Data'!A1:A10"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
