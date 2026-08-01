# ABOUTME: Formula parser that converts Excel formulas to structural AST representation
# ABOUTME: Supports Const, Ref, Function, Unary, and Binary node types per IR spec §4

from __future__ import annotations

from typing import Any


class ASTNode:
    """Base class for AST nodes"""

    def __init__(self, node_type: str):
        self.node_type = node_type

    def to_dict(self) -> dict[str, Any]:
        """Convert node to dictionary representation"""
        raise NotImplementedError

    def to_string(self) -> str:
        """Convert AST node back to Excel formula syntax"""
        raise NotImplementedError(f"{self.__class__.__name__}.to_string() not implemented")

    def to_r1c1_string(self, base_row: int, base_col: int) -> str:
        """Convert AST node to R1C1 formula syntax.

        Args:
            base_row: Row of cell containing formula (1-indexed)
            base_col: Column of cell containing formula (1-indexed)

        Returns:
            Formula string in R1C1 notation.
        """
        raise NotImplementedError(f"{self.__class__.__name__}.to_r1c1_string() not implemented")


class ConstNode(ASTNode):
    """Constant literal node (number, string, boolean, error)"""

    def __init__(self, value: int | float | str | bool):
        super().__init__("Const")
        self.value = value

    def to_dict(self) -> dict[str, Any]:
        return {"type": self.node_type, "value": self.value}

    def to_string(self) -> str:
        if isinstance(self.value, str):
            # String: wrap in quotes, escape internal quotes
            # Excel uses "" for escaped quote (not \")
            escaped = self.value.replace('"', '""')
            return f'"{escaped}"'
        elif isinstance(self.value, bool):
            return "TRUE" if self.value else "FALSE"
        elif isinstance(self.value, (int, float)):
            return str(self.value)
        else:
            # Errors like #REF!, #N/A, or other special values
            return str(self.value)

    def to_r1c1_string(self, base_row: int, base_col: int) -> str:
        """Return constant unchanged. String literals are NEVER converted to R1C1.

        This is critical for INDIRECT/OFFSET formulas where cell references appear
        as string arguments - these must NOT be converted.

        Args:
            base_row: Row of cell containing formula (ignored for constants)
            base_col: Column of cell containing formula (ignored for constants)

        Returns:
            The constant's string representation, unchanged.
        """
        return self.to_string()


class RefNode(ASTNode):
    """Cell or range reference node"""

    def __init__(self, ref: str):
        super().__init__("Ref")
        self.ref = ref  # A1 notation (e.g., "A1", "A1:B10", "Sheet1!A1")

    def to_dict(self) -> dict[str, Any]:
        return {"type": self.node_type, "ref": self.ref}

    def to_string(self) -> str:
        return self.ref  # Already in A1 notation

    def to_r1c1_string(self, base_row: int, base_col: int) -> str:
        """Convert cell reference from A1 to R1C1 notation.

        Args:
            base_row: Row of cell containing formula (1-indexed)
            base_col: Column of cell containing formula (1-indexed)

        Returns:
            Reference in R1C1 notation.
        """
        from xl_marinade.core.ref_converter import a1_to_r1c1

        try:
            # Handle ranges: A1:B10 or Sheet!A1:B10
            if ":" in self.ref:
                sheet_prefix = ""
                range_part = self.ref
                if "!" in self.ref:
                    # Use rsplit to handle sheet names containing '!' (e.g. 'Hi!There'!A1)
                    parts = self.ref.rsplit("!", 1)
                    sheet_prefix = parts[0] + "!"
                    range_part = parts[1]

                start, end = range_part.split(":", 1)
                start_r1c1 = a1_to_r1c1(start, base_row, base_col)
                end_r1c1 = a1_to_r1c1(end, base_row, base_col)

                if sheet_prefix:
                    return f"{sheet_prefix}{start_r1c1}:{end_r1c1}"
                return f"{start_r1c1}:{end_r1c1}"
            else:
                return a1_to_r1c1(self.ref, base_row, base_col)
        except (ValueError, AttributeError):
            # External refs or malformed - return unchanged
            return self.ref


class FunctionNode(ASTNode):
    """Function call node"""

    def __init__(self, name: str, args: list[ASTNode]):
        super().__init__("Function")
        self.name = name
        self.args = args

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": self.node_type,
            "name": self.name,
            "args": [arg.to_dict() for arg in self.args],
        }

    def to_string(self) -> str:
        args_str = ",".join(arg.to_string() for arg in self.args)
        return f"{self.name}({args_str})"

    def to_r1c1_string(self, base_row: int, base_col: int) -> str:
        """Convert function arguments to R1C1 recursively.

        Args:
            base_row: Row of cell containing formula (1-indexed)
            base_col: Column of cell containing formula (1-indexed)

        Returns:
            Function call with arguments in R1C1 notation.
        """
        args_str = ",".join(arg.to_r1c1_string(base_row, base_col) for arg in self.args)
        return f"{self.name}({args_str})"


class UnaryNode(ASTNode):
    """Unary operator node"""

    def __init__(self, operator: str, operand: ASTNode):
        super().__init__("Unary")
        self.operator = operator  # "+", "-", "%"
        self.operand = operand

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": self.node_type,
            "operator": self.operator,
            "operand": self.operand.to_dict(),
        }

    def to_string(self) -> str:
        if self.operator == "%":
            # Postfix percentage
            return f"{self.operand.to_string()}%"
        else:
            # Prefix (+ or -)
            return f"{self.operator}{self.operand.to_string()}"

    def to_r1c1_string(self, base_row: int, base_col: int) -> str:
        """Convert unary operand to R1C1.

        Args:
            base_row: Row of cell containing formula (1-indexed)
            base_col: Column of cell containing formula (1-indexed)

        Returns:
            Unary expression with operand in R1C1 notation.
        """
        if self.operator == "%":
            return f"{self.operand.to_r1c1_string(base_row, base_col)}%"
        else:
            return f"{self.operator}{self.operand.to_r1c1_string(base_row, base_col)}"


class BinaryNode(ASTNode):
    """Binary operator node"""

    def __init__(self, operator: str, left: ASTNode, right: ASTNode):
        super().__init__("Binary")
        self.operator = operator  # "+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">="
        self.left = left
        self.right = right

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": self.node_type,
            "operator": self.operator,
            "left": self.left.to_dict(),
            "right": self.right.to_dict(),
        }

    def to_string(self) -> str:
        """Convert binary node to string with precedence-aware parenthesization."""
        # Get precedence of this operator
        my_precedence = FormulaParser.PRECEDENCE.get(self.operator, 0)

        # Convert left child
        left_str = self.left.to_string()
        # Add parens if left child is binary with lower precedence
        if isinstance(self.left, BinaryNode):
            left_precedence = FormulaParser.PRECEDENCE.get(self.left.operator, 0)
            if left_precedence < my_precedence:
                left_str = f"({left_str})"

        # Convert right child
        right_str = self.right.to_string()
        # Add parens if right child is binary with lower or equal precedence
        # Equal precedence matters for right side (e.g., A-B-C is (A-B)-C not A-(B-C))
        # Exception: exponentiation is right-associative
        if isinstance(self.right, BinaryNode):
            right_precedence = FormulaParser.PRECEDENCE.get(self.right.operator, 0)
            # For exponentiation (^), allow equal precedence on right (right-associative)
            if self.operator == "^":
                if right_precedence < my_precedence:
                    right_str = f"({right_str})"
            else:
                # For other operators, require strictly higher precedence on right
                if right_precedence <= my_precedence:
                    right_str = f"({right_str})"

        return f"{left_str}{self.operator}{right_str}"

    def to_r1c1_string(self, base_row: int, base_col: int) -> str:
        """Convert binary operands to R1C1 with precedence-aware parenthesization.

        Args:
            base_row: Row of cell containing formula (1-indexed)
            base_col: Column of cell containing formula (1-indexed)

        Returns:
            Binary expression with operands in R1C1 notation.
        """
        my_precedence = FormulaParser.PRECEDENCE.get(self.operator, 0)

        left_str = self.left.to_r1c1_string(base_row, base_col)
        if isinstance(self.left, BinaryNode):
            left_precedence = FormulaParser.PRECEDENCE.get(self.left.operator, 0)
            if left_precedence < my_precedence:
                left_str = f"({left_str})"

        right_str = self.right.to_r1c1_string(base_row, base_col)
        if isinstance(self.right, BinaryNode):
            right_precedence = FormulaParser.PRECEDENCE.get(self.right.operator, 0)
            if self.operator == "^":
                if right_precedence < my_precedence:
                    right_str = f"({right_str})"
            else:
                if right_precedence <= my_precedence:
                    right_str = f"({right_str})"

        return f"{left_str}{self.operator}{right_str}"


class FormulaParser:
    """
    Parse Excel formulas to structural AST.

    Supports:
    - Constants (numbers, strings, booleans, errors)
    - Cell/range references (A1, $A$1, A1:B10, Sheet1!A1)
    - Functions (SUM, VLOOKUP, etc.)
    - Operators (arithmetic, comparison, string concatenation)
    - Named ranges (captured as Ref nodes)
    - Structured references (Table1[Column])

    Does NOT support:
    - Formula evaluation (structural only)
    - Array formulas
    - Legacy formulas
    """

    # Operator precedence (higher value = binds tighter)
    #
    # Excel operator precedence rules (from lowest to highest):
    # 1. String concatenation (&)           - Lowest precedence
    # 2. Comparison (=, <>, <, >, <=, >=)   - Equal precedence, left-to-right
    # 3. Addition/Subtraction (+, -)        - Equal precedence, left-to-right
    # 4. Multiplication/Division (*, /)     - Equal precedence, left-to-right
    # 5. Exponentiation (^)                 - Highest precedence, RIGHT-TO-LEFT (right-associative)
    #
    # Associativity:
    # - Most operators are left-associative: A+B+C = (A+B)+C
    # - Exponentiation is RIGHT-associative: A^B^C = A^(B^C)
    #
    # Parenthesization rules (implemented in BinaryNode.to_string()):
    # - Left child: Add parens if child has LOWER precedence than parent
    # - Right child: Add parens if child has LOWER OR EQUAL precedence (except ^ which is right-associative)
    #
    # Examples:
    # - A+B*C   → A+B*C     (no parens: * higher than +)
    # - (A+B)*C → (A+B)*C   (parens required to override precedence)
    # - A-B-C   → A-B-C     (left-associative: (A-B)-C)
    # - A^B^C   → A^B^C     (right-associative: A^(B^C))
    #
    PRECEDENCE = {
        "&": 1,  # String concatenation
        "=": 2,
        "<>": 2,
        "<": 2,
        ">": 2,
        "<=": 2,
        ">=": 2,  # Comparison
        "+": 3,
        "-": 3,  # Addition/subtraction
        "*": 4,
        "/": 4,  # Multiplication/division
        "^": 5,  # Exponentiation (right-associative)
    }

    def __init__(self) -> None:
        self.pos = 0
        self.formula = ""
        self.length = 0

    def parse(self, formula: str) -> ASTNode:
        """
        Parse a formula string to AST.

        Args:
            formula: Excel formula (with or without leading =)

        Returns:
            Root AST node

        Raises:
            ValueError: If formula syntax is invalid
        """
        # Strip leading = if present
        if formula.startswith("="):
            formula = formula[1:]

        self.formula = formula.strip()
        self.length = len(self.formula)
        self.pos = 0

        if not self.formula:
            raise ValueError("Empty formula")

        return self._parse_expression()

    def _current_char(self) -> str | None:
        """Get current character without advancing"""
        if self.pos < self.length:
            return self.formula[self.pos]
        return None

    def _peek_char(self, offset: int = 1) -> str | None:
        """Peek at character at current position + offset"""
        pos = self.pos + offset
        if pos < self.length:
            return self.formula[pos]
        return None

    def _advance(self) -> None:
        """Move position forward by 1"""
        self.pos += 1

    def _skip_whitespace(self) -> None:
        """Skip whitespace characters"""
        while self.pos < self.length and self.formula[self.pos].isspace():
            self.pos += 1

    def _parse_expression(self, min_precedence: int = 0) -> ASTNode:
        """Parse expression with operator precedence"""
        self._skip_whitespace()

        # Parse left operand (primary or unary)
        left = self._parse_unary()

        # Parse binary operators
        while True:
            self._skip_whitespace()

            # Check for binary operator
            op = self._try_parse_operator()
            if op is None:
                break

            precedence = self.PRECEDENCE.get(op, 0)
            if precedence < min_precedence:
                # Put operator back
                self.pos -= len(op)
                break

            # Parse right operand with higher precedence
            right = self._parse_expression(precedence + 1)

            # Create binary node
            left = BinaryNode(op, left, right)

        return left

    def _try_parse_operator(self) -> str | None:
        """Try to parse a binary operator"""
        self._skip_whitespace()

        if self.pos >= self.length:
            return None

        # Try two-character operators first
        if self.pos + 1 < self.length:
            two_char = self.formula[self.pos : self.pos + 2]
            if two_char in ("<=", ">=", "<>"):
                self.pos += 2
                return two_char

        # Try single-character operators
        char = self.formula[self.pos]
        if char in ("+", "-", "*", "/", "^", "&", "=", "<", ">"):
            self.pos += 1
            return char

        return None

    def _parse_unary(self) -> ASTNode:
        """Parse unary expression (unary operator + operand)"""
        self._skip_whitespace()

        # Check for unary operators
        char = self._current_char()
        if char in ("+", "-"):
            self._advance()
            operand = self._parse_unary()
            return UnaryNode(char, operand)

        return self._parse_primary()

    def _parse_primary(self) -> ASTNode:
        """Parse primary expression (atom or parenthesized expression)"""
        self._skip_whitespace()

        char = self._current_char()
        if char is None:
            raise ValueError("Unexpected end of formula")

        node: ASTNode

        # Parenthesized expression
        if char == "(":
            self._advance()
            expr = self._parse_expression()
            self._skip_whitespace()
            if self._current_char() != ")":
                raise ValueError(f"Expected ')' at position {self.pos}")
            self._advance()
            node = expr

        # String literal
        elif char == '"':
            node = self._parse_string()

        # Number, boolean, error, function, or reference
        else:
            node = self._parse_atom()

        # Handle postfix operators (percentage)
        self._skip_whitespace()
        while self._current_char() == "%":
            self._advance()
            node = UnaryNode("%", node)

        return node

    def _parse_string(self) -> ConstNode:
        """Parse string literal"""
        self._advance()  # Skip opening quote

        chars = []
        while self.pos < self.length:
            char = self.formula[self.pos]
            if char == '"':
                # Check for escaped quote
                if self._peek_char() == '"':
                    chars.append('"')
                    self.pos += 2
                else:
                    self._advance()  # Skip closing quote
                    return ConstNode("".join(chars))
            else:
                chars.append(char)
                self._advance()

        raise ValueError("Unterminated string literal")

    def _parse_atom(self) -> ASTNode:
        """Parse atomic value (number, boolean, error, function, or reference)"""
        start = self.pos

        # Collect characters for atom
        chars = []
        while self.pos < self.length:
            char = self.formula[self.pos]

            # Stop at operators, whitespace, or delimiters
            if char in (
                "(",
                ")",
                ",",
                " ",
                "\t",
                "\n",
                "+",
                "-",
                "*",
                "/",
                "^",
                "&",
                "=",
                "<",
                ">",
                "%",
            ):
                break

            # Handle single-quoted identifiers (sheet names with spaces/special chars)
            # Excel uses single quotes: 'My Sheet'!A1
            if char == "'":
                chars.append(char)
                self._advance()
                # Collect until closing quote
                while self.pos < self.length:
                    char = self.formula[self.pos]
                    chars.append(char)
                    self._advance()
                    if char == "'":
                        # Check for escaped quote (two single quotes)
                        if self.pos < self.length and self.formula[self.pos] == "'":
                            chars.append("'")
                            self._advance()
                        else:
                            break  # End of quoted identifier
                continue

            # Handle sheet references with '!'
            if char == "!":
                chars.append(char)
                self._advance()
                # Continue parsing the reference part
                continue

            # Handle structured references with '[' and ']'
            if char == "[":
                chars.append(char)
                self._advance()
                # Collect until closing bracket
                while self.pos < self.length and self.formula[self.pos] != "]":
                    chars.append(self.formula[self.pos])
                    self._advance()
                if self.pos < self.length:
                    chars.append(self.formula[self.pos])  # Add closing bracket
                    self._advance()
                continue

            chars.append(char)
            self._advance()

        atom = "".join(chars)

        if not atom:
            raise ValueError(f"Expected atom at position {start}")

        # Check if function call
        self._skip_whitespace()
        if self._current_char() == "(":
            return self._parse_function(atom)

        # Try to parse as number
        try:
            if "." in atom or "e" in atom.lower():
                return ConstNode(float(atom))
            else:
                return ConstNode(int(atom))
        except ValueError:
            pass

        # Check for boolean
        if atom.upper() == "TRUE":
            return ConstNode(True)
        if atom.upper() == "FALSE":
            return ConstNode(False)

        # Check for error
        if atom.startswith("#"):
            return ConstNode(atom)

        # Must be a reference (cell, range, named range, or structured ref)
        return RefNode(atom)

    def _parse_function(self, name: str) -> FunctionNode:
        """Parse function call"""
        self._advance()  # Skip opening parenthesis

        args: list[ASTNode] = []

        self._skip_whitespace()

        # Handle empty argument list
        if self._current_char() == ")":
            self._advance()
            return FunctionNode(name, args)

        # Parse arguments
        while True:
            self._skip_whitespace()

            # Parse argument expression
            arg = self._parse_expression()
            args.append(arg)

            self._skip_whitespace()

            # Check for more arguments
            char = self._current_char()
            if char == ",":
                self._advance()
                continue
            elif char == ")":
                self._advance()
                break
            else:
                raise ValueError(f"Expected ',' or ')' at position {self.pos}")

        return FunctionNode(name, args)


_formula_ast_cache: dict[str, dict[str, Any]] = {}


def parse_formula(formula: str) -> dict[str, Any]:
    """
    Parse Excel formula to structural AST dictionary.

    Results are cached by formula text — callers must not mutate the returned dict.

    Args:
        formula: Excel formula (with or without leading =)

    Returns:
        Dictionary representation of AST

    Raises:
        ValueError: If formula syntax is invalid

    Example:
        >>> parse_formula("=A1+B1")
        {'type': 'Binary', 'operator': '+', 'left': {'type': 'Ref', 'ref': 'A1'}, 'right': {'type': 'Ref', 'ref': 'B1'}}
    """
    cached = _formula_ast_cache.get(formula)
    if cached is not None:
        return cached
    parser = FormulaParser()
    ast = parser.parse(formula)
    result = ast.to_dict()
    _formula_ast_cache[formula] = result
    return result
