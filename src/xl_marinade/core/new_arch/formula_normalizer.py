# ABOUTME: Token-based formula parser and A1→R1C1 normalizer for memory-efficient extraction
# ABOUTME: Handles shared formulas, array/spill metadata, and whole-column/row expansion

"""
Formula Normalization

Converts Excel A1 formulas to canonical R1C1 notation for deduplication and storage.

Key features:
- Token-based parsing (no AST construction)
- Shared formula shifting (base formula + delta)
- Whole-column/row expansion to explicit bounds
- Function name uppercasing
- Whitespace normalization
- Sheet name canonicalization

Design reference: §6.4 of memory_efficient_extraction_architecture.md
"""

import re
from dataclasses import dataclass

# Excel limits
MAX_ROW = 1_048_576
MAX_COL = 16_384


@dataclass
class FormulaContext:
    """Context for formula normalization."""

    sheet_id: int
    row: int  # 1-based
    col: int  # 1-based
    sheet_name: str


class FormulaToken:
    """Represents a single token in a formula."""

    def __init__(self, type: str, value: str, start: int, end: int):
        self.type = type  # 'cell_ref', 'range_ref', 'function', 'operator', 'literal', 'whitespace', 'sheet_ref'
        self.value = value
        self.start = start
        self.end = end

    def __repr__(self):
        return f"Token({self.type}, {self.value!r})"


class FormulaTokenizer:
    """
    Tokenize Excel formulas without building an AST.

    Recognizes:
    - Cell references (A1, $A$1, Sheet1!A1, 'Sheet Name'!A1)
    - Range references (A1:B2, Sheet1!A1:B2)
    - 3D references (Sheet1:Sheet3!A1:B2)
    - Function names (SUM, VLOOKUP, etc.)
    - Operators (+, -, *, /, ^, &, =, <>, <, >, <=, >=)
    - Literals (numbers, strings, booleans, errors)
    - Structured references (Table[Column], [@Column])
    - External references ([Book.xlsx]Sheet!A1)
    """

    # Regex patterns
    CELL_REF_PATTERN = re.compile(r"\$?[A-Z]{1,3}\$?\d{1,7}", re.IGNORECASE)

    RANGE_REF_PATTERN = re.compile(
        r"\$?[A-Z]{1,3}\$?\d{1,7}:\$?[A-Z]{1,3}\$?\d{1,7}", re.IGNORECASE
    )

    WHOLE_COL_PATTERN = re.compile(r"\$?[A-Z]{1,3}:\$?[A-Z]{1,3}", re.IGNORECASE)

    WHOLE_ROW_PATTERN = re.compile(
        r"\$?\d{1,7}:\$?\d{1,7}",
    )

    FUNCTION_PATTERN = re.compile(r"[A-Z_][A-Z0-9_.]*(?=\()", re.IGNORECASE)

    IDENTIFIER_PATTERN = re.compile(r"[A-Z_][A-Z0-9_.]*", re.IGNORECASE)

    STRING_PATTERN = re.compile(r'"(?:[^"]|"")*"')

    NUMBER_PATTERN = re.compile(r"\d+\.?\d*(?:[eE][+-]?\d+)?")

    SHEET_REF_PATTERN = re.compile(
        r"(?:\[([^\]]+)\])?(?:'([^']+)'|([A-Z_][A-Z0-9_]*))!", re.IGNORECASE
    )

    @staticmethod
    def _extends_as_identifier(formula: str, end: int) -> bool:
        """True if the matched token continues with identifier characters."""
        return end < len(formula) and (formula[end].isalnum() or formula[end] in "_.")

    STRUCTURED_REF_PATTERN = re.compile(r"([A-Z_][A-Z0-9_]*)\[([^\]]+)\]", re.IGNORECASE)

    def tokenize(self, formula: str) -> list[FormulaToken]:
        """
        Tokenize a formula string.

        Args:
            formula: Formula text without leading '='

        Returns:
            List of FormulaToken objects
        """
        tokens = []
        i = 0

        while i < len(formula):
            # Skip whitespace (but track it for normalization)
            if formula[i].isspace():
                start = i
                while i < len(formula) and formula[i].isspace():
                    i += 1
                tokens.append(FormulaToken("whitespace", formula[start:i], start, i))
                continue

            # String literals
            if formula[i] == '"':
                match = self.STRING_PATTERN.match(formula, i)
                if match:
                    tokens.append(FormulaToken("literal", match.group(), i, match.end()))
                    i = match.end()
                    continue

            # Sheet references (including external workbooks and 3D refs)
            if i < len(formula) and (
                formula[i] == "["
                or formula[i] == "'"
                or (i + 1 < len(formula) and formula[i : i + 2] in ["'["])
            ):
                match = self.SHEET_REF_PATTERN.match(formula, i)
                if match:
                    _external_book = match.group(1)  # [Book.xlsx] part
                    quoted_sheet = match.group(2)  # 'Sheet Name' part
                    unquoted_sheet = match.group(3)  # Sheet part

                    _sheet_name = quoted_sheet if quoted_sheet else unquoted_sheet

                    # Check for 3D reference (Sheet1:Sheet3!)
                    end_pos = match.end()
                    if end_pos < len(formula) and formula[end_pos - 1] == "!":
                        # Look ahead for :Sheet!
                        colon_match = re.match(r"([^\s:]+)!", formula[end_pos:])
                        if (
                            colon_match
                            and i > 0
                            and formula[match.start() - 1 : match.start()] != ":"
                        ):
                            # Not a 3D ref, just a regular sheet ref
                            pass
                        else:
                            # Check if there's a : before the next !
                            next_part = formula[end_pos : end_pos + 50]
                            if ":" in next_part and "!" in next_part:
                                colon_idx = next_part.index(":")
                                excl_idx = next_part.index("!")
                                if colon_idx < excl_idx:
                                    # This is a 3D reference
                                    # Parse the second sheet name
                                    second_match = self.SHEET_REF_PATTERN.match(
                                        formula, end_pos + colon_idx + 1
                                    )
                                    if second_match:
                                        end_pos = second_match.end()

                    tokens.append(FormulaToken("sheet_ref", formula[i:end_pos], i, end_pos))
                    i = end_pos
                    continue

            # Spill operator (#) — only when it DIRECTLY follows a reference token
            # (A1#, A1:B2#, R1C1 ref). A leading '#' (error literals #REF!/#N/A)
            # falls through to the operator/literal branches below.
            if (
                formula[i] == "#"
                and tokens
                and tokens[-1].type
                in (
                    "cell_ref",
                    "range_ref",
                    "whole_col_ref",
                    "whole_row_ref",
                    "sheet_ref",
                    "structured_ref",
                    "identifier",
                )
            ):
                tokens.append(FormulaToken("spill_op", "#", i, i + 1))
                i += 1
                continue

            # Implicit-intersection operator (@) — only when it leads a
            # reference/identifier/'(' . The following ref/name token normalizes
            # itself; the '@' carries no target.
            if (
                formula[i] == "@"
                and i + 1 < len(formula)
                and (formula[i + 1].isalpha() or formula[i + 1] in "'$[(")
            ):
                tokens.append(FormulaToken("intersection_op", "@", i, i + 1))
                i += 1
                continue

            # Structured references (Table[Column])
            match = self.STRUCTURED_REF_PATTERN.match(formula, i)
            if match:
                tokens.append(FormulaToken("structured_ref", match.group(), i, match.end()))
                i = match.end()
                continue

            # Function names
            match = self.FUNCTION_PATTERN.match(formula, i)
            if match:
                tokens.append(FormulaToken("function", match.group(), i, match.end()))
                i = match.end()
                continue

            # Range references (must come before cell refs)
            match = self.RANGE_REF_PATTERN.match(formula, i)
            if (
                match
                and not self._extends_as_identifier(formula, match.end())
                and all(_is_valid_a1_ref(part) for part in match.group().split(":"))
            ):
                tokens.append(FormulaToken("range_ref", match.group(), i, match.end()))
                i = match.end()
                continue

            # Whole column references (A:A)
            match = self.WHOLE_COL_PATTERN.match(formula, i)
            if match:
                tokens.append(FormulaToken("whole_col_ref", match.group(), i, match.end()))
                i = match.end()
                continue

            # Whole row references (1:1)
            match = self.WHOLE_ROW_PATTERN.match(formula, i)
            if match:
                tokens.append(FormulaToken("whole_row_ref", match.group(), i, match.end()))
                i = match.end()
                continue

            # Cell references. A token that lexes like a ref but is not a
            # valid one (row 0, column past XFD) can only be a defined name -
            # Excel forbids names that collide with real refs, but allows
            # ref-lookalikes such as T0 or ZZZ1. Same if the token continues
            # with identifier characters (AB1x): let the identifier rule take
            # the whole token.
            match = self.CELL_REF_PATTERN.match(formula, i)
            if (
                match
                and not self._extends_as_identifier(formula, match.end())
                and _is_valid_a1_ref(match.group())
            ):
                tokens.append(FormulaToken("cell_ref", match.group(), i, match.end()))
                i = match.end()
                continue

            # Identifier (e.g., defined name)
            match = self.IDENTIFIER_PATTERN.match(formula, i)
            if match:
                tokens.append(FormulaToken("identifier", match.group(), i, match.end()))
                i = match.end()
                continue

            # Numbers
            match = self.NUMBER_PATTERN.match(formula, i)
            if match:
                tokens.append(FormulaToken("literal", match.group(), i, match.end()))
                i = match.end()
                continue

            # Operators and punctuation
            if i + 1 < len(formula) and formula[i : i + 2] in ["<>", "<=", ">="]:
                tokens.append(FormulaToken("operator", formula[i : i + 2], i, i + 2))
                i += 2
                continue

            if formula[i] in "+-*/^&=<>()[]{},:;%!":
                tokens.append(FormulaToken("operator", formula[i], i, i + 1))
                i += 1
                continue

            # Unknown character - treat as literal
            tokens.append(FormulaToken("literal", formula[i], i, i + 1))
            i += 1

        return tokens


def col_to_a1(col: int) -> str:
    """Convert 1-based column number to Excel column letters."""
    if not (1 <= col <= MAX_COL):
        raise ValueError(f"col out of bounds: {col} (must be 1..{MAX_COL})")

    result = []
    col -= 1

    while True:
        result.append(chr(ord("A") + (col % 26)))
        col //= 26
        if col == 0:
            break
        col -= 1

    return "".join(reversed(result))


def a1_to_col(col_str: str) -> int:
    """Convert Excel column letters to 1-based column number."""
    col_str = col_str.upper()

    if not col_str or not col_str.isalpha():
        raise ValueError(f"Invalid column string: {col_str}")

    col = 0
    for char in col_str:
        col = col * 26 + (ord(char) - ord("A") + 1)

    if not (1 <= col <= MAX_COL):
        raise ValueError(f"Column {col_str} out of bounds (col={col}, max={MAX_COL})")

    return col


def parse_a1_ref(ref: str) -> tuple[int, int, bool, bool]:
    """
    Parse A1 cell reference to (row, col, row_abs, col_abs).

    Args:
        ref: Cell reference like "A1", "$A$1", "B$2"

    Returns:
        Tuple of (row, col, row_absolute, col_absolute)
    """
    ref = ref.upper()

    # Parse $ markers
    col_abs = ref.startswith("$")
    if col_abs:
        ref = ref[1:]

    # Split into column and row
    i = 0
    while i < len(ref) and ref[i].isalpha():
        i += 1

    col_str = ref[:i]
    row_part = ref[i:]

    row_abs = row_part.startswith("$")
    if row_abs:
        row_part = row_part[1:]

    col = a1_to_col(col_str)
    row = int(row_part)

    if not (1 <= row <= MAX_ROW):
        raise ValueError(f"row out of bounds: {row}")

    return row, col, row_abs, col_abs


def _is_valid_a1_ref(ref: str) -> bool:
    """True if ref parses to a real cell (row 1..MAX_ROW, col 1..MAX_COL)."""
    try:
        parse_a1_ref(ref)
        return True
    except ValueError:
        return False


def a1_to_r1c1(ref: str, ctx: FormulaContext) -> str:
    """
    Convert A1 reference to R1C1 notation relative to context.

    Args:
        ref: A1 reference (e.g., "A1", "$A$1")
        ctx: Formula context (sheet, row, col)

    Returns:
        R1C1 string (e.g., "R1C1", "R[-1]C[2]")
    """
    row, col, row_abs, col_abs = parse_a1_ref(ref)

    # Build R1C1 (prefer compact RC form for zero deltas).
    if row_abs:
        r_part = f"R{row}"
    else:
        delta = row - ctx.row
        r_part = "R" if delta == 0 else f"R[{delta}]"

    if col_abs:
        c_part = f"C{col}"
    else:
        delta = col - ctx.col
        c_part = "C" if delta == 0 else f"C[{delta}]"

    return r_part + c_part


def expand_whole_col_ref(ref: str) -> str:
    """
    Expand whole-column reference to explicit bounds.

    Preserves absolute-column markers so $D:$D becomes $D$1:$D$1048576
    rather than D1:D1048576 — the latter loses column-absolute semantics
    and produces position-dependent R1C1 downstream.

    Args:
        ref: Whole-column reference (e.g., "A:A", "$B:$D")

    Returns:
        Expanded range (e.g., "A1:A1048576", "$B$1:$D$1048576")
    """
    parts = ref.split(":")
    if len(parts) != 2:
        raise ValueError(f"Invalid whole-column reference: {ref}")

    def expand_end(col_part: str, row: int) -> str:
        col_abs = col_part.startswith("$")
        col_letters = col_part[1:] if col_abs else col_part
        col_marker = "$" if col_abs else ""
        # When column is absolute, the synthesized row anchors should also be
        # absolute so a1_to_r1c1 emits row-absolute R1C1 (Rn rather than R[d]).
        row_marker = "$" if col_abs else ""
        return f"{col_marker}{col_letters}{row_marker}{row}"

    return f"{expand_end(parts[0], 1)}:{expand_end(parts[1], MAX_ROW)}"


def expand_whole_row_ref(ref: str) -> str:
    """
    Expand whole-row reference to explicit bounds.

    Preserves absolute-row markers so $2:$5 becomes $A$2:$XFD$5 rather
    than A2:XFD5 — the latter loses row-absolute semantics and produces
    position-dependent R1C1 downstream.

    Args:
        ref: Whole-row reference (e.g., "1:1", "$2:$5")

    Returns:
        Expanded range (e.g., "A1:XFD1", "$A$2:$XFD$5")
    """
    parts = ref.split(":")
    if len(parts) != 2:
        raise ValueError(f"Invalid whole-row reference: {ref}")

    max_col_str = col_to_a1(MAX_COL)

    def expand_end(row_part: str, col_letters: str) -> str:
        row_abs = row_part.startswith("$")
        row_digits = row_part[1:] if row_abs else row_part
        row_marker = "$" if row_abs else ""
        # When row is absolute, synthesized col anchors are also absolute so
        # a1_to_r1c1 emits col-absolute R1C1 (Cn rather than C[d]).
        col_marker = "$" if row_abs else ""
        return f"{col_marker}{col_letters}{row_marker}{row_digits}"

    return f"{expand_end(parts[0], 'A')}:{expand_end(parts[1], max_col_str)}"


def normalize_formula(formula: str, ctx: FormulaContext, sheet_name_map: dict[str, str]) -> str:
    """
    Normalize formula to R1C1 notation.

    Args:
        formula: Formula text without leading '='
        ctx: Formula context
        sheet_name_map: Map of case-insensitive sheet names to canonical names

    Returns:
        Normalized R1C1 formula
    """
    if not formula:
        return ""

    tokenizer = FormulaTokenizer()
    tokens = tokenizer.tokenize(formula)

    result = []
    i = 0

    while i < len(tokens):
        token = tokens[i]

        if token.type == "whitespace":
            # Normalize to single space (or omit if between certain operators)
            if result and result[-1] not in ["(", ",", "["]:
                result.append(" ")

        elif token.type == "function":
            # Uppercase function names
            result.append(token.value.upper())

        elif token.type == "cell_ref":
            # Convert to R1C1
            result.append(a1_to_r1c1(token.value, ctx))

        elif token.type == "range_ref":
            # Convert range to R1C1
            parts = token.value.split(":")
            r1c1_start = a1_to_r1c1(parts[0], ctx)
            r1c1_end = a1_to_r1c1(parts[1], ctx)
            result.append(f"{r1c1_start}:{r1c1_end}")

        elif token.type == "whole_col_ref":
            # Expand and convert
            expanded = expand_whole_col_ref(token.value)
            parts = expanded.split(":")
            r1c1_start = a1_to_r1c1(parts[0], ctx)
            r1c1_end = a1_to_r1c1(parts[1], ctx)
            result.append(f"{r1c1_start}:{r1c1_end}")

        elif token.type == "whole_row_ref":
            # Expand and convert
            expanded = expand_whole_row_ref(token.value)
            parts = expanded.split(":")
            r1c1_start = a1_to_r1c1(parts[0], ctx)
            r1c1_end = a1_to_r1c1(parts[1], ctx)
            result.append(f"{r1c1_start}:{r1c1_end}")

        elif token.type == "sheet_ref":
            # Normalize sheet name (case-insensitive lookup)
            # For now, pass through (reference extraction will handle)
            result.append(token.value)

        elif token.type == "structured_ref":
            # Pass through (reference extraction will resolve)
            result.append(token.value)

        elif token.type == "identifier":
            # Pass through (reference extraction resolves defined names)
            result.append(token.value)

        elif token.type in ("spill_op", "intersection_op"):
            # Spill anchor (A1#) and implicit intersection (@A1) carry no extra
            # target: the adjacent reference token already normalizes to the
            # anchor/operand. Drop the marker so the R1C1 string stays clean.
            pass

        else:
            # Literals, operators, etc. - pass through
            result.append(token.value)

        i += 1

    return "".join(result).strip()


def shift_shared_formula(
    base_formula: str,
    base_row: int,
    base_col: int,
    target_row: int,
    target_col: int,
    ctx: FormulaContext,
    sheet_name_map: dict[str, str],
) -> str:
    """
    Shift a shared formula from base cell to target cell.

    Shared formulas are stored once at the master cell and applied to
    a range. Each dependent cell gets the formula shifted by its offset
    from the master.

    Args:
        base_formula: Formula at master cell (A1 notation, no leading '=')
        base_row: Master cell row (1-based)
        base_col: Master cell col (1-based)
        target_row: Target cell row (1-based)
        target_col: Target cell col (1-based)
        ctx: Formula context for target cell
        sheet_name_map: Sheet name mapping

    Returns:
        Normalized R1C1 formula for target cell
    """
    # Shared formulas store the master formula at the base cell. We must shift
    # relative A1 references (no $) by the row/col delta before normalizing.
    row_delta = target_row - base_row
    col_delta = target_col - base_col

    if row_delta == 0 and col_delta == 0:
        return normalize_formula(base_formula, ctx, sheet_name_map)

    tokenizer = FormulaTokenizer()
    tokens = tokenizer.tokenize(base_formula)
    shifted_tokens: list[str] = []

    def format_cell(r: int, c: int, row_abs: bool, col_abs: bool) -> str:
        col_str = col_to_a1(c)
        row_str = str(r)
        if col_abs:
            col_str = f"${col_str}"
        if row_abs:
            row_str = f"${row_str}"
        return f"{col_str}{row_str}"

    def shift_cell_ref(ref: str) -> str:
        row, col, row_abs, col_abs = parse_a1_ref(ref)
        if not row_abs:
            row += row_delta
        if not col_abs:
            col += col_delta
        return format_cell(row, col, row_abs, col_abs)

    def shift_range_ref(ref: str) -> str:
        start_ref, end_ref = ref.split(":", 1)
        start_shifted = shift_cell_ref(start_ref)
        end_shifted = shift_cell_ref(end_ref)
        return f"{start_shifted}:{end_shifted}"

    for token in tokens:
        if token.type == "cell_ref":
            shifted_tokens.append(shift_cell_ref(token.value))
        elif token.type == "range_ref":
            shifted_tokens.append(shift_range_ref(token.value))
        else:
            shifted_tokens.append(token.value)

    shifted_formula = "".join(shifted_tokens)
    return normalize_formula(shifted_formula, ctx, sheet_name_map)


def shift_shared_formula_a1(
    base_formula: str,
    base_row: int,
    base_col: int,
    target_row: int,
    target_col: int,
) -> str:
    """
    Shift a shared formula from base cell to target cell, returning A1 notation.

    Like shift_shared_formula but returns the shifted A1 formula instead of
    converting to R1C1. Used when the caller needs per-cell A1 for INDIRECT
    resolution.
    """
    row_delta = target_row - base_row
    col_delta = target_col - base_col

    if row_delta == 0 and col_delta == 0:
        return base_formula

    tokenizer = FormulaTokenizer()
    tokens = tokenizer.tokenize(base_formula)
    shifted_tokens: list[str] = []

    def shift_cell_ref(ref: str) -> str:
        row, col, row_abs, col_abs = parse_a1_ref(ref)
        if not row_abs:
            row += row_delta
        if not col_abs:
            col += col_delta
        col_str = col_to_a1(col)
        row_str = str(row)
        if col_abs:
            col_str = f"${col_str}"
        if row_abs:
            row_str = f"${row_str}"
        return f"{col_str}{row_str}"

    def shift_range_ref(ref: str) -> str:
        start_ref, end_ref = ref.split(":", 1)
        return f"{shift_cell_ref(start_ref)}:{shift_cell_ref(end_ref)}"

    for token in tokens:
        if token.type == "cell_ref":
            shifted_tokens.append(shift_cell_ref(token.value))
        elif token.type == "range_ref":
            shifted_tokens.append(shift_range_ref(token.value))
        else:
            shifted_tokens.append(token.value)

    return "".join(shifted_tokens)
