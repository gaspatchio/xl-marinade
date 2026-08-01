# ABOUTME: Extract cell/range/external edges from normalized formulas for dependency graph
# ABOUTME: Implements range-first policy and handles dynamic/unresolved references

"""
Reference Extraction

Extracts dependency edges from normalized formulas without evaluation.

Edge types:
- internal: Single cell reference within workbook
- range: Rectangular range reference (stored as first-class edge, not expanded)
- external: External workbook, unresolved name, or dynamic function reference

Design reference: §6.4 of memory_efficient_extraction_architecture.md
"""

import re
from dataclasses import dataclass
from functools import lru_cache
from typing import Any

from .cell_identity import pack as pack_cell_id
from .formula_normalizer import (
    MAX_COL,
    MAX_ROW,
    FormulaContext,
    FormulaTokenizer,
    a1_to_col,
    parse_a1_ref,
)


@dataclass
class Edge:
    """Represents a dependency edge from a formula cell."""

    type: str  # 'internal', 'range', 'external'

    # For internal edges
    to_cell_id: int | None = None

    # For range edges
    to_sheet_id: int | None = None
    to_r1: int | None = None
    to_c1: int | None = None
    to_r2: int | None = None
    to_c2: int | None = None
    to_range_a1: str | None = None
    cell_count: int | None = None
    is_implicit_full_range: bool = False

    # Per-edge provenance (Issue #2). None/'static' for ordinary refs;
    # 'resolved_from_cache' for Issue #1 by-value INDIRECT/OFFSET edges that are
    # valid only for the cached scenario (snapshot-specific). Carried through
    # staging -> range_edges.provenance -> binding_edges.kind (range_dynamic).
    provenance: str | None = None

    # For external edges
    external_ref: str | None = None


class ReferenceExtractor:
    """
    Extract references from normalized formulas.

    Handles:
    - Internal cell references (A1 and R1C1 notation)
    - Range references (range-first policy)
    - 3D references (one edge per sheet)
    - External workbook references
    - Dynamic functions (INDIRECT, OFFSET, etc.)
    - Unresolved names
    - Structured table references
    """

    # Dynamic functions that create opaque references.
    # LET/LAMBDA are intentionally NOT here: their literal cell/range refs are
    # statically knowable. Bound names (LET declarations, LAMBDA params) are
    # blanked by _mask_let_lambda_bindings before extraction so they never
    # become edges, while real refs in value exprs / body / calc are preserved.
    # Functions whose *result target* is dynamic, in two classes:
    # - OPAQUE: address-computing functions. Their call is treated as a black
    #   box: one DYNAMIC marker edge, no raw ref extraction inside (the
    #   semantic-resolution layer owns them).
    # - TRANSPARENT lookups: INDEX/XLOOKUP/CHOOSE evaluate every argument
    #   (lookup keys, arrays, index expressions), so interior refs are genuine
    #   dependencies and are extracted like any other function's arguments —
    #   matching VLOOKUP/MATCH, which were never opaque. The DYNAMIC marker is
    #   still emitted so the resolution layer can narrow the lookup target.
    OPAQUE_DYNAMIC_FUNCTIONS = {"INDIRECT", "OFFSET", "GETPIVOTDATA", "CELL", "INFO"}
    TRANSPARENT_LOOKUP_FUNCTIONS = {"INDEX", "XLOOKUP", "CHOOSE"}
    DYNAMIC_FUNCTIONS = OPAQUE_DYNAMIC_FUNCTIONS | TRANSPARENT_LOOKUP_FUNCTIONS

    # Matches a DYNAMIC_FUNCTIONS name at an identifier boundary, immediately
    # followed (after optional whitespace) by '('. The leading boundary stops
    # substrings of longer identifiers (CELLAR, INDEXED) from matching. Longer
    # names sort first so e.g. INDIRECT wins over a hypothetical IND prefix.
    _DYNAMIC_NAME_PATTERN = re.compile(
        r"(?<![A-Za-z0-9_.])("
        + "|".join(sorted(DYNAMIC_FUNCTIONS, key=len, reverse=True))
        + r")\s*\(",
        re.IGNORECASE,
    )

    # R1C1 reference patterns.
    # Boundary guards: the lookbehind/lookahead prevent matching "RC" fragments
    # inside function names and identifiers (PERCENTILE, SEARCH, SOURCE_DATA, ...).
    # Each component is "R"/"C" alone (current row/col), "[n]" relative, or "n" absolute.
    _R1C1_REF = r"R(?:\[-?\d+\]|\d+)?C(?:\[-?\d+\]|\d+)?"
    _R1C1_PREFIX = r"(?<![A-Za-z0-9_])"
    _R1C1_SUFFIX = r"(?![A-Za-z0-9_(])"

    # Standalone cell refs require at least one explicit row/col component:
    # a completely bare "RC" would be a self-reference (circular), which the
    # normalizer never emits outside ranges or sheet-qualified refs.
    R1C1_CELL_PATTERN = re.compile(
        _R1C1_PREFIX
        + r"(?:R(?:\[-?\d+\]|\d+)C(?:\[-?\d+\]|\d+)?|RC(?:\[-?\d+\]|\d+))"
        + _R1C1_SUFFIX,
        re.IGNORECASE,
    )

    # Range endpoints may be bare "RC" (e.g., "RC:R[9]C" from =A1:A10 in A1).
    R1C1_RANGE_PATTERN = re.compile(
        rf"{_R1C1_PREFIX}{_R1C1_REF}:{_R1C1_REF}{_R1C1_SUFFIX}", re.IGNORECASE
    )

    THREE_D_PATTERN = re.compile(
        rf"(?:(?:'([^']+)')|([A-Z_][A-Z0-9_]*)):(?:(?:'([^']+)')|([A-Z_][A-Z0-9_]*))!({_R1C1_REF}(?::{_R1C1_REF})?){_R1C1_SUFFIX}",
        re.IGNORECASE,
    )

    SHEET_QUALIFIED_PATTERN = re.compile(
        rf"(?:(?:'([^']+)')|([A-Z_][A-Z0-9_]*))!({_R1C1_REF}(?::{_R1C1_REF})?){_R1C1_SUFFIX}",
        re.IGNORECASE,
    )

    # Double-quoted string literal (Excel escapes quotes by doubling: "say ""hi"""),
    # OR a single-quoted sheet-name segment (e.g. 'say "hi"'!A1). The single-quoted
    # alternative is matched first so a double quote inside a quoted sheet name is
    # never mistaken for the start of a string literal.
    STRING_LITERAL_PATTERN = re.compile(r"'[^']*'|\"(?:[^\"]|\"\")*\"")

    def __init__(
        self,
        sheet_name_map: dict[
            str, tuple[int, str]
        ],  # case-insensitive -> (sheet_id, canonical_name)
        defined_names: dict[str, str] = None,  # name -> resolved_ref (or formula)
        table_refs: dict[
            str, tuple[int, int, int, int, int]
        ] = None,  # table[col] -> (sheet_id, r1, c1, r2, c2)
        name_table_map: object | None = None,
        allow_row_col_ranges: bool = True,
    ):
        """
        Initialize reference extractor.

        Args:
            sheet_name_map: Map of sheet names (case-insensitive) to (sheet_id, canonical_name)
            defined_names: Map of defined names to resolved references (optional)
            table_refs: Map of structured references to ranges (optional)
        """
        self.sheet_name_map = sheet_name_map
        self.defined_names = defined_names or {}
        self.table_refs = table_refs or {}
        self.name_table_map = name_table_map
        self.allow_row_col_ranges = allow_row_col_ranges
        self._defined_name_lookup: dict[str, list[str]] = {}
        self._table_name_lookup: dict[str, str] = {}
        if self.name_table_map:
            try:
                for info in self.name_table_map.get_all_names():
                    key = info.name.lower()
                    self._defined_name_lookup.setdefault(key, []).append(info.name)
                for key in self._defined_name_lookup:
                    self._defined_name_lookup[key] = sorted(self._defined_name_lookup[key])
                for table in self.name_table_map.get_all_tables():
                    self._table_name_lookup[table.name.lower()] = table.name
            except Exception:
                self._defined_name_lookup = {}
                self._table_name_lookup = {}
        self.tokenizer = FormulaTokenizer()

    @lru_cache(maxsize=20000)
    def _tokenize_cached(self, formula: str) -> tuple[Any, ...]:
        # Reuse tokenization across repeated formulas to reduce parse overhead.
        return tuple(self.tokenizer.tokenize(formula))

    @classmethod
    def _mask_string_literals(cls, formula: str) -> str:
        """
        Blank out the contents of double-quoted string literals (length-preserving)
        so regex passes never extract references from text like "R2C5".
        Single-quoted sheet-name segments are kept verbatim so sheet names
        containing a double quote are not corrupted.
        """

        def _blank(m: re.Match) -> str:
            text = m.group()
            if text.startswith("'"):
                return text  # quoted sheet-name segment: keep verbatim
            return '"' + " " * (len(text) - 2) + '"'

        return cls.STRING_LITERAL_PATTERN.sub(_blank, formula)

    # LET(/LAMBDA( at an identifier boundary, optional _xlfn. prefix (the real
    # on-disk spelling). '.' in the leading boundary guard would block _xlfn.LET,
    # so the boundary deliberately excludes only word chars/'_' here and the
    # _xlfn. is consumed by the pattern itself.
    _LET_LAMBDA_NAME_PATTERN = re.compile(
        r"(?<![A-Za-z0-9_])(?:_xlfn\.)?(LET|LAMBDA)\s*\(",
        re.IGNORECASE,
    )
    _BARE_IDENT_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")

    @classmethod
    def _mask_let_lambda_bindings(cls, formula: str) -> str:
        """
        Blank out LET name declarations / LAMBDA parameter names AND their body
        usages (length-preserving, spaces) so a bound identifier never becomes a
        dependency edge. Real cell/range refs in value exprs, calc, and body are
        left untouched.

        LET(name1, value1, name2, value2, ..., calc): odd args (0-based even
        indices 0,2,...) are name declarations; the final arg is the calc.
        LAMBDA(param1, ..., paramN, body): all but the last arg are parameters.

        Run this AFTER _mask_string_literals so the comma/paren scan never trips
        on a comma or '(' inside a quoted string.
        """
        if not cls._LET_LAMBDA_NAME_PATTERN.search(formula):
            return formula

        chars = list(formula)
        n = len(formula)

        def blank_span(start: int, end: int) -> None:
            for k in range(start, end):
                if not chars[k].isspace():
                    chars[k] = " "

        pos = 0
        while pos < n:
            match = cls._LET_LAMBDA_NAME_PATTERN.search(formula, pos)
            if not match:
                break
            kind = match.group(1).upper()
            open_paren = match.end() - 1  # match.end() is just past '('

            # Walk paren-balanced from inside the '(' collecting top-level
            # comma-separated arg spans (start, end) relative to `formula`.
            args: list[tuple[int, int]] = []
            depth = 1
            i = open_paren + 1
            arg_start = i
            close_paren = -1
            while i < n and depth > 0:
                ch = formula[i]
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        close_paren = i
                        args.append((arg_start, i))
                        break
                elif ch == "," and depth == 1:
                    args.append((arg_start, i))
                    arg_start = i + 1
                i += 1

            if close_paren == -1:
                # Unbalanced (truncated formula): stop to avoid runaway.
                break

            # Determine which arg spans are declarations.
            if kind == "LET":
                decl_indices = range(0, max(len(args) - 1, 0), 2)
            else:  # LAMBDA
                decl_indices = range(0, max(len(args) - 1, 0))

            bound_names: set[str] = set()
            for idx in decl_indices:
                a_start, a_end = args[idx]
                text = formula[a_start:a_end].strip()
                if cls._BARE_IDENT_RE.match(text):
                    bound_names.add(text.lower())
                    blank_span(a_start, a_end)

            # Blank every word-boundary usage of each bound name within the
            # call scope [open_paren, close_paren], killing body/calc usages.
            # A name followed by '(' is a function call, not a defined-name
            # identifier, so it is left alone (harmless: never an edge).
            scope_text = formula[open_paren : close_paren + 1]
            for name in bound_names:
                usage_re = re.compile(
                    r"(?<![A-Za-z0-9_.!:])" + re.escape(name) + r"(?![A-Za-z0-9_(!:])",
                    re.IGNORECASE,
                )
                for um in usage_re.finditer(scope_text):
                    blank_span(open_paren + um.start(), open_paren + um.end())

            # Advance just past this LET/LAMBDA name so nested LET/LAMBDA inside
            # value/calc/body are processed by a later iteration.
            pos = match.end()

        return "".join(chars)

    def extract_edges(self, formula: str, ctx: FormulaContext) -> list[Edge]:
        """
        Extract all edges from a formula.

        Args:
            formula: Normalized R1C1 formula (or A1 for initial extraction)
            ctx: Formula context

        Returns:
            List of Edge objects
        """
        if not formula:
            return []

        edges = []

        # Mask string literal contents so quoted text like "R2C5" or "RC"
        # is never treated as a reference.
        masked = self._mask_string_literals(formula)

        # Blank LET name declarations / LAMBDA params and their body usages so a
        # bound identifier never becomes an edge. Runs on the string-masked text
        # so the arg splitter ignores commas/parens inside quoted strings. Real
        # cell/range refs in value exprs / calc / body are preserved verbatim.
        masked = self._mask_let_lambda_bindings(masked)

        # Check if formula contains R1C1 notation. Sheet-qualified refs are
        # checked separately because bare "RC" (e.g., Sheet2!RC) is legitimate
        # there but excluded from the standalone cell pattern.
        has_r1c1 = bool(
            self.R1C1_CELL_PATTERN.search(masked)
            or self.R1C1_RANGE_PATTERN.search(masked)
            or self.SHEET_QUALIFIED_PATTERN.search(masked)
        )

        if has_r1c1:
            # Extract R1C1 references directly using regex
            edges = self._extract_r1c1_edges(formula, masked, ctx)
            edges.extend(self._extract_identifier_edges_from_tokens(masked, ctx))
            return edges

        # Otherwise, use token-based extraction for A1 notation. Tokenize the
        # masked text so blanked LET/LAMBDA bound names become whitespace (never
        # edges) while real cell_ref/range_ref/sheet_ref token values are intact.
        tokens = self._tokenize_cached(masked)

        i = 0
        while i < len(tokens):
            token = tokens[i]

            # Dynamic functions: always record the DYNAMIC marker edge.
            if token.type == "function" and token.value.upper() in self.DYNAMIC_FUNCTIONS:
                func_call = self._extract_function_call(tokens, i)
                edges.append(Edge(type="external", external_ref=f"DYNAMIC:{func_call}"))
                if token.value.upper() in self.OPAQUE_DYNAMIC_FUNCTIONS:
                    # Opaque: skip past the call so its arguments emit no edges.
                    i = self._skip_function(tokens, i)
                else:
                    # Transparent lookup: descend into the arguments — lookup
                    # keys, arrays and index expressions are real dependencies.
                    i += 1
                continue

            # Sheet-qualified references
            if token.type == "sheet_ref":
                sheet_ref = token.value

                # Check for 3D reference (Sheet1:Sheet3!)
                if ":" in sheet_ref and "!" in sheet_ref:
                    # Parse 3D reference
                    edges.extend(self._extract_3d_reference(tokens, i, ctx))
                    i += 1
                    continue

                # Check for external workbook reference
                if "[" in sheet_ref:
                    # External reference
                    ref_str = self._extract_external_reference(tokens, i)
                    edges.append(Edge(type="external", external_ref=ref_str))
                    i += 1
                    continue

                # Regular sheet-qualified reference
                # Next token should be cell or range ref
                if i + 1 < len(tokens):
                    next_token = tokens[i + 1]

                    # Resolve sheet name
                    sheet_name = self._parse_sheet_name(sheet_ref)
                    sheet_info = self.sheet_name_map.get(sheet_name.lower())

                    if not sheet_info:
                        # Unresolved sheet
                        edges.append(
                            Edge(
                                type="external",
                                external_ref=f"UNRESOLVED:{sheet_ref}{next_token.value}",
                            )
                        )
                        i += 2
                        continue

                    target_sheet_id, _ = sheet_info

                    if next_token.type == "cell_ref":
                        # Single cell reference
                        edge = self._extract_cell_edge(next_token.value, target_sheet_id, ctx)
                        if edge:
                            edges.append(edge)
                        i += 2
                        continue

                    elif next_token.type in ["range_ref", "whole_col_ref", "whole_row_ref"]:
                        # Range reference
                        edge = self._extract_range_edge(next_token.value, target_sheet_id, ctx)
                        if edge:
                            edges.append(edge)
                        i += 2
                        continue

            # Unqualified cell references (same sheet)
            if token.type == "cell_ref":
                edge = self._extract_cell_edge(token.value, ctx.sheet_id, ctx)
                if edge:
                    edges.append(edge)

            # Unqualified range references (same sheet)
            elif token.type in ["range_ref", "whole_col_ref", "whole_row_ref"]:
                edge = self._extract_range_edge(token.value, ctx.sheet_id, ctx)
                if edge:
                    edges.append(edge)

            # Structured references
            elif token.type == "structured_ref":
                struct_edges = self._extract_structured_ref(token.value, ctx)
                if struct_edges:
                    edges.extend(struct_edges)

            elif token.type == "identifier":
                named_edges = self._extract_defined_name_edges(token.value, ctx)
                if named_edges:
                    edges.extend(named_edges)

            i += 1

        return edges

    def _extract_identifier_edges_from_tokens(
        self, formula: str, ctx: FormulaContext
    ) -> list[Edge]:
        """`formula` should have string literal contents already masked."""
        edges: list[Edge] = []
        # Tokenizer is A1-oriented and can misclassify R1C1 refs (e.g., "RC[-1]") as identifiers,
        # which would create spurious UNRESOLVED edges. Strip R1C1 refs before tokenization and
        # only extract identifiers/structured refs from the remaining text. Sheet-qualified
        # forms go first so bare "RC" refs (e.g., Sheet2!RC) are removed too.
        formula_wo_r1c1 = self.THREE_D_PATTERN.sub("", formula)
        formula_wo_r1c1 = self.SHEET_QUALIFIED_PATTERN.sub("", formula_wo_r1c1)
        formula_wo_r1c1 = self.R1C1_RANGE_PATTERN.sub("", formula_wo_r1c1)
        formula_wo_r1c1 = self.R1C1_CELL_PATTERN.sub("", formula_wo_r1c1)
        tokens = self._tokenize_cached(formula_wo_r1c1)
        i = 0
        while i < len(tokens):
            token = tokens[i]
            if token.type == "structured_ref":
                struct_edges = self._extract_structured_ref(token.value, ctx)
                if struct_edges:
                    edges.extend(struct_edges)
                i += 1
                continue

            if token.type == "identifier":
                named_edges = self._extract_defined_name_edges(token.value, ctx)
                if named_edges:
                    edges.extend(named_edges)
                i += 1
                continue

            if token.type == "sheet_ref" and i + 1 < len(tokens):
                next_token = tokens[i + 1]
                if next_token.type == "identifier":
                    sheet_name = self._parse_sheet_name(token.value)
                    sheet_info = self.sheet_name_map.get(sheet_name.lower())
                    if not sheet_info:
                        edges.append(
                            Edge(
                                type="external",
                                external_ref=f"UNRESOLVED:{token.value}{next_token.value}",
                            )
                        )
                        i += 2
                        continue

                    target_sheet_id, target_sheet_name = sheet_info
                    scoped_ctx = FormulaContext(
                        sheet_id=target_sheet_id,
                        row=ctx.row,
                        col=ctx.col,
                        sheet_name=target_sheet_name,
                    )
                    named_edges = self._extract_defined_name_edges(next_token.value, scoped_ctx)
                    if named_edges:
                        edges.extend(named_edges)
                    i += 2
                    continue

            i += 1

        return edges

    def _find_dynamic_spans(self, masked: str) -> list[tuple[int, int, bool]]:
        """Find paren-balanced spans of DYNAMIC function calls.

        Scans left to right; on each dynamic-function name at an identifier
        boundary, walks forward counting parenthesis depth to the true matching
        close paren (a regex cannot balance parens, so the old `[^)]*` truncated
        nested calls at the first ')'). Returns (start, end, is_opaque) per
        call. Opaque calls are not descended into — the whole call is a black
        box, matching the A1 token path's _skip_function behavior. Transparent
        lookup calls ARE descended into, so a nested opaque call (e.g. an
        OFFSET inside an INDEX argument) is still found and stripped.

        `masked` has string-literal contents blanked, so parens inside quoted
        text cannot perturb the depth count.
        """
        spans: list[tuple[int, int, bool]] = []
        pos = 0
        n = len(masked)
        while pos < n:
            match = self._DYNAMIC_NAME_PATTERN.search(masked, pos)
            if not match:
                break
            is_opaque = match.group(1).upper() in self.OPAQUE_DYNAMIC_FUNCTIONS
            # match.end() is just past the '(' that the pattern consumed.
            depth = 1
            i = match.end()
            while i < n and depth > 0:
                ch = masked[i]
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                i += 1
            if depth != 0:
                # Unbalanced (truncated formula); stop to avoid a runaway span.
                break
            spans.append((match.start(), i, is_opaque))
            # Skip past opaque calls entirely; continue scanning inside
            # transparent ones so nested dynamic calls are found.
            pos = i if is_opaque else match.end()
        return spans

    def _extract_r1c1_edges(self, formula: str, masked: str, ctx: FormulaContext) -> list[Edge]:
        """
        Extract edges from R1C1 formula using regex.

        Args:
            formula: R1C1 formula
            masked: Same formula with string literal contents blanked out
                (length-preserving, so spans map back to `formula`)
            ctx: Formula context

        Returns:
            List of edges
        """
        edges = []

        # Every dynamic call gets a DYNAMIC marker edge. Opaque calls
        # (INDIRECT/OFFSET/...) additionally suppress raw ref extraction from
        # their arguments; transparent lookups (INDEX/XLOOKUP/CHOOSE) keep
        # their arguments in the text so lookup keys, arrays and index
        # expressions still produce edges. Match against the masked text (so
        # string contents can't confuse matching) but record the original
        # call text via the span.
        opaque_spans: list[tuple[int, int]] = []
        for start, end, is_opaque in self._find_dynamic_spans(masked):
            edges.append(Edge(type="external", external_ref=f"DYNAMIC:{formula[start:end]}"))
            if is_opaque:
                opaque_spans.append((start, end))

        # Remove opaque calls from the formula to avoid extracting refs inside.
        formula_without_dynamic = masked
        for start, end in reversed(opaque_spans):
            formula_without_dynamic = (
                formula_without_dynamic[:start] + formula_without_dynamic[end:]
            )

        # Handle 3D references first (Sheet1:Sheet3!R1C1)
        three_d_matches = []
        for match in self.THREE_D_PATTERN.finditer(formula_without_dynamic):
            quoted_sheet1 = match.group(1)
            unquoted_sheet1 = match.group(2)
            quoted_sheet2 = match.group(3)
            unquoted_sheet2 = match.group(4)
            ref = match.group(5)

            sheet1_name = quoted_sheet1 if quoted_sheet1 else unquoted_sheet1
            sheet2_name = quoted_sheet2 if quoted_sheet2 else unquoted_sheet2

            sheet1_info = self.sheet_name_map.get(sheet1_name.lower())
            sheet2_info = self.sheet_name_map.get(sheet2_name.lower())

            if not sheet1_info or not sheet2_info:
                edges.append(Edge(type="external", external_ref=f"UNRESOLVED:{match.group()}"))
                three_d_matches.append(match.span())
                continue

            sheet1_id, _ = sheet1_info
            sheet2_id, _ = sheet2_info

            # Create one edge per sheet in the span
            for sheet_id in range(min(sheet1_id, sheet2_id), max(sheet1_id, sheet2_id) + 1):
                # Parse the reference part
                if ":" in ref:
                    edge = self._parse_r1c1_range(ref, ctx, sheet_id)
                else:
                    edge = self._parse_r1c1_cell(ref, ctx, sheet_id)

                if edge:
                    edges.append(edge)

            three_d_matches.append(match.span())

        # Remove 3D references from formula to avoid double-counting
        formula_without_3d = formula_without_dynamic
        for start, end in reversed(three_d_matches):
            formula_without_3d = formula_without_3d[:start] + formula_without_3d[end:]

        # Handle sheet-qualified references (Sheet2!R1C1)
        for match in self.SHEET_QUALIFIED_PATTERN.finditer(formula_without_3d):
            quoted_sheet = match.group(1)
            unquoted_sheet = match.group(2)
            ref = match.group(3)

            sheet_name = quoted_sheet if quoted_sheet else unquoted_sheet
            sheet_info = self.sheet_name_map.get(sheet_name.lower())

            if not sheet_info:
                edges.append(Edge(type="external", external_ref=f"UNRESOLVED:{match.group()}"))
                continue

            target_sheet_id, _ = sheet_info

            # Parse the reference part
            if ":" in ref:
                edge = self._parse_r1c1_range(ref, ctx, target_sheet_id)
            else:
                edge = self._parse_r1c1_cell(ref, ctx, target_sheet_id)

            if edge:
                edges.append(edge)

        # Remove sheet-qualified refs to avoid double-counting
        formula_without_sheets = self.SHEET_QUALIFIED_PATTERN.sub("", formula_without_3d)

        # Find all R1C1 range references (unqualified)
        for match in self.R1C1_RANGE_PATTERN.finditer(formula_without_sheets):
            range_ref = match.group()
            edge = self._parse_r1c1_range(range_ref, ctx, ctx.sheet_id)
            if edge:
                edges.append(edge)

        # Remove ranges from formula to avoid double-counting
        formula_without_ranges = self.R1C1_RANGE_PATTERN.sub("", formula_without_sheets)

        # Find all R1C1 cell references (unqualified)
        for match in self.R1C1_CELL_PATTERN.finditer(formula_without_ranges):
            cell_ref = match.group()
            edge = self._parse_r1c1_cell(cell_ref, ctx, ctx.sheet_id)
            if edge:
                edges.append(edge)

        return edges

    def _parse_r1c1_cell(self, ref: str, ctx: FormulaContext, target_sheet_id: int) -> Edge | None:
        """
        Parse R1C1 cell reference and create edge.

        Args:
            ref: R1C1 reference (e.g., "R1C1", "R[-1]C[2]")
            ctx: Formula context
            target_sheet_id: Target sheet ID

        Returns:
            Edge or None
        """
        try:
            row, col = self._resolve_r1c1_ref(ref, ctx)
            cell_id = pack_cell_id(target_sheet_id, row, col)

            return Edge(type="internal", to_cell_id=cell_id)
        except (ValueError, OverflowError):
            return None

    def _parse_r1c1_range(self, ref: str, ctx: FormulaContext, target_sheet_id: int) -> Edge | None:
        """
        Parse R1C1 range reference and create edge.

        Args:
            ref: R1C1 range (e.g., "R1C1:R10C2")
            ctx: Formula context
            target_sheet_id: Target sheet ID

        Returns:
            Edge or None
        """
        try:
            parts = ref.split(":")
            if len(parts) != 2:
                return None

            r1, c1 = self._resolve_r1c1_ref(parts[0], ctx)
            r2, c2 = self._resolve_r1c1_ref(parts[1], ctx)

            # Normalize range order
            if r1 > r2:
                r1, r2 = r2, r1
            if c1 > c2:
                c1, c2 = c2, c1

            # Check for 1x1 range
            if r1 == r2 and c1 == c2:
                cell_id = pack_cell_id(target_sheet_id, r1, c1)
                return Edge(type="internal", to_cell_id=cell_id)

            # Build canonical A1 range string
            from .cell_identity import col_to_a1

            range_a1 = f"{col_to_a1(c1)}{r1}:{col_to_a1(c2)}{r2}"

            cell_count = (r2 - r1 + 1) * (c2 - c1 + 1)

            return Edge(
                type="range",
                to_sheet_id=target_sheet_id,
                to_r1=r1,
                to_c1=c1,
                to_r2=r2,
                to_c2=c2,
                to_range_a1=range_a1,
                cell_count=cell_count,
            )
        except (ValueError, OverflowError):
            return None

    def _resolve_r1c1_ref(self, ref: str, ctx: FormulaContext) -> tuple[int, int]:
        """
        Resolve R1C1 reference to absolute (row, col).

        Args:
            ref: R1C1 reference (e.g., "R1C1", "R[-1]C[2]")
            ctx: Formula context

        Returns:
            Tuple of (row, col) as absolute 1-based coordinates
        """
        # Parse row part
        r_match = re.match(r"R(\[?-?\d+\]?)?", ref, re.IGNORECASE)
        if not r_match:
            raise ValueError(f"Invalid R1C1 reference: {ref}")

        r_part = r_match.group(1)
        if not r_part:
            # Just "R" means current row
            row = ctx.row
        elif r_part.startswith("[") and r_part.endswith("]"):
            # Relative reference
            delta = int(r_part[1:-1])
            row = ctx.row + delta
        else:
            # Absolute reference
            row = int(r_part)

        # Parse col part
        c_match = re.search(r"C(\[?-?\d+\]?)?", ref, re.IGNORECASE)
        if not c_match:
            raise ValueError(f"Invalid R1C1 reference: {ref}")

        c_part = c_match.group(1)
        if not c_part:
            # Just "C" means current col
            col = ctx.col
        elif c_part.startswith("[") and c_part.endswith("]"):
            # Relative reference
            delta = int(c_part[1:-1])
            col = ctx.col + delta
        else:
            # Absolute reference
            col = int(c_part)

        # Validate bounds
        if not (1 <= row <= MAX_ROW):
            raise ValueError(f"Row out of bounds: {row}")
        if not (1 <= col <= MAX_COL):
            raise ValueError(f"Col out of bounds: {col}")

        return row, col

    def _extract_cell_edge(
        self, ref: str, target_sheet_id: int, ctx: FormulaContext
    ) -> Edge | None:
        """Extract edge for a single cell reference."""
        try:
            row, col, _, _ = parse_a1_ref(ref)
            cell_id = pack_cell_id(target_sheet_id, row, col)

            return Edge(type="internal", to_cell_id=cell_id)
        except (ValueError, OverflowError):
            # Invalid reference
            return Edge(type="external", external_ref=f"UNRESOLVED:{ref}")

    def _extract_range_edge(
        self, ref: str, target_sheet_id: int, ctx: FormulaContext
    ) -> Edge | None:
        """Extract edge for a range reference."""
        try:
            # Parse range
            if ":" not in ref:
                # Single cell - treat as internal edge
                return self._extract_cell_edge(ref, target_sheet_id, ctx)

            parts = ref.split(":")
            if len(parts) != 2:
                return Edge(type="external", external_ref=f"UNRESOLVED:{ref}")

            part1 = parts[0].replace("$", "")
            part2 = parts[1].replace("$", "")

            # Whole-column reference (e.g., A:A)
            if part1.isalpha() and part2.isalpha():
                if not self.allow_row_col_ranges:
                    return None
                c1 = a1_to_col(part1)
                c2 = a1_to_col(part2)
                r1, r2 = 1, MAX_ROW
                is_implicit_full_range = True
            # Whole-row reference (e.g., 1:3)
            elif part1.isdigit() and part2.isdigit():
                if not self.allow_row_col_ranges:
                    return None
                r1 = int(part1)
                r2 = int(part2)
                c1, c2 = 1, MAX_COL
                is_implicit_full_range = True
            else:
                r1, c1, _, _ = parse_a1_ref(parts[0])
                r2, c2, _, _ = parse_a1_ref(parts[1])
                is_implicit_full_range = False

            # Normalize range order
            if r1 > r2:
                r1, r2 = r2, r1
            if c1 > c2:
                c1, c2 = c2, c1

            # Check for 1x1 range (normalize to cell edge)
            if r1 == r2 and c1 == c2:
                cell_id = pack_cell_id(target_sheet_id, r1, c1)
                return Edge(type="internal", to_cell_id=cell_id)

            # Build canonical A1 range string
            from .cell_identity import col_to_a1

            range_a1 = f"{col_to_a1(c1)}{r1}:{col_to_a1(c2)}{r2}"

            cell_count = (r2 - r1 + 1) * (c2 - c1 + 1)

            return Edge(
                type="range",
                to_sheet_id=target_sheet_id,
                to_r1=r1,
                to_c1=c1,
                to_r2=r2,
                to_c2=c2,
                to_range_a1=range_a1,
                cell_count=cell_count,
                is_implicit_full_range=is_implicit_full_range,
            )
        except (ValueError, OverflowError):
            return Edge(type="external", external_ref=f"UNRESOLVED:{ref}")

    def _extract_3d_reference(self, tokens: list, idx: int, ctx: FormulaContext) -> list[Edge]:
        """
        Extract edges for 3D reference (Sheet1:Sheet3!A1:B2).

        Returns one edge per sheet in the span.
        """
        edges = []

        token = tokens[idx]
        sheet_ref = token.value

        # Parse 3D reference: Sheet1:Sheet3!
        match = re.match(
            r"(?:\[([^\]]+)\])?(?:'([^']+)'|([A-Z_][A-Z0-9_]*)):(?:'([^']+)'|([A-Z_][A-Z0-9_]*))!",
            sheet_ref,
            re.IGNORECASE,
        )
        if not match:
            # Malformed 3D ref
            return [Edge(type="external", external_ref=f"UNRESOLVED:{sheet_ref}")]

        sheet1_name = match.group(2) or match.group(3)
        sheet2_name = match.group(4) or match.group(5)

        # Resolve sheet IDs
        sheet1_info = self.sheet_name_map.get(sheet1_name.lower())
        sheet2_info = self.sheet_name_map.get(sheet2_name.lower())

        if not sheet1_info or not sheet2_info:
            return [Edge(type="external", external_ref=f"UNRESOLVED:{sheet_ref}")]

        sheet1_id, _ = sheet1_info
        sheet2_id, _ = sheet2_info

        # Get the reference part (next token)
        if idx + 1 >= len(tokens):
            return [Edge(type="external", external_ref=f"UNRESOLVED:{sheet_ref}")]

        next_token = tokens[idx + 1]
        ref_str = next_token.value

        # Create one edge per sheet in the span
        for sheet_id in range(min(sheet1_id, sheet2_id), max(sheet1_id, sheet2_id) + 1):
            if next_token.type == "cell_ref":
                edge = self._extract_cell_edge(ref_str, sheet_id, ctx)
            else:
                edge = self._extract_range_edge(ref_str, sheet_id, ctx)

            if edge:
                edges.append(edge)

        return edges

    def _extract_external_reference(self, tokens: list, idx: int) -> str:
        """Extract external workbook reference string."""
        token = tokens[idx]
        sheet_ref = token.value

        # Get the reference part
        if idx + 1 < len(tokens):
            next_token = tokens[idx + 1]
            return f"{sheet_ref}{next_token.value}"

        return sheet_ref

    def _extract_structured_ref(self, ref: str, ctx: FormulaContext) -> list[Edge] | None:
        """
        Extract edge for structured table reference.

        For now, emit as UNRESOLVED (full table resolution requires table metadata).
        """
        if self.name_table_map:
            resolved = self.name_table_map.resolve_table_reference(ref)
            if not resolved and self._table_name_lookup:
                match = re.match(r"^(\w+)\[(.+)\]$", ref)
                if match:
                    table_name = match.group(1)
                    column_spec = match.group(2)
                    canonical = self._table_name_lookup.get(table_name.lower())
                    if canonical:
                        resolved = self.name_table_map.resolve_table_reference(
                            f"{canonical}[{column_spec}]"
                        )
            if resolved:
                edges = []
                for range_ref in resolved:
                    target_sheet_id, ref_part = self._resolve_sheet_ref(range_ref, ctx)
                    if target_sheet_id is None or not ref_part:
                        edges.append(Edge(type="external", external_ref=f"UNRESOLVED:{ref}"))
                        continue
                    edge = self._extract_range_edge(ref_part, target_sheet_id, ctx)
                    if edge:
                        edges.append(edge)
                return edges
        # Check if this is a row-level reference ([@Column])
        if ref.startswith("[@") and ref.endswith("]"):
            # Row-level reference - would need table metadata to resolve
            # For now, emit as unresolved
            return [Edge(type="external", external_ref=f"UNRESOLVED:{ref}")]

        # Column-level reference (Table[Column])
        # Would need table metadata to resolve to range
        return [Edge(type="external", external_ref=f"UNRESOLVED:{ref}")]

    def _extract_defined_name_edges(self, name: str, ctx: FormulaContext) -> list[Edge] | None:
        if not self.name_table_map:
            return None

        resolved = self.name_table_map.resolve_name(name, scope=ctx.sheet_name)
        if not resolved and self._defined_name_lookup:
            candidates = self._defined_name_lookup.get(name.lower(), [])
            for candidate in candidates:
                resolved = self.name_table_map.resolve_name(candidate, scope=ctx.sheet_name)
                if resolved:
                    break

        if not resolved:
            # Dynamic (formula-valued) names resolve to [] just like an unknown
            # name, so probe is_dynamic_name. Defined names are case-insensitive
            # in Excel, but is_dynamic_name keys are case-sensitive, so try the
            # registered candidate spelling (matched case-insensitively) the same
            # way the static-resolution path above rescues a case-mismatched ref.
            for probe in (name, *self._defined_name_lookup.get(name.lower(), [])):
                if self.name_table_map.is_dynamic_name(probe, scope=ctx.sheet_name):
                    return [Edge(type="external", external_ref=f"DYNAMIC:{probe}")]
            return None

        edges = []
        for range_ref in resolved:
            target_sheet_id, ref_part = self._resolve_sheet_ref(range_ref, ctx)
            if target_sheet_id is None or not ref_part:
                edges.append(Edge(type="external", external_ref=f"UNRESOLVED:{name}"))
                continue
            edge = self._extract_range_edge(ref_part, target_sheet_id, ctx)
            if edge:
                edges.append(edge)
        return edges

    def _resolve_sheet_ref(
        self, range_ref: str, ctx: FormulaContext
    ) -> tuple[int | None, str | None]:
        if "!" in range_ref:
            sheet_part, ref_part = range_ref.split("!", 1)
            sheet_part = sheet_part.strip("'").replace("''", "'")
            sheet_info = self.sheet_name_map.get(sheet_part.lower())
            if not sheet_info:
                return None, None
            return sheet_info[0], ref_part

        return ctx.sheet_id, range_ref

    def _parse_sheet_name(self, sheet_ref: str) -> str:
        """Extract sheet name from sheet reference."""
        # Remove trailing !
        if sheet_ref.endswith("!"):
            sheet_ref = sheet_ref[:-1]

        # Handle quoted names
        if sheet_ref.startswith("'") and sheet_ref.endswith("'"):
            return sheet_ref[1:-1].replace("''", "'")

        return sheet_ref

    def _extract_function_call(self, tokens: list, start_idx: int) -> str:
        """
        Extract the full function call as a string.

        Args:
            tokens: Token list
            start_idx: Index of function token

        Returns:
            Normalized function call string
        """
        result = [tokens[start_idx].value.upper()]

        # Find matching parentheses
        i = start_idx + 1
        if i >= len(tokens) or tokens[i].value != "(":
            return result[0]

        result.append("(")
        i += 1

        paren_depth = 1
        while i < len(tokens) and paren_depth > 0:
            token = tokens[i]

            if token.value == "(":
                paren_depth += 1
            elif token.value == ")":
                paren_depth -= 1

            if token.type != "whitespace":
                result.append(token.value)

            i += 1

        return "".join(result)

    def _skip_function(self, tokens: list, start_idx: int) -> int:
        """
        Skip past a function call.

        Returns the index after the closing parenthesis.
        """
        i = start_idx + 1
        if i >= len(tokens) or tokens[i].value != "(":
            return i

        i += 1
        paren_depth = 1

        while i < len(tokens) and paren_depth > 0:
            if tokens[i].value == "(":
                paren_depth += 1
            elif tokens[i].value == ")":
                paren_depth -= 1
            i += 1

        return i


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
