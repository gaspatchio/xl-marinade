# ABOUTME: Binding detection and identity computation for contiguous rectangular ranges.
# ABOUTME: Computes cells_structure_hash, binding_id, and collapses cell edges to binding edges.

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass

from xl_marinade.core.ref_converter import (
    convert_formula_to_r1c1,
    parse_cell_address,
    quote_sheet_name,
)


@dataclass
class Binding:
    """
    Contiguous rectangular range with computed identities.

    Attributes:
        binding_id: SHA-256 hash (hex64) uniquely identifying this binding
        debug_label: Human-readable identifier (optional)
        sheet: Sheet name
        address_a1: Full A1 range (e.g., "B5:D10")
        top_left_a1: Top-left cell address (e.g., "B5")
        shape_rows: Number of rows in binding
        shape_cols: Number of columns in binding
        cells_structure_hash: Hash of sorted (address, formula) tuples
        cells: List of cell addresses in this binding
    """

    binding_id: str
    debug_label: str | None
    sheet: str
    address_a1: str
    top_left_a1: str
    shape_rows: int
    shape_cols: int
    binding_type: str
    cells_structure_hash: str
    cells: list[str]
    label_candidates_json: str | None = "{}"
    relationships_json: str | None = "{}"
    extraction_source: str | None = "standard"
    spatial_candidates_json: str | None = "{}"


@dataclass
class BindingEdge:
    """
    Dependency edge between two bindings.

    Attributes:
        from_binding_id: Source binding ID
        to_binding_id: Target binding ID
    """

    from_binding_id: str
    to_binding_id: str


def _parse_a1_address(address: str) -> tuple[str, int, int]:
    """
    Parse A1 address into components.

    Args:
        address: A1 address (e.g., "Sheet1!B5" or "B5")

    Returns:
        Tuple of (sheet, row, col)
    """
    parsed = parse_cell_address(address)
    sheet = str(parsed.get("sheet", ""))
    row = int(parsed.get("row", 0))
    col = int(parsed.get("col", 0))
    return (sheet, row, col)


def _col_to_letter(col: int) -> str:
    """
    Convert column number to letter(s).

    Args:
        col: Column number (1-indexed, A=1, B=2, etc.)

    Returns:
        Column letter(s) (e.g., "A", "Z", "AA", "AB")
    """
    result = ""
    while col > 0:
        col -= 1
        result = chr(col % 26 + ord("A")) + result
        col //= 26
    return result


def _format_a1_range(
    sheet: str, top_row: int, top_col: int, bottom_row: int, bottom_col: int
) -> str:
    """
    Format range as A1 notation.

    Args:
        sheet: Sheet name
        top_row: Top row (1-indexed)
        top_col: Top column (1-indexed)
        bottom_row: Bottom row (1-indexed)
        bottom_col: Bottom column (1-indexed)

    Returns:
        A1 range string (e.g., "Sheet1!B5:D10")
    """
    top_left = f"{_col_to_letter(top_col)}{top_row}"
    quoted_sheet = quote_sheet_name(sheet)

    if top_row == bottom_row and top_col == bottom_col:
        # Single cell
        return f"{quoted_sheet}!{top_left}"

    bottom_right = f"{_col_to_letter(bottom_col)}{bottom_row}"
    return f"{quoted_sheet}!{top_left}:{bottom_right}"


def _get_r1c1_signature(cell_addr: str, formula_a1: str) -> str:
    """
    Get R1C1 formula signature for pattern matching (IR Spec §4.5).

    Args:
        cell_addr: Cell address (e.g., "Sheet1!B5")
        formula_a1: Formula in A1 notation (e.g., "=A5+B4")

    Returns:
        R1C1 formula signature (e.g., "=RC[-1]+R[-1]C")
        Empty string for non-formula cells.

    Note: Cells with identical R1C1 signatures belong to same formula block.
          Uses existing convert_formula_to_r1c1() from ref_converter module.
    """
    if not formula_a1 or not formula_a1.strip():
        return ""

    # Parse cell address to get row/col
    parsed = parse_cell_address(cell_addr)
    row = parsed.get("row", 0)
    col = parsed.get("col", 0)

    if row == 0 or col == 0:
        return ""

    # Ensure row and col are ints (type narrowing for mypy)
    assert isinstance(row, int) and isinstance(col, int)

    # Use existing conversion function
    return convert_formula_to_r1c1(formula_a1, row, col)


# H3 (RC1): whole-extent reference canonicalisation for grouping signatures.
_EXCEL_MAX_ROWS = 1048576
_EXCEL_MAX_COLS = 16384
_R1C1_RANGE_RE = re.compile(
    r"R(\[-?\d+\]|-?\d+)?C(\[-?\d+\]|-?\d+)?:R(\[-?\d+\]|-?\d+)?C(\[-?\d+\]|-?\d+)?"
)


def _r1c1_offset(part: str | None) -> int | None:
    """Numeric value of an R1C1 row/col part ("[-8]" -> -8, "5" -> 5, "" -> None)."""
    if not part:
        return None
    return int(part[1:-1]) if part.startswith("[") else int(part)


def _canonicalize_whole_extent_refs(r1c1_sig: str) -> str:
    """Collapse whole-column / whole-row R1C1 ranges to a position-invariant form
    so fill-down (or fill-across) cells whose formula references an ENTIRE column
    ('Sheet'!C:C) or row share ONE grouping signature.

    Excel serialises a whole-column A1 ref to a row-RELATIVE R1C1 range whose row
    offsets shift with the base row (R[1-base]C[x]:R[1048576-base]C[x]); each cell
    then gets a distinct signature and the column shatters into N 1x1 bindings
    (RC1 fragmentation). The row span of such a range is always
    ``_EXCEL_MAX_ROWS - 1`` regardless of the base row, so it is detectable
    without knowing the base row; dropping the (varying) row part yields a
    base-invariant signature. This rewrites ONLY the signature used as a grouping
    key — never the stored canonical formula, edges, or families.
    """
    if not r1c1_sig or ":" not in r1c1_sig:
        return r1c1_sig

    def repl(m: re.Match) -> str:
        r1, c1, r2, c2 = m.group(1), m.group(2), m.group(3), m.group(4)
        rv1, rv2 = _r1c1_offset(r1), _r1c1_offset(r2)
        cv1, cv2 = _r1c1_offset(c1), _r1c1_offset(c2)
        # whole column: full row span, identical column part on both ends
        if rv1 is not None and rv2 is not None and (rv2 - rv1) == _EXCEL_MAX_ROWS - 1 and c1 == c2:
            return f"C{c1 or ''}:C{c2 or ''}"
        # whole row: full column span, identical row part on both ends
        if cv1 is not None and cv2 is not None and (cv2 - cv1) == _EXCEL_MAX_COLS - 1 and r1 == r2:
            return f"R{r1 or ''}:R{r2 or ''}"
        return m.group(0)

    return _R1C1_RANGE_RE.sub(repl, r1c1_sig)


def _cells_form_rectangle(cells: list[tuple[int, int, str]]) -> bool:
    """
    Check if cells form a perfect contiguous rectangle.

    Args:
        cells: List of (row, col, address) tuples

    Returns:
        True if cells form rectangle with no gaps, False otherwise

    Algorithm:
        1. Find bounding box (min/max row/col)
        2. Check if cell count == (rows * cols)
        3. Verify all positions in bounding box are present
    """
    if not cells:
        return False

    rows = [r for r, c, _ in cells]
    cols = [c for r, c, _ in cells]

    min_row, max_row = min(rows), max(rows)
    min_col, max_col = min(cols), max(cols)

    expected_count = (max_row - min_row + 1) * (max_col - min_col + 1)

    if len(cells) != expected_count:
        return False

    # Verify all positions present
    cell_positions = {(r, c) for r, c, _ in cells}
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            if (r, c) not in cell_positions:
                return False

    return True


def _find_contiguous_rectangles(
    cells: list[tuple[int, int, str]],
) -> list[list[tuple[int, int, str]]]:
    """
    Partition cells into maximal contiguous rectangular regions.

    Uses flood-fill to find connected components, then validates each is rectangular.

    Args:
        cells: List of (row, col, address) tuples with same R1C1 signature

    Returns:
        List of cell groups, each forming a contiguous rectangle

    Algorithm:
        1. Build adjacency graph (cells touching horizontally or vertically)
        2. Find connected components via flood-fill
        3. For each component, validate it forms a rectangle
        4. If rectangular, keep as block; else split into individual cells
    """
    if not cells:
        return []

    # Build position map for O(1) adjacency lookup
    cell_positions = {(r, c): (r, c, addr) for r, c, addr in cells}

    # Find connected components
    visited = set()
    components = []

    for row, col, _addr in sorted(cells):  # Deterministic order
        if (row, col) in visited:
            continue

        # Flood-fill to find connected component
        component = []
        stack = [(row, col)]

        while stack:
            r, c = stack.pop()
            if (r, c) in visited:
                continue
            if (r, c) not in cell_positions:
                continue

            visited.add((r, c))
            component.append(cell_positions[(r, c)])

            # Check 4 adjacent cells (up, down, left, right)
            for dr, dc in [(-1, 0), (1, 0), (0, -1), (0, 1)]:
                adj = (r + dr, c + dc)
                if adj in cell_positions and adj not in visited:
                    stack.append(adj)

        if component:
            components.append(component)

    # Validate each component is rectangular
    rectangles = []
    for component in components:
        if _cells_form_rectangle(component):
            rectangles.append(component)
        else:
            # Non-rectangular component (jagged edges / gaps): decompose into
            # maximal row-span sub-rectangles instead of emitting one binding
            # per cell. Singletonising same-formula cells here produces
            # thousands of fragmented bindings for real-world models.
            addr_by_pos = {(r, c): addr for r, c, addr in component}
            subrects = _split_positions_into_rectangles([(r, c) for r, c, _ in component])
            for min_r, max_r, min_c, max_c in subrects:
                sub_cells = [
                    (r, c, addr_by_pos[(r, c)])
                    for r in range(min_r, max_r + 1)
                    for c in range(min_c, max_c + 1)
                    if (r, c) in addr_by_pos
                ]
                if sub_cells:
                    rectangles.append(sub_cells)

    return rectangles


def _split_positions_into_rectangles(
    positions: list[tuple[int, int]],
) -> list[tuple[int, int, int, int]]:
    """Split a set of (row, col) positions into maximal row-span rectangles.

    Identical logic to grouping_native._split_component_into_rectangles but
    co-located here to avoid an import cycle. Processes rows top-to-bottom,
    tracking active (start_col, end_col) column-runs; extends a run vertically
    while the next row matches exactly, otherwise finalises it.
    """
    if not positions:
        return []

    row_to_cols: dict[int, list[int]] = {}
    for row, col in positions:
        row_to_cols.setdefault(row, []).append(col)

    rectangles: list[tuple[int, int, int, int]] = []
    active: dict[tuple[int, int], tuple[int, int]] = {}

    for row in sorted(row_to_cols):
        cols = sorted(row_to_cols[row])
        runs: list[tuple[int, int]] = []
        start = prev = cols[0]
        for col in cols[1:]:
            if col == prev + 1:
                prev = col
                continue
            runs.append((start, prev))
            start = prev = col
        runs.append((start, prev))

        current_keys: set[tuple[int, int]] = set()
        for start_col, end_col in runs:
            key = (start_col, end_col)
            if key in active:
                min_row, max_row = active[key]
                if max_row == row - 1:
                    active[key] = (min_row, row)
                else:
                    rectangles.append((min_row, max_row, start_col, end_col))
                    active[key] = (row, row)
            else:
                active[key] = (row, row)
            current_keys.add(key)

        for key in list(active.keys()):
            if key not in current_keys:
                min_row, max_row = active.pop(key)
                rectangles.append((min_row, max_row, key[0], key[1]))

    for key, (min_row, max_row) in active.items():
        rectangles.append((min_row, max_row, key[0], key[1]))

    rectangles.sort()
    return rectangles


def _detect_contiguous_blocks(
    cells: list[tuple[int, int, str]],
    cell_formulas: dict[str, str],
    formula_is_r1c1: bool = False,
    cell_signatures: dict[str, str] | None = None,
) -> list[list[str]]:
    """
    Group cells into maximal contiguous rectangular blocks (IR Spec §4.5).

    Args:
        cells: List of (row, col, address) tuples for single sheet
        cell_formulas: Mapping from cell address to formula

    Returns:
        List of cell address groups, each representing a binding

    Algorithm:
        1. Group cells by R1C1 formula signature
        2. For each formula group, find contiguous rectangular regions
        3. Use connected-component detection for multiple separate rectangles
        4. Return all blocks (multi-cell rectangles + remaining 1x1 cells)

    Determinism:
        - Process in sorted (row, col) order
        - Flood-fill is deterministic (sorted stack processing)
        - Final blocks sorted by top-left position
    """
    blocks = []

    # Group by R1C1 signature.
    # Prefer caller-provided canonical signatures (e.g., formulas.formula_r1c1
    # from the IR DB) over re-deriving R1C1 per cell — re-derivation has a
    # parser bug for mixed absolute/relative range refs that produces a
    # distinct signature per cell and singletonises whole rows.
    formula_groups: dict[str, list[tuple[int, int, str]]] = {}
    for row, col, cell_addr in cells:
        sig = cell_signatures.get(cell_addr) if cell_signatures is not None else None
        if sig:
            r1c1_sig = sig
        else:
            formula = cell_formulas.get(cell_addr, "")
            if formula:
                if formula_is_r1c1:
                    r1c1_sig = formula
                else:
                    # Use row/col directly (already parsed) instead of re-parsing cell_addr
                    r1c1_sig = (
                        convert_formula_to_r1c1(formula, row, col)
                        if row and col
                        else _get_r1c1_signature(cell_addr, formula)
                    )
            else:
                # Value-only cells: each gets unique signature (won't group)
                # Design decision: Only formulas group (by R1C1 signature).
                # This keeps constants as 1×1 bindings for granular tracking.
                # TODO: Revisit if assumption tables should group as structural units
                r1c1_sig = f"__value_only_{cell_addr}"

        # H3 (RC1): canonicalise whole-column/row refs so a fill-down column that
        # references an entire column shares one signature instead of shattering
        # into N 1x1 bindings. Skip the value-only sentinel (unique by design).
        if not r1c1_sig.startswith("__value_only_"):
            r1c1_sig = _canonicalize_whole_extent_refs(r1c1_sig)

        if r1c1_sig not in formula_groups:
            formula_groups[r1c1_sig] = []
        formula_groups[r1c1_sig].append((row, col, cell_addr))

    # For each formula group, find contiguous rectangles
    for _sig, group_cells in sorted(formula_groups.items()):
        if len(group_cells) == 1:
            # Single cell: always valid 1x1 binding
            blocks.append([group_cells[0][2]])  # Extract address
        else:
            # Multiple cells: find rectangular regions
            rectangles = _find_contiguous_rectangles(group_cells)
            for rect_cells in rectangles:
                block_addrs = [addr for _, _, addr in rect_cells]
                blocks.append(block_addrs)

    # Sort blocks by top-left position for determinism
    def block_sort_key(block: list[str]) -> tuple[int, int]:
        # Parse first cell address to get position
        if not block:
            return (999999, 999999)
        parsed = parse_cell_address(block[0])
        row_val = parsed.get("row", 999999)
        col_val = parsed.get("col", 999999)
        # Type narrowing for mypy
        assert isinstance(row_val, int) and isinstance(col_val, int)
        return (row_val, col_val)

    blocks.sort(key=block_sort_key)

    return blocks


def group_cells_into_bindings(
    cells: list[str],
    cell_formulas: dict[str, str],
    workbook_guid: str = "",
    formula_is_r1c1: bool = False,
    cell_signatures: dict[str, str] | None = None,
) -> list[Binding]:
    """
    Group visited cells into contiguous rectangular bindings.

    Per ADR-002: Each binding is a contiguous rectangular range on one sheet.
    Groups cells greedily by sheet, finding maximal rectangular regions.

    Args:
        cells: List of cell addresses (e.g., ["Sheet1!A1", "Sheet1!A2"])
        cell_formulas: Mapping from cell address to formula_a1 (empty string if no formula)
        workbook_guid: Workbook GUID for binding_id computation (default: "")

    Returns:
        List of Binding objects with computed identities

    Algorithm:
        1. Group cells by sheet
        2. For each sheet, find contiguous rectangular regions
        3. Compute identities (cells_structure_hash, binding_id) for each binding

    Determinism:
        - Processes sheets in sorted order
        - Processes cells in sorted order (row, col)
        - Hash computation uses sorted tuples

    Example:
        >>> cells = ["Sheet1!A1", "Sheet1!A2", "Sheet1!B1"]
        >>> formulas = {"Sheet1!A1": "=10", "Sheet1!A2": "=20", "Sheet1!B1": "=30"}
        >>> bindings = group_cells_into_bindings(cells, formulas)
        >>> len(bindings)  # May be 1 or multiple depending on contiguity
    """
    if not cells:
        return []

    # Group cells by sheet
    sheet_cells: dict[str, list[tuple[int, int, str]]] = {}
    for cell in cells:
        sheet, row, col = _parse_a1_address(cell)
        if sheet not in sheet_cells:
            sheet_cells[sheet] = []
        sheet_cells[sheet].append((row, col, cell))

    # Process each sheet
    bindings: list[Binding] = []

    for sheet in sorted(sheet_cells.keys()):
        cells_list = sheet_cells[sheet]

        # Sort cells by row, then col
        cells_list.sort(key=lambda x: (x[0], x[1]))

        # Detect contiguous blocks (multi-cell bindings)
        blocks = _detect_contiguous_blocks(
            cells_list,
            cell_formulas,
            formula_is_r1c1=formula_is_r1c1,
            cell_signatures=cell_signatures,
        )

        # Create binding for each block
        for block_cells in blocks:
            binding = _create_binding_for_cells(
                sheet=sheet,
                cells=block_cells,
                cell_formulas=cell_formulas,
                workbook_guid=workbook_guid,
            )
            bindings.append(binding)

    return bindings


def _create_binding_for_cells(
    sheet: str,
    cells: list[str],
    cell_formulas: dict[str, str],
    workbook_guid: str,
) -> Binding:
    """
    Create a Binding object for a group of cells.

    Args:
        sheet: Sheet name
        cells: List of cell addresses in this binding
        cell_formulas: Mapping from cell address to formula
        workbook_guid: Workbook GUID for binding_id computation

    Returns:
        Binding object with computed identities
    """
    # Parse all cell addresses
    parsed_cells = [_parse_a1_address(cell) for cell in cells]
    rows = [row for _, row, _ in parsed_cells]
    cols = [col for _, _, col in parsed_cells]

    # Compute bounding box
    min_row, max_row = min(rows), max(rows)
    min_col, max_col = min(cols), max(cols)

    # Compute shape
    shape_rows = max_row - min_row + 1
    shape_cols = max_col - min_col + 1

    # Top-left cell
    top_left_a1 = f"{_col_to_letter(min_col)}{min_row}"

    # Full address range
    address_a1 = _format_a1_range(sheet, min_row, min_col, max_row, max_col)

    # Compute cells_structure_hash
    cells_structure_hash = compute_cells_structure_hash(cells, cell_formulas)

    # Compute binding_id
    binding_id = compute_binding_id(
        workbook_guid=workbook_guid,
        sheet=sheet,
        top_left_a1=top_left_a1,
        shape_rows=shape_rows,
        shape_cols=shape_cols,
        cells_structure_hash=cells_structure_hash,
    )

    # Debug label
    debug_label = f"{sheet}::{top_left_a1}::{_col_to_letter(max_col)}{max_row}"

    return Binding(
        binding_id=binding_id,
        debug_label=debug_label,
        sheet=sheet,
        address_a1=address_a1,
        top_left_a1=top_left_a1,
        shape_rows=shape_rows,
        shape_cols=shape_cols,
        binding_type="formula",
        cells_structure_hash=cells_structure_hash,
        cells=sorted(cells),
    )


def compute_cells_structure_hash(
    cells: list[str],
    cell_formulas: dict[str, str],
) -> str:
    """
    Compute cells_structure_hash for a binding.

    Per IR Spec §3.1:
    - Hash of the set of (cell_address_a1, formula_a1) tuples
    - Tuples are sorted lexicographically for order independence
    - Uses SHA-256, lowercase hex, 64 characters

    Args:
        cells: List of cell addresses in binding
        cell_formulas: Mapping from cell address to formula_a1

    Returns:
        SHA-256 hash (hex64) of sorted tuples

    Example:
        >>> cells = ["Sheet1!A1", "Sheet1!A2"]
        >>> formulas = {"Sheet1!A1": "=10", "Sheet1!A2": "=A1*2"}
        >>> hash_val = compute_cells_structure_hash(cells, formulas)
        >>> len(hash_val)
        64
        >>> hash_val.islower()
        True
    """
    # Create tuples of (address, formula)
    tuples = []
    for cell in cells:
        formula = cell_formulas.get(cell, "")
        tuples.append((cell, formula))

    # Sort tuples lexicographically (address then formula)
    tuples.sort()

    # Serialize as JSON
    json_str = json.dumps(tuples, separators=(",", ":"), sort_keys=True)

    # Compute SHA-256 hash
    hash_obj = hashlib.sha256(json_str.encode("utf-8"))
    return hash_obj.hexdigest()


def compute_binding_id(
    workbook_guid: str,
    sheet: str,
    top_left_a1: str,
    shape_rows: int,
    shape_cols: int,
    cells_structure_hash: str,
) -> str:
    """
    Compute binding_id per IR Spec §3.1.

    Formula:
        binding_id = SHA256(workbook_guid + "::" + sheet + "::" +
                           top_left_a1 + "::" + rows + "x" + cols + "::" +
                           cells_structure_hash)

    Args:
        workbook_guid: Workbook GUID (hex64)
        sheet: Sheet name
        top_left_a1: Top-left cell address (e.g., "B5")
        shape_rows: Number of rows
        shape_cols: Number of columns
        cells_structure_hash: Hash of cell structure (hex64)

    Returns:
        SHA-256 hash (hex64) of binding identity components

    Example:
        >>> guid = "a" * 64
        >>> hash_val = "b" * 64
        >>> binding_id = compute_binding_id(guid, "Sheet1", "B5", 6, 3, hash_val)
        >>> len(binding_id)
        64
        >>> binding_id.islower()
        True
    """
    # Construct identity string
    identity_str = (
        f"{workbook_guid}::{sheet}::{top_left_a1}::"
        f"{shape_rows}x{shape_cols}::{cells_structure_hash}"
    )

    # Compute SHA-256 hash
    hash_obj = hashlib.sha256(identity_str.encode("utf-8"))
    return hash_obj.hexdigest()


def collapse_cell_edges_to_binding_edges(
    cell_edges: list[tuple[str, str]],
    cell_to_binding: dict[str, list[str]],
) -> list[BindingEdge]:
    """
    Collapse cell-level edges to binding-level edges.

    Per IR Spec §5:
    - If any cell in Binding A depends on any cell in Binding B,
      create binding-level edge A → B
    - Deduplicates edges (same from_binding + to_binding)
    - Supports overlapping bindings (one-to-many mapping)

    Args:
        cell_edges: List of (from_cell, to_cell) tuples
        cell_to_binding: Mapping from cell address to list of binding_ids

    Returns:
        List of unique BindingEdge objects

    Example:
        >>> edges = [("Sheet1!B1", "Sheet1!A1")]
        >>> mapping = {
        ...     "Sheet1!A1": ["binding1", "binding2"],
        ...     "Sheet1!B1": ["binding3"]
        ... }
        >>> binding_edges = collapse_cell_edges_to_binding_edges(edges, mapping)
        >>> len(binding_edges)
        2
    """
    # Collect unique binding-level edges
    binding_edge_set: set[tuple[str, str]] = set()

    for from_cell, to_cell in cell_edges:
        from_bindings = cell_to_binding.get(from_cell, [])
        to_bindings = cell_to_binding.get(to_cell, [])

        # Iterate over all combinations of source and target bindings
        for from_binding in from_bindings:
            for to_binding in to_bindings:
                # Skip self-loops (binding depends on itself)
                if from_binding == to_binding:
                    continue

                binding_edge_set.add((from_binding, to_binding))

    # Convert to BindingEdge objects and sort for determinism
    binding_edges = [
        BindingEdge(from_binding_id=from_id, to_binding_id=to_id)
        for from_id, to_id in binding_edge_set
    ]

    # Sort by (from_binding_id, to_binding_id) for deterministic output
    binding_edges.sort(key=lambda e: (e.from_binding_id, e.to_binding_id))

    return binding_edges


def create_structure_hash_entries(
    bindings: list[Binding],
) -> list[dict[str, str]]:
    """
    Create structure_hashes table entries for bindings.

    Per Story 5 requirements:
    - hash_type = "bindings"
    - hash_key = binding_id
    - hash_value = cells_structure_hash

    Args:
        bindings: List of Binding objects

    Returns:
        List of dictionaries with hash_type, hash_key, hash_value

    Example:
        >>> binding = Binding(
        ...     binding_id="abc123",
        ...     cells_structure_hash="def456",
        ...     # ... other fields
        ... )
        >>> entries = create_structure_hash_entries([binding])
        >>> entries[0]["hash_type"]
        'bindings'
        >>> entries[0]["hash_key"]
        'abc123'
    """
    entries = []
    for binding in bindings:
        entries.append(
            {
                "hash_type": "bindings",
                "hash_key": binding.binding_id,
                "hash_value": binding.cells_structure_hash,
            }
        )

    return entries
