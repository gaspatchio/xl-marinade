# ABOUTME: Refinement logic to merge singleton anchor constants and multi-cell bindings.
# ABOUTME: Implements "Anchor-Chain Grouping" and "Constant-Formula Merging".


from xl_marinade.core.bindings import (
    Binding,
    _col_to_letter,
    _format_a1_range,
    _get_r1c1_signature,
    _parse_a1_address,
    compute_binding_id,
    compute_cells_structure_hash,
)
from xl_marinade.core.grouping.geometry import parse_a1_address
from xl_marinade.core.ref_converter import quote_sheet_name
from xl_marinade.core.reverse_index import ReverseIndex


class RefinementEngine:
    """
    Refines binding groups by identifying and merging logical patterns.

    Implements two refinement passes:
    1. "Anchor-Chain Grouping": Merging singleton constants into formula chains
    2. "Constant-Formula Merging": Merging multi-cell constant ranges with adjacent formula bindings
    """

    def __init__(self, workbook_guid: str, cell_formulas: dict[str, str]):
        self.workbook_guid = workbook_guid
        self.cell_formulas = cell_formulas
        # Cache R1C1 signatures for repeated comparisons during refinement.
        self._r1c1_cache: dict[str, str] = {}

    def _get_r1c1_cached(self, cell_addr: str) -> str:
        formula = self.cell_formulas.get(cell_addr)
        if not formula:
            return ""
        cached = self._r1c1_cache.get(cell_addr)
        if cached is not None:
            return cached
        signature = _get_r1c1_signature(cell_addr, formula)
        self._r1c1_cache[cell_addr] = signature
        return signature

    def split_heterogeneous_bindings(self, bindings: list[Binding]) -> list[Binding]:
        """
        Recursively split 2D bindings that contain heterogeneous R1C1 formulas.

        Priority: Column-wise splitting first, then row-wise.

        Args:
            bindings: List of initial bindings.

        Returns:
            List of bindings with heterogeneous 2D bindings split into homogeneous parts.
        """
        new_bindings = []

        for binding in bindings:
            # Only check 2D formula bindings (width > 1 AND height > 1)
            if (
                binding.binding_type == "formula"
                and binding.shape_rows > 1
                and binding.shape_cols > 1
            ):
                # Recursive split
                split_parts = self._split_binding_recursive(binding)
                new_bindings.extend(split_parts)
            else:
                # Pass through unchanged
                new_bindings.append(binding)

        return new_bindings

    def _split_binding_recursive(self, binding: Binding) -> list[Binding]:
        """Recursive helper to split bindings."""
        # Check columns first
        col_split = self._try_split_columns(binding)
        if col_split:
            # Recurse on parts
            results = []
            for part in col_split:
                results.extend(self._split_binding_recursive(part))
            return results

        # Check rows second
        row_split = self._try_split_rows(binding)
        if row_split:
            # Recurse on parts
            results = []
            for part in row_split:
                results.extend(self._split_binding_recursive(part))
            return results

        # No split needed
        return [binding]

    def _try_split_columns(self, binding: Binding) -> list[Binding] | None:
        """Try to split binding at first column with different formula."""
        if binding.shape_cols <= 1:
            return None

        # Parse geometry
        try:
            if ":" in binding.address_a1:
                addr_part = binding.address_a1.split("!")[-1]
                box = parse_a1_address(addr_part)
            else:
                # Handle single cell case (though shape_cols > 1 check should prevent this)
                addr_part = binding.address_a1.split("!")[-1]
                box = parse_a1_address(addr_part)
        except Exception:
            return None

        # Get representative cell for first column
        first_col = box.min_col
        first_col_cell = self._get_cell_at(binding.sheet, box.min_row, first_col)

        # Be robust if cell not found in formulas (shouldn't happen for valid binding)
        first_formula_r1c1 = self._get_r1c1_cached(first_col_cell)
        if not first_formula_r1c1:
            return None  # Can't compare

        # Iterate through columns to find split point
        # Start from min_col + 1 (second column)
        for col_idx in range(box.min_col + 1, box.max_col + 1):
            # Check top cell of this column
            cell = self._get_cell_at(binding.sheet, box.min_row, col_idx)
            current_formula_r1c1 = self._get_r1c1_cached(cell)

            if current_formula_r1c1 != first_formula_r1c1:
                # Found split point!
                # Split is BEFORE col_idx
                # Left part: min_col to col_idx - 1
                # Right part: col_idx to max_col

                # NOTE: _create_sub_binding creates valid bindings with correct address_a1
                # which will be picked up by recursion
                left_part = self._create_sub_binding(
                    binding, box.min_row, box.max_row, box.min_col, col_idx - 1
                )
                right_part = self._create_sub_binding(
                    binding, box.min_row, box.max_row, col_idx, box.max_col
                )

                return [left_part, right_part]

        return None

    def _try_split_rows(self, binding: Binding) -> list[Binding] | None:
        """Try to split binding at first row with different formula."""
        if binding.shape_rows <= 1:
            return None

        # Parse geometry
        try:
            if ":" in binding.address_a1:
                addr_part = binding.address_a1.split("!")[-1]
                box = parse_a1_address(addr_part)
            else:
                addr_part = binding.address_a1.split("!")[-1]
                box = parse_a1_address(addr_part)
        except Exception:
            return None

        # Get representative cell for first row
        first_row = box.min_row
        first_row_cell = self._get_cell_at(binding.sheet, first_row, box.min_col)

        first_formula_r1c1 = self._get_r1c1_cached(first_row_cell)
        if not first_formula_r1c1:
            return None

        # Iterate through rows
        for row_idx in range(box.min_row + 1, box.max_row + 1):
            cell = self._get_cell_at(binding.sheet, row_idx, box.min_col)
            current_formula_r1c1 = self._get_r1c1_cached(cell)

            if current_formula_r1c1 != first_formula_r1c1:
                # Found split point!
                top_part = self._create_sub_binding(
                    binding, box.min_row, row_idx - 1, box.min_col, box.max_col
                )
                bottom_part = self._create_sub_binding(
                    binding, row_idx, box.max_row, box.min_col, box.max_col
                )

                return [top_part, bottom_part]

        return None

    def _get_cell_at(self, sheet: str, row: int, col: int) -> str:
        """Construct A1 address for cell."""
        col_letter = _col_to_letter(col)
        quoted_sheet = quote_sheet_name(sheet)
        return f"{quoted_sheet}!{col_letter}{row}"

    def _create_sub_binding(
        self,
        parent: Binding,
        min_row: int,
        max_row: int,
        min_col: int,
        max_col: int,
    ) -> Binding:
        """Create a new binding subset from geometry."""
        # Collect subset of cells
        # Optimization: Filter existing sorted cells instead of regenerating
        # Or simpler: Just generate them? Generating is safer for geometry correctness.

        # Generate all cells in range
        new_cells = []
        quoted_sheet = quote_sheet_name(parent.sheet)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                cell = f"{quoted_sheet}!{_col_to_letter(c)}{r}"
                new_cells.append(cell)

        # Verify they exist in parent cells (sanity check)
        # Note: Parent cells might be sparse? No, bindings are rectangular blocks.
        # group_cells_into_bindings creates rectangular blocks but might have
        # holes if filtered. Refined grouping assumes dense blocks usually.
        # Assume dense for now as splitting logic assumes geometry implies cells.

        shape_rows = max_row - min_row + 1
        shape_cols = max_col - min_col + 1

        top_left_a1 = f"{_col_to_letter(min_col)}{min_row}"
        address_a1 = _format_a1_range(parent.sheet, min_row, min_col, max_row, max_col)

        cells_structure_hash = compute_cells_structure_hash(new_cells, self.cell_formulas)

        binding_id = compute_binding_id(
            workbook_guid=self.workbook_guid,
            sheet=parent.sheet,
            top_left_a1=top_left_a1,
            shape_rows=shape_rows,
            shape_cols=shape_cols,
            cells_structure_hash=cells_structure_hash,
        )

        return Binding(
            binding_id=binding_id,
            debug_label=f"{parent.sheet}::{address_a1}",
            sheet=parent.sheet,
            address_a1=address_a1,
            top_left_a1=top_left_a1,
            shape_rows=shape_rows,
            shape_cols=shape_cols,
            binding_type=parent.binding_type,
            cells_structure_hash=cells_structure_hash,
            cells=new_cells,
            # Inherit candidates? Maybe re-extract later?
            label_candidates_json=parent.label_candidates_json,
        )

    def refine_bindings(
        self,
        bindings: list[Binding],
        reverse_index: ReverseIndex,
        forward_index: dict[str, list[str]] | None = None,
    ) -> list[Binding]:
        """
        Refine bindings through two-step merging process.

        Step 1: Singleton anchor merging (existing logic from Story 8)
        Step 2: Multi-cell constant-formula merging (Story 10)

        Args:
            bindings: List of current bindings
            reverse_index: Dependency index for validating relationships
            forward_index: Cell → precedents mapping (for multi-cell merging)

        Returns:
            New list of bindings with merges applied
        """
        if forward_index is None:
            forward_index = {}
        # Separate candidates
        # We consider two types of anchors:
        # 1. True constants (binding_type="constant")
        # 2. Pseudo-constants (binding_type="formula" but empty formula string)
        #    This happens if they weren't caught by the constant grouping step
        singletons = []
        formulas = []

        for b in bindings:
            if b.binding_type == "constant":
                if b.shape_rows == 1 and b.shape_cols == 1:
                    singletons.append(b)
            elif b.binding_type == "formula":
                formulas.append(b)
                # Check if it's actually a constant singleton
                if b.shape_rows == 1 and b.shape_cols == 1:
                    # Check if formula is empty
                    cell_addr = b.cells[0]
                    if not self.cell_formulas.get(cell_addr):
                        singletons.append(b)

        # Sort for determinism
        singletons.sort(key=lambda b: b.binding_id)

        # Keep track of merged bindings to avoid double-processing
        merged_constant_ids = set()
        merged_formula_ids = set()

        new_bindings = []

        # Index formula bindings by top-left coordinate for fast adjacency lookup
        # Key: (sheet, row, col) of top-left cell
        formula_map = {}
        for fb in formulas:
            sheet, row, col = _parse_a1_address(fb.top_left_a1)
            # Note: _parse_a1_address returns sheet="" if top_left_a1 has no sheet prefix
            # Binding.top_left_a1 usually DOES NOT have sheet prefix (e.g. "A8")
            # So we use fb.sheet for the key
            formula_map[(fb.sheet, row, col)] = fb

        # Try to merge each singleton
        for anchor in singletons:
            if anchor.binding_id in merged_constant_ids:
                continue

            anchor_sheet, anchor_row, anchor_col = _parse_a1_address(anchor.top_left_a1)
            # Use anchor.sheet, not parsed empty sheet

            # Check Vertical Adjacency (Formula is below Anchor)
            # Anchor: (r, c), Formula: (r+1, c)
            candidate_below = formula_map.get((anchor.sheet, anchor_row + 1, anchor_col))

            # Check Horizontal Adjacency (Formula is right of Anchor)
            # Anchor: (r, c), Formula: (r, c+1)
            candidate_right = formula_map.get((anchor.sheet, anchor_row, anchor_col + 1))

            merged = False

            # Try vertical merge first (arbitrary preference, but common in financial models)
            if (
                candidate_below
                and candidate_below.binding_id not in merged_formula_ids
                and self._can_merge(anchor, candidate_below, "vertical", reverse_index)
            ):
                new_binding = self._merge_bindings(anchor, candidate_below)
                new_bindings.append(new_binding)
                merged_constant_ids.add(anchor.binding_id)
                merged_formula_ids.add(candidate_below.binding_id)
                merged = True

            # If not merged vertically, try horizontal
            if (
                not merged
                and candidate_right
                and candidate_right.binding_id not in merged_formula_ids
                and self._can_merge(anchor, candidate_right, "horizontal", reverse_index)
            ):
                new_binding = self._merge_bindings(anchor, candidate_right)
                new_bindings.append(new_binding)
                merged_constant_ids.add(anchor.binding_id)
                merged_formula_ids.add(candidate_right.binding_id)
                merged = True

        # Add remaining bindings from step 1
        for b in bindings:
            if b.binding_id in merged_constant_ids or b.binding_id in merged_formula_ids:
                continue
            new_bindings.append(b)

        # STEP 2: Multi-Cell Constant-Formula Merging
        if forward_index:
            new_bindings = self._merge_multicell_constants(
                new_bindings, forward_index, merged_constant_ids, merged_formula_ids
            )

        return new_bindings

    def _can_merge(
        self, anchor: Binding, chain: Binding, direction: str, reverse_index: ReverseIndex
    ) -> bool:
        """
        Check if anchor can be merged into chain.

        Criteria:
        1. Dimensions align (Vertical: both width 1; Horizontal: both height 1)
        2. Dependency exists (Chain[0] depends on Anchor)
        """
        # 1. Dimension check
        if direction == "vertical" and chain.shape_cols != 1:
            # Must be 1 column wide to be a vector
            return False
        if direction == "horizontal" and chain.shape_rows != 1:
            # Must be 1 row high to be a vector
            return False

        # 2. Dependency check
        # Chain's first cell must depend on Anchor's (only) cell

        # CRITICAL: chain.cells is sorted alphabetically (A1, A10, A2, ...)
        # We MUST determine the true top-left cell using binding metadata
        # Reconstruct the address from top_left_a1 and sheet
        if "!" in chain.top_left_a1:
            chain_start_cell = chain.top_left_a1
        else:
            quoted_sheet = quote_sheet_name(chain.sheet)
            chain_start_cell = f"{quoted_sheet}!{chain.top_left_a1}"

        # Similarly for anchor (though it's a singleton, so less ambiguous)
        anchor_cell = anchor.cells[0]

        # Check if chain_start_cell is a dependent of anchor_cell
        dependents = reverse_index.get_dependents(anchor_cell)

        return chain_start_cell in dependents

    def _merge_bindings(self, anchor: Binding, chain: Binding) -> Binding:
        """
        Create a new binding merging anchor and chain.
        """
        # Combine cells
        new_cells = sorted(anchor.cells + chain.cells)

        # Parse all cell addresses to compute new geometry
        # We can optimize by just using the union of bounding boxes,
        # but parsing ensures correctness
        parsed_cells = [_parse_a1_address(cell) for cell in new_cells]
        rows = [row for _, row, _ in parsed_cells]
        cols = [col for _, _, col in parsed_cells]

        min_row, max_row = min(rows), max(rows)
        min_col, max_col = min(cols), max(cols)

        shape_rows = max_row - min_row + 1
        shape_cols = max_col - min_col + 1

        top_left_a1 = f"{_col_to_letter(min_col)}{min_row}"
        address_a1 = _format_a1_range(anchor.sheet, min_row, min_col, max_row, max_col)

        # Recompute hash
        # Note: cell_formulas should contain formulas for chain cells and empty/values for anchor
        cells_structure_hash = compute_cells_structure_hash(new_cells, self.cell_formulas)

        # Recompute ID
        binding_id = compute_binding_id(
            workbook_guid=self.workbook_guid,
            sheet=anchor.sheet,
            top_left_a1=top_left_a1,
            shape_rows=shape_rows,
            shape_cols=shape_cols,
            cells_structure_hash=cells_structure_hash,
        )

        # Debug label
        debug_label = f"{anchor.sheet}::{address_a1}"

        return Binding(
            binding_id=binding_id,
            debug_label=debug_label,
            sheet=anchor.sheet,
            address_a1=address_a1,
            top_left_a1=top_left_a1,
            shape_rows=shape_rows,
            shape_cols=shape_cols,
            binding_type="formula",  # Inherit formula type (dominance)
            cells_structure_hash=cells_structure_hash,
            cells=new_cells,
        )

    def _merge_multicell_constants(
        self,
        bindings: list[Binding],
        forward_index: dict[str, list[str]],
        merged_constant_ids: set[str],
        merged_formula_ids: set[str],
    ) -> list[Binding]:
        """
        Merge multi-cell constant bindings with adjacent formula bindings.

        Args:
            bindings: List of bindings after singleton anchor merging
            forward_index: Cell → precedents mapping
            merged_constant_ids: Set of already-merged constant binding IDs (updated in place)
            merged_formula_ids: Set of already-merged formula binding IDs (updated in place)

        Returns:
            New list of bindings with multi-cell merges applied
        """
        # Separate constant and formula bindings
        constant_bindings = [b for b in bindings if b.binding_type == "constant"]
        formula_bindings = [b for b in bindings if b.binding_type == "formula"]

        # Sort for determinism
        constant_bindings.sort(key=lambda b: b.binding_id)
        formula_bindings.sort(key=lambda b: b.binding_id)

        new_bindings = []
        step2_merged_constant_ids = set()
        step2_merged_formula_ids = set()

        # Try to merge each constant binding
        for const_b in constant_bindings:
            if (
                const_b.binding_id in merged_constant_ids
                or const_b.binding_id in step2_merged_constant_ids
            ):
                continue

            # Find adjacent formula binding with dependency
            formula_match = self._find_adjacent_formula_binding(
                const_b, formula_bindings, forward_index
            )

            if (
                formula_match
                and formula_match.binding_id not in merged_formula_ids
                and formula_match.binding_id not in step2_merged_formula_ids
            ):
                # Merge them
                merged = self._merge_bindings(const_b, formula_match)
                new_bindings.append(merged)
                step2_merged_constant_ids.add(const_b.binding_id)
                step2_merged_formula_ids.add(formula_match.binding_id)

        # Add remaining bindings
        for b in bindings:
            if (
                b.binding_id in step2_merged_constant_ids
                or b.binding_id in step2_merged_formula_ids
            ):
                continue
            new_bindings.append(b)

        # Update tracking sets
        merged_constant_ids.update(step2_merged_constant_ids)
        merged_formula_ids.update(step2_merged_formula_ids)

        return new_bindings

    def _is_adjacent_vertical(self, binding_a: Binding, binding_b: Binding) -> bool:
        """
        Check if binding_b is directly below binding_a (or vice versa).

        Adjacency means: same column span, consecutive rows, same sheet.

        Args:
            binding_a: First binding
            binding_b: Second binding

        Returns:
            True if vertically adjacent, False otherwise
        """
        # Must be on same sheet
        if not binding_a or not binding_b:
            return False
        if binding_a.sheet != binding_b.sheet:
            return False

        # Parse A1 addresses to bounding boxes (without sheet prefix)
        # binding.address_a1 may be like "Sheet1!A1:A10" or just "A1:A10"
        try:
            addr_a = binding_a.address_a1.split("!")[-1]
            addr_b = binding_b.address_a1.split("!")[-1]
            box_a = parse_a1_address(addr_a)
            box_b = parse_a1_address(addr_b)
        except (ValueError, IndexError):
            # Malformed address, can't determine adjacency
            return False

        # Same column span?
        if box_a.min_col != box_b.min_col or box_a.max_col != box_b.max_col:
            return False

        # Consecutive rows?
        # A above B: box_a.max_row + 1 == box_b.min_row
        # OR B above A: box_b.max_row + 1 == box_a.min_row
        return box_a.max_row + 1 == box_b.min_row or box_b.max_row + 1 == box_a.min_row

    def _is_adjacent_horizontal(self, binding_a: Binding, binding_b: Binding) -> bool:
        """
        Check if binding_b is directly right of binding_a (or vice versa).

        Adjacency means: same row span, consecutive columns, same sheet.

        Args:
            binding_a: First binding
            binding_b: Second binding

        Returns:
            True if horizontally adjacent, False otherwise
        """
        # Must be on same sheet
        if not binding_a or not binding_b:
            return False
        if binding_a.sheet != binding_b.sheet:
            return False

        try:
            addr_a = binding_a.address_a1.split("!")[-1]
            addr_b = binding_b.address_a1.split("!")[-1]
            box_a = parse_a1_address(addr_a)
            box_b = parse_a1_address(addr_b)
        except (ValueError, IndexError):
            return False

        # Same row span?
        if box_a.min_row != box_b.min_row or box_a.max_row != box_b.max_row:
            return False

        # Consecutive columns?
        # A left of B: box_a.max_col + 1 == box_b.min_col
        # OR B left of A: box_b.max_col + 1 == box_a.min_col
        return box_a.max_col + 1 == box_b.min_col or box_b.max_col + 1 == box_a.min_col

    def _check_inter_binding_dependency(
        self,
        constant_binding: Binding,
        formula_binding: Binding,
        forward_index: dict[str, list[str]],
    ) -> bool:
        """
        Check if formula_binding depends on constant_binding.

        Returns True if ANY cell in formula_binding has a precedent that intersects
        constant_binding. Bindings are rectangular by construction (see Binding
        docstring), so we test sheet match + bbox overlap instead of materialising
        every cell in the precedent range — orders of magnitude faster on workbooks
        with large lookup-table references.

        Falls back to the legacy cell-set membership test if bbox parsing fails for
        either side (defensive — should never trigger on well-formed bindings).

        Args:
            constant_binding: The constant binding (potential input table)
            formula_binding: The formula binding (potential consumer)
            forward_index: Cell → precedents mapping (from traversal)
                May contain individual cells (e.g. "Sheet1!A1") OR range strings
                (e.g. "Sheet1!A1:A10")

        Returns:
            True if at least one formula cell depends on at least one constant cell
        """
        if not forward_index:
            return False

        # Pre-compute constant binding bbox once.
        const_sheet = constant_binding.sheet
        const_addr = constant_binding.address_a1.split("!")[-1]
        try:
            const_bbox = parse_a1_address(const_addr)
        except (ValueError, IndexError):
            return self._check_inter_binding_dependency_legacy(
                constant_binding, formula_binding, forward_index
            )

        for formula_cell in formula_binding.cells:
            for prec in forward_index.get(formula_cell, []):
                # Parse precedent sheet + address.
                if "!" in prec:
                    prec_sheet_part, prec_addr = prec.split("!", 1)
                    prec_sheet = prec_sheet_part.strip("'")
                else:
                    prec_sheet = ""
                    prec_addr = prec

                if prec_sheet != const_sheet:
                    continue

                try:
                    prec_bbox = parse_a1_address(prec_addr)
                except (ValueError, IndexError):
                    continue

                if const_bbox.overlaps(prec_bbox):
                    # Both bindings rectangular ⇒ bbox overlap ≡ cell-set intersection.
                    return True

        return False

    def _check_inter_binding_dependency_legacy(
        self,
        constant_binding: Binding,
        formula_binding: Binding,
        forward_index: dict[str, list[str]],
    ) -> bool:
        """
        Cell-set membership implementation retained as a safety fallback and as a
        reference oracle for parity testing. See _check_inter_binding_dependency
        for the production fast path.
        """
        if not forward_index:
            return False

        constant_cells_set = set(constant_binding.cells)

        for formula_cell in formula_binding.cells:
            precedents = forward_index.get(formula_cell, [])
            for prec in precedents:
                if ":" in prec:
                    try:
                        expanded_cells = self._expand_range_to_cells(prec)
                        if any(cell in constant_cells_set for cell in expanded_cells):
                            return True
                    except Exception:
                        continue
                else:
                    if prec in constant_cells_set:
                        return True

        return False

    def _expand_range_to_cells(self, range_ref: str) -> list[str]:
        """
        Expand a range reference to individual cell addresses.

        Args:
            range_ref: Range in A1 notation (e.g., "Sheet1!A1:A10" or "A1:B5")

        Returns:
            List of individual cell addresses (e.g., ["Sheet1!A1", "Sheet1!A2", ...])

        Raises:
            ValueError: If range format is invalid
        """
        # Parse sheet name if present
        if "!" in range_ref:
            sheet_part, addr_part = range_ref.split("!", 1)
            # Remove quotes if present
            sheet_name = sheet_part.strip("'")
        else:
            sheet_name = None
            addr_part = range_ref

        # Parse range using BoundingBox
        box = parse_a1_address(addr_part)

        # Generate all cells in range
        cells = []
        for row in range(box.min_row, box.max_row + 1):
            for col in range(box.min_col, box.max_col + 1):
                col_letter = _col_to_letter(col)
                cell_addr = f"{col_letter}{row}"

                # Add sheet prefix if we have one
                if sheet_name:
                    quoted_sheet = quote_sheet_name(sheet_name)
                    cell_addr = f"{quoted_sheet}!{cell_addr}"

                cells.append(cell_addr)

        return cells

    def _find_adjacent_formula_binding(
        self,
        constant_binding: Binding,
        formula_bindings: list[Binding],
        forward_index: dict[str, list[str]],
    ) -> Binding | None:
        """
        Find an adjacent formula binding that depends on the constant binding.

        Applies deterministic tie-breaking: vertical before horizontal,
        lexicographically smallest binding_id wins ties.

        Args:
            constant_binding: The constant binding to find a match for
            formula_bindings: List of formula bindings (already sorted by binding_id)
            forward_index: Cell dependency index

        Returns:
            Adjacent dependent formula binding, or None if not found
        """
        # Collect vertical candidates (above and below)
        vertical_candidates = [
            fb
            for fb in formula_bindings
            if self._is_adjacent_vertical(constant_binding, fb)
            and self._check_inter_binding_dependency(constant_binding, fb, forward_index)
        ]

        # Vertical takes precedence
        if vertical_candidates:
            # Sort by binding_id for deterministic tie-breaking
            vertical_candidates.sort(key=lambda b: b.binding_id)
            return vertical_candidates[0]

        # If no vertical, try horizontal
        horizontal_candidates = [
            fb
            for fb in formula_bindings
            if self._is_adjacent_horizontal(constant_binding, fb)
            and self._check_inter_binding_dependency(constant_binding, fb, forward_index)
        ]

        if horizontal_candidates:
            horizontal_candidates.sort(key=lambda b: b.binding_id)
            return horizontal_candidates[0]

        return None
