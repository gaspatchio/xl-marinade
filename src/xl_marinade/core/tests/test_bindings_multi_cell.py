# ABOUTME: Tests for multi-cell binding detection (Story 8b)
# ABOUTME: Validates R1C1 formula matching, rectangle validation, and connected-component grouping

from xl_marinade.core.bindings import (
    _cells_form_rectangle,
    _detect_contiguous_blocks,
    _find_contiguous_rectangles,
    _get_r1c1_signature,
    group_cells_into_bindings,
)


class TestR1C1SignatureExtraction:
    """Test R1C1 signature extraction for formula pattern matching."""

    def test_r1c1_signature_basic(self) -> None:
        """Test R1C1 signature extraction using existing infrastructure."""
        # Test: B5 with formula "=A5" should give "=RC[-1]"
        sig = _get_r1c1_signature("Sheet1!B5", "=A5")
        assert sig == "=RC[-1]"

        # Test: C5 with formula "=B5" should also give "=RC[-1]" (same signature!)
        sig2 = _get_r1c1_signature("Sheet1!C5", "=B5")
        assert sig2 == "=RC[-1]"

        # Test: Empty formula
        sig3 = _get_r1c1_signature("Sheet1!A1", "")
        assert sig3 == ""

    def test_r1c1_signature_relative_refs(self) -> None:
        """Test R1C1 signatures for relative references."""
        # Column-relative formula
        sig1 = _get_r1c1_signature("Sheet1!C10", "=B10")
        assert sig1 == "=RC[-1]"

        # Row-relative formula
        sig2 = _get_r1c1_signature("Sheet1!B10", "=B9")
        assert sig2 == "=R[-1]C"

        # Both relative
        sig3 = _get_r1c1_signature("Sheet1!C10", "=B9")
        assert sig3 == "=R[-1]C[-1]"


class TestRectangleValidation:
    """Test rectangle validation for contiguous cell groups."""

    def test_rectangle_validation_valid(self) -> None:
        """Test rectangle validation for valid 2x3 rectangle."""
        cells = [
            (2, 2, "B2"),
            (2, 3, "C2"),
            (2, 4, "D2"),
            (3, 2, "B3"),
            (3, 3, "C3"),
            (3, 4, "D3"),
        ]
        assert _cells_form_rectangle(cells) is True

    def test_rectangle_validation_single_cell(self) -> None:
        """Test single cell is a valid 1x1 rectangle."""
        cells = [(5, 5, "E5")]
        assert _cells_form_rectangle(cells) is True

    def test_rectangle_validation_gap(self) -> None:
        """Test rectangle validation rejects L-shaped region."""
        cells = [(2, 2, "B2"), (2, 3, "C2"), (3, 2, "B3")]  # Missing (3,3)
        assert _cells_form_rectangle(cells) is False

    def test_rectangle_validation_horizontal_line(self) -> None:
        """Test horizontal line (1xN) is valid rectangle."""
        cells = [(5, 2, "B5"), (5, 3, "C5"), (5, 4, "D5")]
        assert _cells_form_rectangle(cells) is True

    def test_rectangle_validation_vertical_line(self) -> None:
        """Test vertical line (Mx1) is valid rectangle."""
        cells = [(2, 5, "E2"), (3, 5, "E3"), (4, 5, "E4")]
        assert _cells_form_rectangle(cells) is True


class TestConnectedComponentGrouping:
    """Test flood-fill connected component detection."""

    def test_find_single_rectangle(self) -> None:
        """Test finding single contiguous rectangle."""
        cells = [
            (2, 2, "Sheet1!B2"),
            (2, 3, "Sheet1!C2"),
            (3, 2, "Sheet1!B3"),
            (3, 3, "Sheet1!C3"),
        ]
        rectangles = _find_contiguous_rectangles(cells)

        assert len(rectangles) == 1
        assert len(rectangles[0]) == 4

    def test_find_multiple_rectangles_with_gap(self) -> None:
        """Test finding TWO separate rectangles with gap."""
        cells = [
            # First rectangle: B2:C2
            (2, 2, "Sheet1!B2"),
            (2, 3, "Sheet1!C2"),
            # Gap at row 3
            # Second rectangle: B4:C4
            (4, 2, "Sheet1!B4"),
            (4, 3, "Sheet1!C4"),
        ]
        rectangles = _find_contiguous_rectangles(cells)

        # Should find 2 separate rectangles
        assert len(rectangles) == 2
        assert all(len(r) == 2 for r in rectangles)

    def test_l_shaped_decomposes_to_singles(self) -> None:
        """Test L-shaped region decomposes to individual 1x1 cells."""
        cells = [
            (2, 2, "Sheet1!B2"),
            (2, 3, "Sheet1!C2"),  # Top of L
            (3, 2, "Sheet1!B3"),  # Bottom of L (missing C3)
        ]
        rectangles = _find_contiguous_rectangles(cells)

        # L-shape is not rectangular, should decompose to 3x 1x1
        assert len(rectangles) == 3
        assert all(len(r) == 1 for r in rectangles)


class TestContiguousBlockDetection:
    """Test full block detection algorithm with R1C1 matching."""

    def test_block_detection_single_rectangle(self) -> None:
        """Test detecting single 1x3 horizontal block."""
        cells = [(2, 2, "Sheet1!B2"), (2, 3, "Sheet1!C2"), (2, 4, "Sheet1!D2")]
        formulas = {"Sheet1!B2": "=A2", "Sheet1!C2": "=B2", "Sheet1!D2": "=C2"}
        blocks = _detect_contiguous_blocks(cells, formulas)

        # All have same R1C1 signature (=RC[-1]), should group
        assert len(blocks) == 1
        assert len(blocks[0]) == 3

    def test_block_detection_multiple_rectangles(self) -> None:
        """Test detecting TWO separate 1x3 blocks with same formula."""
        cells = [
            # First block: B2:D2
            (2, 2, "Sheet1!B2"),
            (2, 3, "Sheet1!C2"),
            (2, 4, "Sheet1!D2"),
            # Second block: B5:D5 (gap at rows 3-4)
            (5, 2, "Sheet1!B5"),
            (5, 3, "Sheet1!C5"),
            (5, 4, "Sheet1!D5"),
        ]
        formulas = {
            "Sheet1!B2": "=A2",
            "Sheet1!C2": "=B2",
            "Sheet1!D2": "=C2",
            "Sheet1!B5": "=A5",
            "Sheet1!C5": "=B5",
            "Sheet1!D5": "=C5",
        }
        blocks = _detect_contiguous_blocks(cells, formulas)

        # All 6 cells have same R1C1, but form TWO separate rectangles
        assert len(blocks) == 2
        assert all(len(b) == 3 for b in blocks)

    def test_block_detection_l_shaped_decomposes(self) -> None:
        """Test L-shaped region decomposes to individual cells."""
        cells = [
            (2, 2, "Sheet1!B2"),
            (2, 3, "Sheet1!C2"),  # Top of L
            (3, 2, "Sheet1!B3"),  # Bottom of L (missing C3)
        ]
        formulas = {"Sheet1!B2": "=A2", "Sheet1!C2": "=B2", "Sheet1!B3": "=A3"}
        blocks = _detect_contiguous_blocks(cells, formulas)

        # L-shaped is not rectangular, should decompose to 3x 1x1
        assert len(blocks) == 3
        assert all(len(b) == 1 for b in blocks)

    def test_block_detection_different_formulas_dont_group(self) -> None:
        """Test cells with different R1C1 signatures don't group."""
        cells = [(2, 2, "Sheet1!B2"), (2, 3, "Sheet1!C2"), (2, 4, "Sheet1!D2")]
        formulas = {
            "Sheet1!B2": "=A2",  # R1C1: =RC[-1]
            "Sheet1!C2": "=A3",  # R1C1: =R[1]C[-2] (different!)
            "Sheet1!D2": "=C2",  # R1C1: =RC[-1] (same as B2)
        }
        blocks = _detect_contiguous_blocks(cells, formulas)

        # B2 and D2 have same R1C1 but aren't contiguous, so 3 separate bindings
        assert len(blocks) == 3

    def test_block_detection_value_only_cells(self) -> None:
        """Test value-only cells each get unique signature (don't group)."""
        cells = [(2, 2, "Sheet1!B2"), (2, 3, "Sheet1!C2"), (2, 4, "Sheet1!D2")]
        formulas = {
            "Sheet1!B2": "",  # Value-only
            "Sheet1!C2": "",  # Value-only
            "Sheet1!D2": "",  # Value-only
        }
        blocks = _detect_contiguous_blocks(cells, formulas)

        # Value-only cells don't group (each has unique signature)
        assert len(blocks) == 3
        assert all(len(b) == 1 for b in blocks)


class TestGroupCellsIntoBindings:
    """Integration tests for full binding detection."""

    def test_group_horizontal_range(self) -> None:
        """Test grouping 1xN horizontal range into single binding."""
        cells = ["Sheet1!B2", "Sheet1!C2", "Sheet1!D2"]
        formulas = {"Sheet1!B2": "=A2", "Sheet1!C2": "=B2", "Sheet1!D2": "=C2"}
        bindings = group_cells_into_bindings(cells, formulas, "test-wb")

        assert len(bindings) == 1
        assert bindings[0].shape_rows == 1
        assert bindings[0].shape_cols == 3
        assert bindings[0].address_a1 == "Sheet1!B2:D2"

    def test_group_vertical_range(self) -> None:
        """Test grouping Mx1 vertical range into single binding."""
        cells = ["Sheet1!B2", "Sheet1!B3", "Sheet1!B4"]
        formulas = {"Sheet1!B2": "=A2", "Sheet1!B3": "=A3", "Sheet1!B4": "=A4"}
        bindings = group_cells_into_bindings(cells, formulas, "test-wb")

        assert len(bindings) == 1
        assert bindings[0].shape_rows == 3
        assert bindings[0].shape_cols == 1
        assert bindings[0].address_a1 == "Sheet1!B2:B4"

    def test_group_rectangular_block(self) -> None:
        """Test grouping MxN rectangular block."""
        cells = ["Sheet1!B2", "Sheet1!C2", "Sheet1!D2", "Sheet1!B3", "Sheet1!C3", "Sheet1!D3"]
        formulas = {
            "Sheet1!B2": "=A2",
            "Sheet1!C2": "=B2",
            "Sheet1!D2": "=C2",
            "Sheet1!B3": "=A3",
            "Sheet1!C3": "=B3",
            "Sheet1!D3": "=C3",
        }
        bindings = group_cells_into_bindings(cells, formulas, "test-wb")

        assert len(bindings) == 1
        assert bindings[0].shape_rows == 2
        assert bindings[0].shape_cols == 3

    def test_mixed_grouping(self) -> None:
        """Test some cells group, others stay 1x1."""
        cells = [
            "Sheet1!B2",
            "Sheet1!C2",  # Can group (same R1C1)
            "Sheet1!B4",  # Separate (gap at row 3)
        ]
        formulas = {"Sheet1!B2": "=A2", "Sheet1!C2": "=B2", "Sheet1!B4": "=A4"}
        bindings = group_cells_into_bindings(cells, formulas, "test-wb")

        # Should be 2 bindings: B2:C2 (1x2) and B4 (1x1)
        assert len(bindings) == 2

    def test_backward_compatibility_single_cells(self) -> None:
        """Test single cells still create 1x1 bindings (backward compatibility)."""
        cells = ["Sheet1!A1"]
        formulas = {"Sheet1!A1": "=10"}
        bindings = group_cells_into_bindings(cells, formulas, "test-wb")

        assert len(bindings) == 1
        assert bindings[0].shape_rows == 1
        assert bindings[0].shape_cols == 1
        assert bindings[0].address_a1 == "Sheet1!A1"

    def test_multiple_separate_blocks_same_formula(self) -> None:
        """Test multiple disconnected regions with same R1C1 create separate bindings."""
        cells = [
            "Sheet1!B2",
            "Sheet1!C2",
            "Sheet1!D2",  # Block 1
            "Sheet1!B5",
            "Sheet1!C5",
            "Sheet1!D5",  # Block 2 (gap at rows 3-4)
        ]
        formulas = {
            "Sheet1!B2": "=A2",
            "Sheet1!C2": "=B2",
            "Sheet1!D2": "=C2",
            "Sheet1!B5": "=A5",
            "Sheet1!C5": "=B5",
            "Sheet1!D5": "=C5",
        }
        bindings = group_cells_into_bindings(cells, formulas, "test")

        # Should create TWO 1x3 bindings (separate connected components)
        assert len(bindings) == 2
        assert all(b.shape_rows == 1 and b.shape_cols == 3 for b in bindings)

        # Verify addresses
        addresses = sorted([b.address_a1 for b in bindings])
        assert addresses == ["Sheet1!B2:D2", "Sheet1!B5:D5"]


class TestDeterminism:
    """Test determinism requirements for multi-cell bindings."""

    def test_multi_cell_binding_determinism(self) -> None:
        """Test multi-cell bindings produce identical IDs across runs."""
        cells = ["Sheet1!B2", "Sheet1!C2", "Sheet1!D2"]
        formulas = {"Sheet1!B2": "=A2", "Sheet1!C2": "=B2", "Sheet1!D2": "=C2"}

        # Run twice
        bindings1 = group_cells_into_bindings(cells, formulas, "test-wb")
        bindings2 = group_cells_into_bindings(cells, formulas, "test-wb")

        assert len(bindings1) == len(bindings2)
        assert bindings1[0].binding_id == bindings2[0].binding_id
        assert bindings1[0].cells_structure_hash == bindings2[0].cells_structure_hash

    def test_cell_order_independence(self) -> None:
        """Test different cell processing order produces same bindings."""
        cells_forward = ["Sheet1!B2", "Sheet1!C2", "Sheet1!D2"]
        cells_reverse = ["Sheet1!D2", "Sheet1!C2", "Sheet1!B2"]
        formulas = {"Sheet1!B2": "=A2", "Sheet1!C2": "=B2", "Sheet1!D2": "=C2"}

        bindings1 = group_cells_into_bindings(cells_forward, formulas, "test-wb")
        bindings2 = group_cells_into_bindings(cells_reverse, formulas, "test-wb")

        # Should produce identical bindings regardless of input order
        assert len(bindings1) == len(bindings2)
        assert bindings1[0].binding_id == bindings2[0].binding_id
        assert bindings1[0].address_a1 == bindings2[0].address_a1


class TestEdgeCases:
    """Test edge cases and error conditions."""

    def test_empty_cells_list(self) -> None:
        """Test empty cells list returns empty bindings."""
        bindings = group_cells_into_bindings([], {}, "test-wb")
        assert bindings == []

    def test_cells_with_no_formulas(self) -> None:
        """Test cells with no formulas create 1x1 bindings each."""
        cells = ["Sheet1!A1", "Sheet1!A2"]
        formulas = {"Sheet1!A1": "", "Sheet1!A2": ""}
        bindings = group_cells_into_bindings(cells, formulas, "test-wb")

        # Each value-only cell is separate binding
        assert len(bindings) == 2
        assert all(b.shape_rows == 1 and b.shape_cols == 1 for b in bindings)

    def test_diagonal_cells_dont_group(self) -> None:
        """Test diagonal adjacency doesn't count as contiguous."""
        cells = [
            "Sheet1!B2",
            "Sheet1!C3",  # Diagonal neighbors, not contiguous
        ]
        formulas = {
            "Sheet1!B2": "=A1",
            "Sheet1!C3": "=B2",  # Same R1C1 if it were =B2 from C3
        }
        bindings = group_cells_into_bindings(cells, formulas, "test-wb")

        # Diagonal cells should NOT group (only horizontal/vertical adjacency)
        assert len(bindings) == 2
