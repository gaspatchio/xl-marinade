# ABOUTME: Stage 4b — Cell matching between canonicalized IR models.
# ABOUTME: Global coordinate matches + local binding-relative matches.

from __future__ import annotations

from xl_marinade.core.ir_diff.model import (
    BindingMatch,
    CellKey,
    CellMatch,
    IRModel,
)


def match_cells(
    a: IRModel,
    b: IRModel,
    binding_match: BindingMatch,
) -> CellMatch:
    """Match cells between canonicalized A and B models.

    Two passes:
    1. Global coordinate match: cells at the same (sheet, row, col) after canonicalization.
    2. Local binding-relative match: for matched binding pairs, pair unmatched cells
       by relative offset within the binding.

    Each cell appears in at most one pair (uniqueness invariant).

    Args:
        a: Canonicalized Version A model.
        b: Canonicalized Version B model.
        binding_match: Result of binding matching.

    Returns:
        CellMatch with matched pairs and unmatched cells.
    """
    result = CellMatch()
    matched_a: set[CellKey] = set()
    matched_b: set[CellKey] = set()

    b_cells_by_key: dict[CellKey, bool] = dict.fromkeys(b.cells, True)

    # --- Pass 1: Global coordinate match ---
    for ck_a in sorted(a.cells.keys()):
        if ck_a in b_cells_by_key:
            result.matched[ck_a] = ck_a
            matched_a.add(ck_a)
            matched_b.add(ck_a)

    # --- Pass 2: Local binding-relative match ---
    # Invert cell_to_binding ONCE per model (O(total memberships)) instead of
    # scanning the full cell map once per matched binding pair, which is
    # O(cells x bindings) and dominates the whole diff on large workbooks.
    a_members_by_binding = _all_binding_member_cells(a)
    b_members_by_binding = _all_binding_member_cells(b)

    for ak, bk in sorted(binding_match.matched.items()):
        a_desc = a.bindings.get(ak)
        b_desc = b.bindings.get(bk)
        if not a_desc or not b_desc:
            continue

        # Find member cells by offset
        a_members = a_members_by_binding.get(ak, {})
        b_members = b_members_by_binding.get(bk, {})

        # Match by shared relative offset
        for offset in sorted(set(a_members.keys()) & set(b_members.keys())):
            ck_a = a_members[offset]
            ck_b = b_members[offset]
            if ck_a not in matched_a and ck_b not in matched_b:
                result.matched[ck_a] = ck_b
                matched_a.add(ck_a)
                matched_b.add(ck_b)

    # --- Unmatched cells ---
    result.removed = sorted(ck for ck in a.cells if ck not in matched_a)
    result.added = sorted(ck for ck in b.cells if ck not in matched_b)

    return result


def _all_binding_member_cells(
    model: IRModel,
) -> dict:
    """Invert cell_to_binding: binding key -> {(row_offset, col_offset): CellKey}.

    Offsets are relative to each binding's top-left corner.
    """
    members: dict = {}
    for ck, bkeys in model.cell_to_binding.items():
        for bk in bkeys:
            members.setdefault(bk, {})[(ck.row - bk.top_left_row, ck.col - bk.top_left_col)] = ck
    return members
