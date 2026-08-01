# ABOUTME: Stage 2b — Row/column axis edit detection per matched sheet.
# ABOUTME: Produces rho (row map) and kappa (col map) using LCS on structural signatures.

from __future__ import annotations

import hashlib
import json

from xl_marinade.core.ir_diff.lcs import alignment_to_maps, lcs_alignment
from xl_marinade.core.ir_diff.model import (
    AxisMap,
    IRModel,
    SheetMatch,
)


def detect_axis_edits(
    a: IRModel,
    b: IRModel,
    sheet_match: SheetMatch,
) -> dict[str, AxisMap]:
    """Detect row/column insertions/deletions for each matched sheet pair.

    For each matched sheet, computes row and column signatures and aligns them
    using lexicographically minimal LCS. Unmatched runs become ROWS_INSERTED,
    ROWS_DELETED, COLS_INSERTED, COLS_DELETED.

    Args:
        a: Version A IR model.
        b: Version B IR model.
        sheet_match: Result of sheet matching.

    Returns:
        Dict mapping A sheet name -> AxisMap.
    """
    result: dict[str, AxisMap] = {}

    # One pass over each model's cells (instead of a full-sheet scan per
    # sheet, plus a full-sheet sort per row/col inside _row_sig/_col_sig,
    # which is O(rows x cells log cells) and dominates large diffs).
    rows_tok_a, cols_tok_a = _gather_axis_tokens(a)
    rows_tok_b, cols_tok_b = _gather_axis_tokens(b)

    for sa in sorted(sheet_match.matched.keys()):
        sb = sheet_match.matched[sa]

        # Occupied rows and cols per sheet
        rows_a = sorted(rows_tok_a.get(sa, {}))
        rows_b = sorted(rows_tok_b.get(sb, {}))
        cols_a = sorted(cols_tok_a.get(sa, {}))
        cols_b = sorted(cols_tok_b.get(sb, {}))

        # Row alignment via LCS on signatures
        row_sigs_a = [_axis_sig(rows_tok_a[sa][r]) for r in rows_a]
        row_sigs_b = [_axis_sig(rows_tok_b[sb][r]) for r in rows_b]
        row_alignment = lcs_alignment(row_sigs_a, row_sigs_b)
        row_map, rows_inserted, rows_deleted = alignment_to_maps(
            row_alignment,
            rows_a,
            rows_b,
        )

        # Fallback: rows that exist at the same index in both A and B but have
        # different signatures should still be mapped as identity (content change,
        # not structural insert/delete). This prevents treating formula/value edits
        # as spurious row insertions/deletions.
        row_map, rows_inserted, rows_deleted = _add_identity_fallback(
            row_map,
            rows_inserted,
            rows_deleted,
            set(rows_a),
            set(rows_b),
        )

        # Column alignment
        col_sigs_a = [_axis_sig(cols_tok_a[sa][c]) for c in cols_a]
        col_sigs_b = [_axis_sig(cols_tok_b[sb][c]) for c in cols_b]
        col_alignment = lcs_alignment(col_sigs_a, col_sigs_b)
        col_map, cols_inserted, cols_deleted = alignment_to_maps(
            col_alignment,
            cols_a,
            cols_b,
        )

        # Same fallback for columns
        col_map, cols_inserted, cols_deleted = _add_identity_fallback(
            col_map,
            cols_inserted,
            cols_deleted,
            set(cols_a),
            set(cols_b),
        )

        result[sa] = AxisMap(
            row_map=row_map,
            col_map=col_map,
            rows_inserted=rows_inserted,
            rows_deleted=rows_deleted,
            cols_inserted=cols_inserted,
            cols_deleted=cols_deleted,
        )

    return result


def _add_identity_fallback(
    key_map: dict[int, int],
    inserted_runs: list[tuple[int, int]],
    deleted_runs: list[tuple[int, int]],
    keys_a: set[int],
    keys_b: set[int],
) -> tuple[dict[int, int], list[tuple[int, int]], list[tuple[int, int]]]:
    """Add identity mappings for keys present in both A and B but unmatched by LCS.

    When a row/col exists at the same index in both databases but has different
    content (formula/value changed), the LCS won't match it. This fallback maps
    such rows/cols as identity (content change, not structural edit), and removes
    them from the inserted/deleted runs.

    Returns updated (key_map, inserted_runs, deleted_runs).
    """
    already_mapped_a = set(key_map.keys())
    already_mapped_b = set(key_map.values())

    # Keys present in both but not yet mapped
    common_unmapped = (keys_a - already_mapped_a) & (keys_b - already_mapped_b)

    if not common_unmapped:
        return key_map, inserted_runs, deleted_runs

    # Add identity mappings
    new_map = dict(key_map)
    for k in sorted(common_unmapped):
        new_map[k] = k

    # Rebuild inserted/deleted runs excluding the now-mapped keys
    mapped_b = set(new_map.values())
    mapped_a = set(new_map.keys())

    new_inserted = _rebuild_runs_excluding(inserted_runs, mapped_b)
    new_deleted = _rebuild_runs_excluding(deleted_runs, mapped_a)

    return new_map, new_inserted, new_deleted


def _rebuild_runs_excluding(
    runs: list[tuple[int, int]],
    exclude: set[int],
) -> list[tuple[int, int]]:
    """Rebuild contiguous runs after excluding some keys."""
    new_runs = []
    for start, count in runs:
        # Expand run into individual keys
        keys = [start + i for i in range(count)]
        # Filter out excluded keys and rebuild runs
        remaining = [k for k in keys if k not in exclude]
        # Group into contiguous runs
        if remaining:
            run_start = remaining[0]
            run_len = 1
            for i in range(1, len(remaining)):
                if remaining[i] == remaining[i - 1] + 1:
                    run_len += 1
                else:
                    new_runs.append((run_start, run_len))
                    run_start = remaining[i]
                    run_len = 1
            new_runs.append((run_start, run_len))
    return new_runs


# ---------------------------------------------------------------------------
# Signature computation
# ---------------------------------------------------------------------------


def _gather_axis_tokens(
    model: IRModel,
) -> tuple[dict[str, dict[int, list]], dict[str, dict[int, list]]]:
    """Group cell tokens per sheet by row and by column in one pass.

    Returns:
        (rows, cols) where rows[sheet][row] = [(col, token), ...] and
        cols[sheet][col] = [(row, token), ...], unsorted.
    """
    rows: dict[str, dict[int, list]] = {}
    cols: dict[str, dict[int, list]] = {}

    for ck, csig in model.cells.items():
        # Token: (formula_or_value_marker, normalized_r1c1, data_type, is_array, is_spilled)
        if csig.formula_r1c1:
            token = (
                "F",
                csig.formula_r1c1,
                csig.data_type or "",
                csig.is_array_formula,
                csig.is_spilled,
            )
        else:
            token = ("V", csig.value_sha256 or "", csig.data_type or "")
        rows.setdefault(ck.sheet, {}).setdefault(ck.row, []).append((ck.col, token))
        cols.setdefault(ck.sheet, {}).setdefault(ck.col, []).append((ck.row, token))

    return rows, cols


def _axis_sig(keyed_tokens: list) -> str:
    """Hash a row's (col, token) list — or a column's (row, token) list —
    ordered by the cross-axis coordinate, matching the original ordering of
    a full (row, col) sort filtered to one row/col."""
    keyed_tokens.sort(key=lambda kt: kt[0])
    tokens = [t for _, t in keyed_tokens]
    raw = json.dumps(tokens, separators=(",", ":"), sort_keys=True)
    return hashlib.sha256(raw.encode()).hexdigest()[:16]
