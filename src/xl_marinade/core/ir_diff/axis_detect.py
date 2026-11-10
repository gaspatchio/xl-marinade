# ABOUTME: Stage 2b — Row/column axis edit detection per matched sheet.
# ABOUTME: Produces rho (row map) and kappa (col map) using LCS on structural signatures.

from __future__ import annotations

import hashlib
import json
import re
from collections.abc import Callable

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

        row_map, rows_inserted, rows_deleted = _align_axis(
            rows_tok_a.get(sa, {}), rows_tok_b.get(sb, {}), rows_a, rows_b, "row"
        )
        col_map, cols_inserted, cols_deleted = _align_axis(
            cols_tok_a.get(sa, {}), cols_tok_b.get(sb, {}), cols_a, cols_b, "col"
        )

        result[sa] = AxisMap(
            row_map=row_map,
            col_map=col_map,
            rows_inserted=rows_inserted,
            rows_deleted=rows_deleted,
            cols_inserted=cols_inserted,
            cols_deleted=cols_deleted,
            deleted_row_keys=frozenset(set(rows_a) - set(row_map)),
            deleted_col_keys=frozenset(set(cols_a) - set(col_map)),
        )

    return result


def _align_axis(
    tok_a: dict[int, list],
    tok_b: dict[int, list],
    keys_a: list[int],
    keys_b: list[int],
    axis: str,
) -> tuple[dict[int, int], list[tuple[int, int]], list[tuple[int, int]]]:
    """Align one axis (rows or cols) of a matched sheet pair.

    Four passes, each confined to the stretches the previous ones left
    unmatched, so an earlier, stricter match is never overridden:

    1. strict: the full content signature (the original alignment);
    2. loose: formulas with this axis's reference coordinates blanked out,
       so a row/column whose relative references now span an inserted or
       deleted one still matches itself (#53);
    3. label: only the non-formula cells (header labels, constants), so a
       column whose formula changed -- e.g. to #REF! after a column it read
       was deleted -- still matches by its header (#53);
    4. identity: a key present on both sides inside the same unmatched
       stretch is a content edit in place, not a delete plus an insert.

    Returns (key_map, inserted_runs, deleted_runs) as alignment_to_maps does.
    """
    strict_a = [_axis_sig(tok_a[k]) for k in keys_a]
    strict_b = [_axis_sig(tok_b[k]) for k in keys_b]
    alignment = lcs_alignment(strict_a, strict_b)
    # The later passes only look inside gaps with keys left on both sides,
    # and only compute signatures for those keys: an unedited axis (every
    # key strictly matched on one side) costs nothing extra.
    if len(alignment) < min(len(keys_a), len(keys_b)):
        alignment = _refine_gaps(
            alignment,
            lambda i: _axis_sig(tok_a[keys_a[i]], loose=axis),
            lambda j: _axis_sig(tok_b[keys_b[j]], loose=axis),
            len(keys_a),
            len(keys_b),
        )
        alignment = _refine_gaps(
            alignment,
            lambda i: _label_sig(tok_a[keys_a[i]], ("a", keys_a[i])),
            lambda j: _label_sig(tok_b[keys_b[j]], ("b", keys_b[j])),
            len(keys_a),
            len(keys_b),
        )
        alignment = _identity_in_gaps(alignment, keys_a, keys_b)
    return alignment_to_maps(alignment, keys_a, keys_b)


def _identity_in_gaps(
    alignment: list[tuple[int, int]],
    keys_a: list[int],
    keys_b: list[int],
) -> list[tuple[int, int]]:
    """Pair equal keys left unmatched within the same gap between matches.

    Replaces a global fallback that paired any key unmatched on both sides
    with itself. After an interior deletion that paired columns across the
    gap with whatever now sat at their old address -- and could cross
    existing matches (A col 47 -> B col 47 while A col 48 -> B col 46),
    which is what placed #53's deletions at the sheet's tail.
    """
    bounds = [(-1, -1), *alignment, (len(keys_a), len(keys_b))]
    out: list[tuple[int, int]] = []
    for (pa, pb), (na, nb) in zip(bounds, bounds[1:], strict=False):
        b_pos = {keys_b[j]: j for j in range(pb + 1, nb)}
        for i in range(pa + 1, na):
            j = b_pos.get(keys_a[i])
            if j is not None:
                out.append((i, j))
        if na < len(keys_a):
            out.append((na, nb))
    return out


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


def _axis_sig(keyed_tokens: list, loose: str | None = None) -> str:
    """Hash a row's (col, token) list — or a column's (row, token) list —
    ordered by the cross-axis coordinate, matching the original ordering of
    a full (row, col) sort filtered to one row/col.

    ``loose`` ("row" or "col") hashes formulas with that axis's coordinate
    blanked out of every reference: see ``_refine_gaps``.
    """
    keyed_tokens.sort(key=lambda kt: kt[0])
    tokens = [t for _, t in keyed_tokens]
    if loose is not None:
        tokens = [_loosen(t, loose) for t in tokens]
    raw = json.dumps(tokens, separators=(",", ":"), sort_keys=True)
    return hashlib.sha256(raw.encode()).hexdigest()[:16]


# ---------------------------------------------------------------------------
# Gap refinement (issue #53)
# ---------------------------------------------------------------------------

# One R1C1 reference: an optional row part and an optional column part, at
# least one present. The lookarounds keep it out of names and function calls
# (ROUND(, COUNT(, Proj_RC). String literals are skipped by _loosen.
_R1C1_REF = re.compile(
    r"(?<![A-Za-z0-9_.!])"
    r"(?=[RC])"
    r"(?P<row>R(?:\[-?\d+\]|\d+)?)?"
    r"(?P<col>C(?:\[-?\d+\]|\d+)?)?"
    r"(?![A-Za-z0-9_(\[])"
)
_STRING_LITERAL = re.compile(r'"[^"]*"')


def _label_sig(keyed_tokens: list, own: tuple[str, int]) -> str:
    """Hash only the non-formula cells (labels, constants) of a row/column.

    A row/column with none gets a signature unique to itself (``own``), so
    rows/columns cannot match each other on having no labels at all.
    """
    values = [t for _, t in sorted(keyed_tokens, key=lambda kt: kt[0]) if t[0] == "V"]
    if not values:
        return f"unlabelled:{own[0]}:{own[1]}"
    raw = json.dumps(values, separators=(",", ":"), sort_keys=True)
    return hashlib.sha256(raw.encode()).hexdigest()[:16]


def _loosen_formula(formula: str, axis: str) -> str:
    """Blank the ``axis`` coordinate ("row" or "col") of every reference."""

    def sub(m: re.Match[str]) -> str:
        row, col = m.group("row"), m.group("col")
        if not row and not col:
            return m.group(0)
        if axis == "col":
            return (row or "") + ("C*" if col else "")
        return ("R*" if row else "") + (col or "")

    parts: list[str] = []
    pos = 0
    for lit in _STRING_LITERAL.finditer(formula):
        parts.append(_R1C1_REF.sub(sub, formula[pos : lit.start()]))
        parts.append(lit.group(0))
        pos = lit.end()
    parts.append(_R1C1_REF.sub(sub, formula[pos:]))
    return "".join(parts)


def _loosen(token: tuple, axis: str) -> tuple:
    if token[0] != "F":
        return token
    return (token[0], _loosen_formula(token[1], axis), *token[2:])


def _refine_gaps(
    alignment: list[tuple[int, int]],
    sig_a: Callable[[int], str],
    sig_b: Callable[[int], str],
    len_a: int,
    len_b: int,
) -> list[tuple[int, int]]:
    """Re-align each unmatched stretch between strict matches on loose signatures.

    Deleting an interior column changes the R1C1 text of every formula that
    reaches across the gap (RC[-5] becomes RC[-3]), so every column to its
    right gets a new strict signature and the strict LCS matches none of
    them. The identity fallback then paired them by position, which put
    the deletion at the sheet's tail and turned every shifted column into a
    "label change" (#53). Loose signatures ignore the coordinate on the
    deleted axis, so a column keeps its signature when it moves; its header
    label and values still tell it apart from its neighbours.

    Strict matches are kept as they are: only the stretches between them
    (and before the first / after the last) are re-aligned. ``sig_a(i)`` /
    ``sig_b(j)`` give the signature of the i-th / j-th key, and are only
    called for keys inside such a stretch.
    """
    bounds = [(-1, -1), *alignment, (len_a, len_b)]
    refined: list[tuple[int, int]] = []
    for (pa, pb), (na, nb) in zip(bounds, bounds[1:], strict=False):
        if na - pa > 1 and nb - pb > 1:
            gap = lcs_alignment(
                [sig_a(i) for i in range(pa + 1, na)],
                [sig_b(j) for j in range(pb + 1, nb)],
            )
            refined.extend((pa + 1 + i, pb + 1 + j) for i, j in gap)
        if na < len_a:
            refined.append((na, nb))
    return refined
