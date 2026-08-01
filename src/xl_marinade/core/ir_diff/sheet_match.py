# ABOUTME: Stage 2a — Sheet matching with exact-name pass and rename detection.
# ABOUTME: Uses binding/cell/coord overlap scores with fixed conservative thresholds.

from __future__ import annotations

from collections import Counter

from xl_marinade.core.ir_diff.model import (
    IRModel,
    SheetMatch,
)

# Fixed thresholds (per research report section 4.1)
RENAME_THRESHOLD_BIND = 0.80
RENAME_UNIQUENESS_MARGIN = 0.10


def match_sheets(a: IRModel, b: IRModel) -> SheetMatch:
    """Match sheets between two IR models.

    Pass 1: Exact name match.
    Pass 2: Rename detection for unmatched sheets using overlap scores.

    Args:
        a: Version A IR model.
        b: Version B IR model.

    Returns:
        SheetMatch with matched pairs, renamed pairs, and unmatched sheets.
    """
    result = SheetMatch()

    set_a = set(a.sheet_names)
    set_b = set(b.sheet_names)

    # Pass 1: exact name match
    exact = sorted(set_a & set_b)
    for s in exact:
        result.matched[s] = s

    unmatched_a = sorted(set_a - set_b)
    unmatched_b = sorted(set_b - set_a)

    if not unmatched_a or not unmatched_b:
        result.removed = unmatched_a
        result.added = unmatched_b
        return result

    # Pass 2: rename detection
    # Compute fingerprints per sheet for scoring
    fp_a = {s: _sheet_fingerprint(a, s) for s in unmatched_a}
    fp_b = {s: _sheet_fingerprint(b, s) for s in unmatched_b}

    # Score all pairs
    scores: list[tuple[str, str, float, float, float]] = []
    for sa in unmatched_a:
        for sb in unmatched_b:
            s_bind, s_cell, s_coord = _score_sheet_pair(fp_a[sa], fp_b[sb])
            scores.append((sa, sb, s_bind, s_cell, s_coord))

    # Build best-score maps for mutual-best checking
    best_for_a: dict[str, list[tuple[str, float, float, float]]] = {}
    best_for_b: dict[str, list[tuple[str, float, float, float]]] = {}

    for sa, sb, s_bind, s_cell, s_coord in scores:
        best_for_a.setdefault(sa, []).append((sb, s_bind, s_cell, s_coord))
        best_for_b.setdefault(sb, []).append((sa, s_bind, s_cell, s_coord))

    # Sort candidates by score tuple descending
    for k in best_for_a:
        best_for_a[k].sort(key=lambda x: (x[1], x[2], x[3]), reverse=True)
    for k in best_for_b:
        best_for_b[k].sort(key=lambda x: (x[1], x[2], x[3]), reverse=True)

    # Greedily accept rename pairs
    remaining_a = set(unmatched_a)
    remaining_b = set(unmatched_b)

    while remaining_a and remaining_b:
        best_pair = None
        best_score = (-1.0, -1.0, -1.0)

        for sa in sorted(remaining_a):
            candidates = [c for c in best_for_a.get(sa, []) if c[0] in remaining_b]
            if not candidates:
                continue
            top = candidates[0]
            sb, s_bind, s_cell, s_coord = top

            # Check acceptance criteria
            if s_bind < RENAME_THRESHOLD_BIND:
                continue

            # Check uniqueness margin
            if len(candidates) > 1:
                runner_up = candidates[1]
                if s_bind - runner_up[1] < RENAME_UNIQUENESS_MARGIN:
                    if s_bind == runner_up[1] and s_cell - runner_up[2] < RENAME_UNIQUENESS_MARGIN:
                        continue

            # Check mutual best: sb's best must also be sa
            b_candidates = [c for c in best_for_b.get(sb, []) if c[0] in remaining_a]
            if not b_candidates or b_candidates[0][0] != sa:
                continue

            # Check uniqueness from b's perspective too
            if len(b_candidates) > 1:
                b_runner = b_candidates[1]
                b_top_bind = b_candidates[0][1]
                if b_top_bind - b_runner[1] < RENAME_UNIQUENESS_MARGIN:
                    if (
                        b_top_bind == b_runner[1]
                        and b_candidates[0][2] - b_runner[2] < RENAME_UNIQUENESS_MARGIN
                    ):
                        continue

            score_tuple = (s_bind, s_cell, s_coord)
            if score_tuple > best_score:
                best_score = score_tuple
                best_pair = (sa, sb, s_bind, s_cell, s_coord)

        if best_pair is None:
            break

        sa, sb, s_bind, s_cell, s_coord = best_pair
        result.matched[sa] = sb
        result.renamed.append((sa, sb, s_bind, s_cell, s_coord))
        remaining_a.discard(sa)
        remaining_b.discard(sb)

    result.removed = sorted(remaining_a)
    result.added = sorted(remaining_b)

    return result


# ---------------------------------------------------------------------------
# Sheet fingerprinting for rename scoring
# ---------------------------------------------------------------------------


def _sheet_fingerprint(model: IRModel, sheet: str) -> dict:
    """Build a fingerprint for a sheet for rename scoring."""
    # Binding fingerprints (formula pattern + shape, ignoring absolute position)
    bind_fps: list[tuple] = []
    for bkey, bdesc in model.bindings.items():
        if bkey.sheet == sheet:
            bind_fps.append(
                (
                    bdesc.binding_type,
                    bkey.shape_rows,
                    bkey.shape_cols,
                    bdesc.formula_r1c1 or "__const__",
                )
            )

    # Cell structural tokens (formula + data_type, ignoring position)
    cell_tokens: list[tuple] = []
    for ck, csig in model.cells.items():
        if ck.sheet == sheet:
            cell_tokens.append(
                (
                    csig.formula_r1c1 or "__value__",
                    csig.data_type or "__unknown__",
                )
            )

    # Coordinate tokens (row, col, formula/value kind)
    coord_tokens: list[tuple] = []
    for ck, csig in model.cells.items():
        if ck.sheet == sheet:
            coord_tokens.append(
                (
                    ck.row,
                    ck.col,
                    csig.formula_r1c1 or "__value__",
                    csig.data_type or "__unknown__",
                )
            )

    return {
        "bind_counter": Counter(bind_fps),
        "cell_counter": Counter(cell_tokens),
        "coord_counter": Counter(coord_tokens),
    }


def _multiset_overlap(ca: Counter, cb: Counter) -> float:
    """Compute multiset overlap / max(|a|, |b|)."""
    total_a = sum(ca.values())
    total_b = sum(cb.values())
    if total_a == 0 and total_b == 0:
        return 1.0
    if total_a == 0 or total_b == 0:
        return 0.0
    overlap = sum((ca & cb).values())
    return overlap / max(total_a, total_b)


def _score_sheet_pair(fp_a: dict, fp_b: dict) -> tuple[float, float, float]:
    """Score a sheet pair for rename detection."""
    s_bind = _multiset_overlap(fp_a["bind_counter"], fp_b["bind_counter"])
    s_cell = _multiset_overlap(fp_a["cell_counter"], fp_b["cell_counter"])
    s_coord = _multiset_overlap(fp_a["coord_counter"], fp_b["coord_counter"])
    return s_bind, s_cell, s_coord
