# ABOUTME: Stage 4a — Binding matching across canonicalized IR models.
# ABOUTME: Four sub-passes: exact position, moved fingerprint, overlap graph, component classification.

from __future__ import annotations

from collections import defaultdict

from xl_marinade.core.ir_diff.model import (
    BindingKey,
    BindingMatch,
    CellKey,
    IRModel,
)


def match_bindings(a: IRModel, b: IRModel) -> BindingMatch:
    """Match bindings between canonicalized A and B models.

    Four sub-passes (B1-B4):
    - B1: Exact canonical-position match (sheet + top_left + shape + type)
    - B2: Moved-binding match (unique fingerprint on same sheet)
    - B3: Overlap graph from cell membership
    - B4: Component classification (1x1, 1xN, Nx1, NxM)

    Args:
        a: Canonicalized Version A model.
        b: Canonicalized Version B model.

    Returns:
        BindingMatch with matched pairs, match types, and unmatched bindings.
    """
    result = BindingMatch()
    unmatched_a = set(a.bindings.keys())
    unmatched_b = set(b.bindings.keys())

    # --- B1: Exact canonical-position match ---
    # Match bindings with identical (sheet, top_left, shape, type)
    a_by_pos = _index_by_position_type(a, unmatched_a)
    b_by_pos = _index_by_position_type(b, unmatched_b)

    for pos_key in sorted(a_by_pos.keys() & b_by_pos.keys()):
        ak = a_by_pos[pos_key]
        bk = b_by_pos[pos_key]
        result.matched[ak] = bk
        result.match_type[ak] = "exact"
        unmatched_a.discard(ak)
        unmatched_b.discard(bk)

    # --- B2: Moved-binding match ---
    # For unmatched bindings, match by unique (sheet, shape, type, fingerprint)
    a_by_fp: dict[tuple, list[BindingKey]] = defaultdict(list)
    b_by_fp: dict[tuple, list[BindingKey]] = defaultdict(list)

    for ak in sorted(unmatched_a):
        adesc = a.bindings[ak]
        fp_key = (ak.sheet, ak.shape_rows, ak.shape_cols, adesc.binding_type, adesc.binding_fp_rel)
        a_by_fp[fp_key].append(ak)

    for bk in sorted(unmatched_b):
        bdesc = b.bindings[bk]
        fp_key = (bk.sheet, bk.shape_rows, bk.shape_cols, bdesc.binding_type, bdesc.binding_fp_rel)
        b_by_fp[fp_key].append(bk)

    for fp_key in sorted(a_by_fp.keys() & b_by_fp.keys()):
        a_list = a_by_fp[fp_key]
        b_list = b_by_fp[fp_key]
        # Only match if fingerprint is unique on both sides
        if len(a_list) == 1 and len(b_list) == 1:
            ak = a_list[0]
            bk = b_list[0]
            result.matched[ak] = bk
            result.match_type[ak] = "moved"
            unmatched_a.discard(ak)
            unmatched_b.discard(bk)

    # --- B3/B4: Overlap graph and component classification ---
    if unmatched_a and unmatched_b:
        _overlap_match(a, b, unmatched_a, unmatched_b, result)

    result.removed = sorted(unmatched_a)
    result.added = sorted(unmatched_b)

    return result


def _index_by_position_type(
    model: IRModel,
    keys: set[BindingKey],
) -> dict[tuple, BindingKey]:
    """Index bindings by (sheet, row, col, shape_rows, shape_cols, type)."""
    index: dict[tuple, BindingKey] = {}
    for bk in keys:
        bdesc = model.bindings[bk]
        pos_key = (
            bk.sheet,
            bk.top_left_row,
            bk.top_left_col,
            bk.shape_rows,
            bk.shape_cols,
            bdesc.binding_type,
        )
        index[pos_key] = bk  # Assumes unique per position
    return index


def _overlap_match(
    a: IRModel,
    b: IRModel,
    unmatched_a: set[BindingKey],
    unmatched_b: set[BindingKey],
    result: BindingMatch,
) -> None:
    """Build overlap graph and classify components.

    For each unmatched binding pair (a, b), compute cell-membership overlap.
    Build bipartite graph, find connected components, classify as
    1x1, 1xN (split), Nx1 (merge), NxM (restructure).
    """
    # Build cell -> binding maps for unmatched bindings only
    a_cell_to_bind: dict[CellKey, BindingKey] = {}
    for ck, bkeys in a.cell_to_binding.items():
        for bk in bkeys:
            if bk in unmatched_a:
                a_cell_to_bind[ck] = bk

    b_cell_to_bind: dict[CellKey, BindingKey] = {}
    for ck, bkeys in b.cell_to_binding.items():
        for bk in bkeys:
            if bk in unmatched_b:
                b_cell_to_bind[ck] = bk

    # Find overlapping cells → build bipartite edges with scores
    overlap_counts: dict[tuple[BindingKey, BindingKey], int] = defaultdict(int)

    for ck in a_cell_to_bind:
        if ck in b_cell_to_bind:
            ak = a_cell_to_bind[ck]
            bk = b_cell_to_bind[ck]
            overlap_counts[(ak, bk)] += 1

    if not overlap_counts:
        # No overlap — all remaining are pure adds/removes
        return

    # Build edges with scores
    edges: list[tuple[BindingKey, BindingKey, int, float, float, float]] = []
    for (ak, bk), overlap in overlap_counts.items():
        a_size = len(a.bindings[ak].members_by_offset) or 1
        b_size = len(b.bindings[bk].members_by_offset) or 1
        coverage_a = overlap / a_size
        coverage_b = overlap / b_size
        union_size = a_size + b_size - overlap
        jaccard = overlap / union_size if union_size > 0 else 0.0
        edges.append((ak, bk, overlap, coverage_a, coverage_b, jaccard))

    # Build adjacency for component detection
    a_neighbors: dict[BindingKey, set[BindingKey]] = defaultdict(set)
    b_neighbors: dict[BindingKey, set[BindingKey]] = defaultdict(set)
    for ak, bk, *_ in edges:
        a_neighbors[ak].add(bk)
        b_neighbors[bk].add(ak)

    # Find connected components via BFS
    visited_a: set[BindingKey] = set()
    visited_b: set[BindingKey] = set()

    components: list[tuple[set[BindingKey], set[BindingKey]]] = []

    for ak in sorted(a_neighbors.keys()):
        if ak in visited_a:
            continue
        # BFS
        comp_a: set[BindingKey] = set()
        comp_b: set[BindingKey] = set()
        queue_a = [ak]
        while queue_a:
            next_queue_a: list[BindingKey] = []
            for node_a in queue_a:
                if node_a in visited_a:
                    continue
                visited_a.add(node_a)
                comp_a.add(node_a)
                for nb in sorted(a_neighbors[node_a]):
                    if nb not in visited_b:
                        visited_b.add(nb)
                        comp_b.add(nb)
                        for nb2 in sorted(b_neighbors[nb]):
                            if nb2 not in visited_a:
                                next_queue_a.append(nb2)
            queue_a = next_queue_a
        components.append((comp_a, comp_b))

    # Also capture B-only nodes with edges (not yet visited)
    for bk in sorted(b_neighbors.keys()):
        if bk not in visited_b:
            comp_b = {bk}
            comp_a: set[BindingKey] = set()
            for nb in sorted(b_neighbors[bk]):
                comp_a.add(nb)
            if comp_a or comp_b:
                components.append((comp_a, comp_b))

    # Classify each component
    for comp_a, comp_b in components:
        _classify_component(comp_a, comp_b, edges, result, unmatched_a, unmatched_b)


def _classify_component(
    comp_a: set[BindingKey],
    comp_b: set[BindingKey],
    all_edges: list[tuple[BindingKey, BindingKey, int, float, float, float]],
    result: BindingMatch,
    unmatched_a: set[BindingKey],
    unmatched_b: set[BindingKey],
) -> None:
    """Classify a connected component and pair bindings."""
    # Filter edges to this component
    comp_edges = [
        (ak, bk, overlap, cov_a, cov_b, jacc)
        for ak, bk, overlap, cov_a, cov_b, jacc in all_edges
        if ak in comp_a and bk in comp_b
    ]

    if not comp_edges:
        return

    n_a = len(comp_a)
    n_b = len(comp_b)

    if n_a == 1 and n_b == 1:
        # 1x1: direct match
        ak = next(iter(comp_a))
        bk = next(iter(comp_b))
        result.matched[ak] = bk
        result.match_type[ak] = "overlap"
        unmatched_a.discard(ak)
        unmatched_b.discard(bk)
        return

    # Sort edges by deterministic score: (overlap desc, coverage_b desc, jaccard desc,
    # then lex by binding keys for tie-break)
    comp_edges.sort(key=lambda e: (-e[2], -e[4], -e[5], e[0], e[1]))

    # Greedy pairing
    paired_a: set[BindingKey] = set()
    paired_b: set[BindingKey] = set()

    for ak, bk, overlap, cov_a, cov_b, jacc in comp_edges:
        if ak in paired_a or bk in paired_b:
            continue
        result.matched[ak] = bk
        result.match_type[ak] = "overlap"
        paired_a.add(ak)
        paired_b.add(bk)
        unmatched_a.discard(ak)
        unmatched_b.discard(bk)

        # Record lineage for splits/merges
        if n_a == 1 and n_b > 1:
            result.lineage[ak] = {"type": "split", "component_size": n_b}
        elif n_a > 1 and n_b == 1:
            result.lineage[ak] = {"type": "merge", "component_size": n_a}
        elif n_a > 1 and n_b > 1:
            result.lineage[ak] = {"type": "restructure", "component_a": n_a, "component_b": n_b}

    # Residual unmatched within this component already stay in unmatched_a/unmatched_b
