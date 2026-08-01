# ABOUTME: Lexicographically minimal Longest Common Subsequence alignment.
# ABOUTME: Returns the unique alignment with the smallest sequence of matched index pairs.

from __future__ import annotations

from collections.abc import Hashable, Sequence


def lcs_alignment(
    seq_a: Sequence[Hashable],
    seq_b: Sequence[Hashable],
) -> list[tuple[int, int]]:
    """Compute the lexicographically minimal LCS alignment.

    Returns a list of (index_in_a, index_in_b) pairs representing the
    longest common subsequence. Among all maximum-cardinality alignments,
    this returns the one with the lexicographically smallest sequence of
    matched index pairs.

    Args:
        seq_a: First sequence of hashable elements.
        seq_b: Second sequence of hashable elements.

    Returns:
        List of (i, j) pairs where seq_a[i] == seq_b[j], in ascending order.
    """
    n = len(seq_a)
    m = len(seq_b)

    if n == 0 or m == 0:
        return []

    # Trim the common prefix and suffix before the O(n*m) DP. Equal ends are
    # always part of some maximum-cardinality LCS, and the lexicographically
    # minimal alignment always takes the earliest available match, so the
    # trimmed solution composes back to the exact untrimmed answer. On real
    # workbook diffs the sequences are near-identical, which turns the DP
    # from O(n*m) into O(changed region squared).
    p = 0
    lim = min(n, m)
    while p < lim and seq_a[p] == seq_b[p]:
        p += 1
    s = 0
    while s < lim - p and seq_a[n - 1 - s] == seq_b[m - 1 - s]:
        s += 1

    mid = _lcs_dp(seq_a[p : n - s], seq_b[p : m - s])
    return (
        [(i, i) for i in range(p)]
        + [(i + p, j + p) for i, j in mid]
        + [(n - s + k, m - s + k) for k in range(s)]
    )


def _lcs_dp(
    seq_a: Sequence[Hashable],
    seq_b: Sequence[Hashable],
) -> list[tuple[int, int]]:
    """Core O(n*m) lexicographically-minimal LCS on pre-trimmed sequences."""
    n = len(seq_a)
    m = len(seq_b)

    if n == 0 or m == 0:
        return []

    # Build DP table: L[i][j] = LCS length of seq_a[:i] and seq_b[:j]
    L = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            if seq_a[i - 1] == seq_b[j - 1]:
                L[i][j] = L[i - 1][j - 1] + 1
            else:
                L[i][j] = max(L[i - 1][j], L[i][j - 1])

    # Backtrack to recover lexicographically minimal alignment.
    # "Lexicographically minimal" means: prefer matching at the smallest
    # (i, j) index pair possible at each step from the top-left.
    #
    # We backtrack from (n, m) but build the result in reverse.
    # At each cell: if seq_a[i-1] == seq_b[j-1] AND this match is part of
    # an optimal solution (L[i][j] == L[i-1][j-1] + 1), take the match.
    # Otherwise, prefer going UP (i-1, j) over LEFT (i, j-1) when tied.
    # Going UP keeps i small → lexicographically smaller indices.
    result = []
    i, j = n, m
    while i > 0 and j > 0:
        if seq_a[i - 1] == seq_b[j - 1] and L[i][j] == L[i - 1][j - 1] + 1:
            result.append((i - 1, j - 1))
            i -= 1
            j -= 1
        elif L[i - 1][j] >= L[i][j - 1]:
            # Prefer UP (smaller i) when tied → lex-minimal
            i -= 1
        else:
            j -= 1

    result.reverse()
    return result


def alignment_to_maps(
    alignment: list[tuple[int, int]],
    keys_a: Sequence[int],
    keys_b: Sequence[int],
) -> tuple[dict[int, int], list[tuple[int, int]], list[tuple[int, int]]]:
    """Convert an LCS alignment to a key map plus inserted/deleted runs.

    Args:
        alignment: List of (idx_a, idx_b) pairs from lcs_alignment().
        keys_a: Original keys (e.g., row numbers) for sequence A.
        keys_b: Original keys (e.g., row numbers) for sequence B.

    Returns:
        Tuple of:
        - key_map: dict mapping keys_a[i] -> keys_b[j] for matched pairs
        - inserted_runs: list of (first_key_in_b, count) for unmatched B runs
        - deleted_runs: list of (first_key_in_a, count) for unmatched A runs
    """
    matched_a_indices = {ia for ia, _ in alignment}
    matched_b_indices = {ib for _, ib in alignment}

    # Key map
    key_map = {}
    for ia, ib in alignment:
        key_map[keys_a[ia]] = keys_b[ib]

    # Inserted runs (in B, not matched)
    inserted_runs = _find_runs(keys_b, matched_b_indices)

    # Deleted runs (in A, not matched)
    deleted_runs = _find_runs(keys_a, matched_a_indices)

    return key_map, inserted_runs, deleted_runs


def _find_runs(
    keys: Sequence[int],
    matched_indices: set[int],
) -> list[tuple[int, int]]:
    """Find contiguous runs of unmatched indices, return as (first_key, count)."""
    runs = []
    i = 0
    n = len(keys)
    while i < n:
        if i not in matched_indices:
            start = i
            while i < n and i not in matched_indices:
                i += 1
            runs.append((keys[start], i - start))
        else:
            i += 1
    return runs
