# ABOUTME: Deterministic table candidate detection from 1D bindings (Sprint 10 Story 7)
# ABOUTME: Groups adjacent column vectors into table candidates for downstream semantic index

"""
Table Candidate Detection

Detects table candidates from 1D bindings (column vectors) by grouping adjacent
columns with consistent row spans. This addresses the common actuarial pattern where
tables are represented as multiple 1D bindings rather than a single 2D binding.

Design:
- Deterministic grouping based on geometry (sheet_id, r1, r2, contiguous columns)
- Stable candidate IDs using SHA-256 hashing
- Confidence scoring with explainable reasons (top 3)
- Supports both 'vector' (single column) and 'grid' (multiple columns) candidates

Algorithm:
1. Group bindings by (sheet_id, r1, r2) for column vectors
2. Sort by c1 within each group
3. Merge contiguous segments (no gaps by default)
4. Create candidates meeting minimum thresholds
5. Assign stable ordinals to member bindings

Reference: docs/phase2_documentation_agent/backlog/sprint10/STORY_sprint10_07_ir_table_candidates_from_1d_bindings.md
"""

import hashlib
import json
import sqlite3


def _compute_candidate_id(sheet_id: int, r1: int, c1: int, r2: int, c2: int, kind: str) -> str:
    """
    Compute stable candidate ID from geometry.

    Args:
        sheet_id: Sheet ID
        r1, c1, r2, c2: Bounding box coordinates
        kind: Candidate kind ('vector' or 'grid')

    Returns:
        Candidate ID with prefix 'tc_' and 16-char hex digest
    """
    key = f"{sheet_id}:{r1}:{c1}:{r2}:{c2}:{kind}"
    digest = hashlib.sha256(key.encode("utf-8")).hexdigest()[:16]
    return f"tc_{digest}"


def _compute_bbox_from_bindings(
    conn: sqlite3.Connection, binding_ids: list[str]
) -> tuple[int, int, int, int, int]:
    """
    Compute bounding box from binding list.

    Args:
        conn: Database connection
        binding_ids: List of binding IDs

    Returns:
        Tuple of (sheet_id, r1, c1, r2, c2)
    """
    if not binding_ids:
        raise ValueError("binding_ids cannot be empty")

    # Get bounding boxes for all bindings
    placeholders = ",".join("?" * len(binding_ids))
    rows = conn.execute(
        f"""
        SELECT 
            b.sheet_id,
            c.row AS r,
            c.col AS c
        FROM bindings b
        JOIN cells c ON b.top_left_cell_id = c.cell_id
        WHERE b.binding_id IN ({placeholders})
    """,
        binding_ids,
    ).fetchall()

    if not rows:
        raise ValueError("No cells found for bindings")

    sheet_id = rows[0][0]

    # Compute bbox from all bindings
    min_r = min(r for _, r, _ in rows)
    min_c = min(c for _, _, c in rows)

    # Get shape for each binding to compute max bounds
    shape_rows = conn.execute(
        f"""
        SELECT binding_id, shape_rows, shape_cols
        FROM bindings
        WHERE binding_id IN ({placeholders})
    """,
        binding_ids,
    ).fetchall()

    shape_map = {bid: (sr, sc) for bid, sr, sc in shape_rows}

    max_r = min_r
    max_c = min_c

    for binding_id in binding_ids:
        row = next((r for _, r, c in rows if _ == sheet_id), min_r)
        col = next((c for _, r, c in rows if _ == sheet_id), min_c)
        sr, sc = shape_map.get(binding_id, (1, 1))
        max_r = max(max_r, row + sr - 1)
        max_c = max(max_c, col + sc - 1)

    return sheet_id, min_r, min_c, max_r, max_c


def _col_to_a1(col: int) -> str:
    """Convert column number (1-based) to A1 notation."""
    result = ""
    col = col - 1  # Convert to 0-based
    while col >= 0:
        result = chr(ord("A") + (col % 26)) + result
        col = col // 26 - 1
    return result


def _range_to_a1(r1: int, c1: int, r2: int, c2: int) -> str:
    """Convert range coordinates to A1 notation (no sheet)."""
    start = f"{_col_to_a1(c1)}{r1}"
    if r1 == r2 and c1 == c2:
        return start
    end = f"{_col_to_a1(c2)}{r2}"
    return f"{start}:{end}"


def extract_table_candidates(*, conn: sqlite3.Connection) -> None:
    """
    Extract table candidates from 1D bindings.

    This is a deterministic post-pass that runs after grouping/refinement.
    It groups adjacent column vectors (shape_rows > 1, shape_cols = 1) into
    table candidates.

    Args:
        conn: Database connection (must have bindings table populated)

    Side effects:
        Writes to table_candidates and table_candidate_members tables
    """
    # Configuration (deterministic thresholds)
    MIN_ROWS_FOR_GRID = 3
    MIN_COLS_FOR_GRID = 2
    MIN_LENGTH_FOR_VECTOR = 5

    # Step 1: Load all column vector bindings (shape_rows > 1, shape_cols = 1)
    column_vectors = conn.execute("""
        SELECT 
            b.binding_id,
            b.sheet_id,
            c.row AS r1,
            c.col AS c1,
            b.shape_rows,
            b.shape_cols,
            b.classification
        FROM bindings b
        JOIN cells c ON b.top_left_cell_id = c.cell_id
        WHERE b.shape_rows > 1 AND b.shape_cols = 1
        ORDER BY b.sheet_id, c.row, c.col
    """).fetchall()

    if not column_vectors:
        return

    # Step 2: Group by (sheet_id, r1, r2)
    groups: dict[tuple[int, int, int], list[tuple[str, int, int, str | None]]] = {}

    for binding_id, sheet_id, r1, c1, shape_rows, shape_cols, classification in column_vectors:
        r2 = r1 + shape_rows - 1
        key = (sheet_id, r1, r2)
        groups.setdefault(key, []).append((binding_id, c1, shape_rows, classification))

    # Step 3: For each group, find contiguous segments
    candidates_to_insert = []
    members_to_insert = []

    for (sheet_id, r1, r2), bindings_in_group in groups.items():
        # Sort by c1
        bindings_in_group.sort(key=lambda x: x[1])

        # Find contiguous segments (no gaps)
        segments = []
        current_segment = [bindings_in_group[0]]

        for i in range(1, len(bindings_in_group)):
            prev_binding_id, prev_c1, prev_shape_rows, prev_classification = current_segment[-1]
            curr_binding_id, curr_c1, curr_shape_rows, curr_classification = bindings_in_group[i]

            # Check if contiguous (current c1 == prev c1 + 1)
            if curr_c1 == prev_c1 + 1:
                current_segment.append(bindings_in_group[i])
            else:
                # Gap detected - finalize current segment
                segments.append(current_segment)
                current_segment = [bindings_in_group[i]]

        # Add final segment
        segments.append(current_segment)

        # Step 4: Create candidates from segments
        for segment in segments:
            segment_len = len(segment)
            shape_rows = segment[0][2]

            # Decide candidate kind and whether to create
            if segment_len >= MIN_COLS_FOR_GRID and shape_rows >= MIN_ROWS_FOR_GRID:
                kind = "grid"
                create_candidate = True
            elif segment_len == 1:
                # Single column - check if it qualifies as vector
                binding_id, c1, shape_rows, classification = segment[0]
                if shape_rows >= MIN_LENGTH_FOR_VECTOR or classification in ("input", "assumption"):
                    kind = "vector"
                    create_candidate = True
                else:
                    create_candidate = False
            else:
                # Small segment (< MIN_COLS_FOR_GRID) - skip
                create_candidate = False

            if not create_candidate:
                continue

            # Compute bounding box
            c1 = segment[0][1]
            c2 = segment[-1][1]

            # Compute candidate ID
            candidate_id = _compute_candidate_id(sheet_id, r1, c1, r2, c2, kind)

            # Compute range_a1
            range_a1 = _range_to_a1(r1, c1, r2, c2)

            # Compute confidence and reasons
            confidence = 0.8  # Base confidence
            reasons = []

            if kind == "grid":
                reasons.append(f"contiguous_{segment_len}_cols")
                if shape_rows >= 10:
                    reasons.append("sufficient_rows")
                    confidence = 0.9
            else:  # vector
                if shape_rows >= MIN_LENGTH_FOR_VECTOR:
                    reasons.append(f"length_{shape_rows}")
                if segment[0][3] in ("input", "assumption"):
                    reasons.append(f"classification_{segment[0][3]}")
                    confidence = 0.85

            reasons.append(f"same_row_span_{r1}_{r2}")

            # Limit to top 3 reasons
            reasons_top3 = reasons[:3]
            reasons_json = json.dumps(reasons_top3)

            # Add candidate
            candidates_to_insert.append(
                (candidate_id, sheet_id, kind, r1, c1, r2, c2, range_a1, confidence, reasons_json)
            )

            # Add members with stable ordinals
            for ordinal, (binding_id, _, _, _) in enumerate(segment):
                members_to_insert.append(
                    (
                        candidate_id,
                        ordinal,
                        binding_id,
                        "values",  # role_hint (all are values for now)
                    )
                )

    # Step 5: Write to database
    if candidates_to_insert:
        conn.executemany(
            """
            INSERT INTO table_candidates (
                candidate_id, sheet_id, kind,
                r1, c1, r2, c2,
                range_a1, confidence, reasons_top3_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
            candidates_to_insert,
        )

    if members_to_insert:
        conn.executemany(
            """
            INSERT INTO table_candidate_members (
                candidate_id, ordinal, binding_id, role_hint
            ) VALUES (?, ?, ?, ?)
        """,
            members_to_insert,
        )

    conn.commit()
