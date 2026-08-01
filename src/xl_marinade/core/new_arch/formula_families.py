# ABOUTME: Deterministic formula family detection grouping bindings by canonical_a1 on one sheet.
# ABOUTME: Groups non-adjacent bindings whose formulas share the same canonical A1 signature into families.

"""
Formula Family Detection

Groups bindings that share the same canonical_a1 formula signature on the same
sheet into formula families. This addresses the common actuarial pattern where
repeating calculation blocks (e.g., one per asset) produce many separate
bindings that are semantically the same variable.

Design:
- Grouping key: (sheet_id, formula_canonical_a1) from the bindings + formulas tables
- Bindings whose formula has formula_canonical_a1 IS NULL (meaningfulness gate)
  are excluded from family-building; they remain singletons by design.
- Deterministic family IDs using SHA-256 hashing with 'ff_' prefix
- Representative binding: topmost-leftmost member (smallest row, then col)
- Minimum family size: 2 (singletons are not families)

The `formula_id` column persisted on `formula_families` is the
representative binding's formula_id. Multiple formula_ids may collapse into
the same canonical_a1 cluster; the representative's formula_id stands in for
downstream R1C1 lookups.

Algorithm:
1. GROUP BY (sheet_id, canonical_a1) on bindings joined to formulas
   with HAVING COUNT(*) >= 2
2. For each group, fetch member positions to determine ordinals and representative
3. Compute stable family_id from (sheet_id, canonical_a1)
4. Bulk insert into formula_families and formula_family_members
"""

import hashlib
import sqlite3


def _compute_family_id(sheet_id: int, canonical_a1: str) -> str:
    """
    Compute stable family ID from grouping key.

    Args:
        sheet_id: Sheet ID
        canonical_a1: Canonical A1 formula signature (from formulas.formula_canonical_a1)

    Returns:
        Family ID with prefix 'ff_' and 16-char hex digest
    """
    key = f"{sheet_id}:{canonical_a1}"
    digest = hashlib.sha256(key.encode("utf-8")).hexdigest()[:16]
    return f"ff_{digest}"


def extract_formula_families(*, conn: sqlite3.Connection) -> None:
    """
    Extract formula families from bindings.

    This is a deterministic post-pass that runs after grouping/refinement.
    It groups bindings sharing the same (sheet_id, canonical_a1) into families.

    Args:
        conn: Database connection (must have bindings + formulas tables populated)

    Side effects:
        Writes to formula_families and formula_family_members tables
    """
    # Step 1: Find all canonical_a1 groups with 2+ members on each sheet
    groups = conn.execute("""
        SELECT
            b.sheet_id,
            f.formula_canonical_a1,
            COUNT(*) AS member_count
        FROM bindings b
        JOIN formulas f ON b.formula_id = f.formula_id
        WHERE b.formula_id IS NOT NULL
          AND f.formula_canonical_a1 IS NOT NULL
        GROUP BY b.sheet_id, f.formula_canonical_a1
        HAVING COUNT(*) >= 2
        ORDER BY b.sheet_id, f.formula_canonical_a1
    """).fetchall()

    if not groups:
        return

    families_to_insert = []
    members_to_insert = []

    for sheet_id, canonical_a1, _member_count in groups:
        # Step 2: Fetch member bindings sorted by position
        member_rows = conn.execute(
            """
            SELECT b.binding_id, b.formula_id, c.row, c.col
            FROM bindings b
            JOIN formulas f ON b.formula_id = f.formula_id
            JOIN cells c ON b.top_left_cell_id = c.cell_id
            WHERE b.sheet_id = ? AND f.formula_canonical_a1 = ?
            ORDER BY c.row, c.col
        """,
            (sheet_id, canonical_a1),
        ).fetchall()

        if len(member_rows) < 2:
            continue

        # Step 3: Compute family ID
        family_id = _compute_family_id(sheet_id, canonical_a1)

        # Representative is first member (topmost-leftmost).
        # Persisted formula_id is the representative binding's formula_id;
        # other members may have different formula_ids that share canonical_a1.
        representative_binding_id, representative_formula_id, _, _ = member_rows[0]

        families_to_insert.append(
            (
                family_id,
                sheet_id,
                representative_formula_id,
                len(member_rows),
                representative_binding_id,
            )
        )

        # Step 4: Create member entries with stable ordinals
        for ordinal, (binding_id, _fid, _row, _col) in enumerate(member_rows):
            members_to_insert.append(
                (
                    family_id,
                    ordinal,
                    binding_id,
                )
            )

    # Step 5: Write to database
    if families_to_insert:
        conn.executemany(
            """
            INSERT INTO formula_families (
                family_id, sheet_id, formula_id,
                member_count, representative_binding_id
            ) VALUES (?, ?, ?, ?, ?)
        """,
            families_to_insert,
        )

    if members_to_insert:
        conn.executemany(
            """
            INSERT INTO formula_family_members (
                family_id, ordinal, binding_id
            ) VALUES (?, ?, ?)
        """,
            members_to_insert,
        )

    conn.commit()
