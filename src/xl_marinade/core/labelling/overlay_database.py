# ABOUTME: Semantic overlay database persistence layer
# ABOUTME: Handles creating, writing, and loading overlay databases with IR attachment

import json
import sqlite3
from pathlib import Path

from .mutation_engine import DETERMINISTIC_TIMESTAMP, BindingOverlay, OverlayState

# Schema and version constants
SCHEMA_SQL_PATH = Path(__file__).parent / "overlay_schema.sql"
OVERLAY_VERSION = "0.3"


def _bool_to_sqlite(value: bool | None) -> int | None:
    """Convert Python bool to SQLite INTEGER (NULL, 0, or 1).

    SQLite doesn't have a native boolean type, so we represent:
    - None → NULL
    - False → 0
    - True → 1

    Args:
        value: Python boolean or None

    Returns:
        SQLite integer representation or None
    """
    if value is None:
        return None
    return 1 if value else 0


def initialize_overlay_db(ir_db_path: str) -> sqlite3.Connection:
    """
    Create empty semantic overlay with entry for each IR binding.

    Args:
        ir_db_path: Path to Phase 1 IR database

    Returns:
        In-memory SQLite connection with initialized overlay

    Raises:
        FileNotFoundError: If IR database doesn't exist
    """
    if not Path(ir_db_path).exists():
        raise FileNotFoundError(f"IR database not found: {ir_db_path}")

    # Create in-memory overlay
    overlay_conn = sqlite3.connect(":memory:")
    overlay_conn.execute("PRAGMA foreign_keys = ON")

    # Load and execute schema
    with open(SCHEMA_SQL_PATH, encoding="utf-8") as f:
        schema = f.read()
    overlay_conn.executescript(schema)

    # Connect to IR (read-only)
    ir_conn = sqlite3.connect(f"file:{ir_db_path}?mode=ro", uri=True)

    # Insert metadata - try fast schema first, fall back to legacy
    try:
        ir_guid = ir_conn.execute(
            "SELECT value FROM ir_metadata WHERE key = 'workbook_sha256' LIMIT 1"
        ).fetchone()[0]
    except sqlite3.OperationalError:
        # Fall back to legacy schema
        ir_guid = ir_conn.execute("SELECT workbook_guid FROM meta LIMIT 1").fetchone()[0]

    created_timestamp = DETERMINISTIC_TIMESTAMP

    overlay_conn.execute("INSERT INTO metadata VALUES (?, ?)", ("overlay_version", OVERLAY_VERSION))
    overlay_conn.execute("INSERT INTO metadata VALUES (?, ?)", ("ir_db_path", ir_db_path))
    overlay_conn.execute("INSERT INTO metadata VALUES (?, ?)", ("ir_workbook_guid", ir_guid))
    overlay_conn.execute("INSERT INTO metadata VALUES (?, ?)", ("created_at", created_timestamp))

    # Create semantic_variables entry for each IR binding
    # Try fast schema first, fall back to legacy
    try:
        bindings = ir_conn.execute("SELECT binding_id FROM agent_bindings").fetchall()
    except sqlite3.OperationalError:
        # Fall back to legacy schema
        bindings = ir_conn.execute("SELECT binding_id FROM bindings").fetchall()

    for (binding_id,) in bindings:
        overlay_conn.execute(
            "INSERT INTO semantic_variables (binding_id, is_active, is_composite) VALUES (?, 1, 0)",
            (binding_id,),
        )

    ir_conn.close()
    overlay_conn.commit()

    return overlay_conn


def attach_ir_to_overlay(overlay_conn: sqlite3.Connection, ir_db_path: str) -> None:
    """
    Attach Phase 1 IR as read-only database for queries.

    Args:
        overlay_conn: Overlay database connection
        ir_db_path: Path to Phase 1 IR database
    """
    # Validate path exists before attaching
    if not Path(ir_db_path).exists():
        raise FileNotFoundError(f"IR database not found: {ir_db_path}")

    overlay_conn.execute(f"ATTACH DATABASE 'file:{ir_db_path}?mode=ro' AS ir")


def write_overlay_to_db(
    overlay: OverlayState, mutations_path: str, ir_db_path: str, output_path: str
) -> None:
    """
    Write overlay state to SQLite database file.

    ⚠️ ARCHITECTURAL BOUNDARY WARNING ⚠️

    This function should ONLY be called AFTER replay_mutations() has produced an
    OverlayState projection from mutations.json. DO NOT use this function to apply
    manual database updates or bypass the mutation log.

    REQUIRED WORKFLOW:
      1. Create/modify mutations via MutationLogger
      2. Save mutations to mutations.json
      3. Call replay_mutations(ir_db, mutations.json) → OverlayState
      4. Call write_overlay_to_db(overlay, ...) ← YOU ARE HERE

    This enforces Single Source of Truth: mutations.json is authoritative, the
    database is a materialized projection.

    INCORRECT USAGE (will cause state divergence):
      ❌ Manually INSERT/UPDATE semantic_overlay.db rows
      ❌ Call this function with hand-crafted OverlayState objects
      ✅ Always use: replay_mutations() → write_overlay_to_db()

    Args:
        overlay: OverlayState from mutation replay (read-only projection)
        mutations_path: Path to mutations.json (for provenance metadata)
        ir_db_path: Path to Phase 1 IR (for provenance metadata)
        output_path: Where to write semantic_overlay.db

    Raises:
        FileNotFoundError: If ir_db_path doesn't exist
        ValueError: If overlay appears to be hand-crafted (no mutations_applied)
    """
    # Runtime assertion: verify overlay came from replay (has mutations)
    if not overlay.mutations_applied:
        raise ValueError(
            "write_overlay_to_db called with empty mutations_applied. "
            "This function requires an OverlayState from replay_mutations(), "
            "not a hand-crafted object. See docstring for correct usage."
        )
    # Initialize empty overlay database
    overlay_conn = initialize_overlay_db(ir_db_path)

    # Update metadata
    overlay_conn.execute(
        "UPDATE metadata SET value = ? WHERE key = 'created_at'", (DETERMINISTIC_TIMESTAMP,)
    )
    total_mutations = str(len(overlay.mutations_applied))

    overlay_conn.execute("INSERT INTO metadata VALUES (?, ?)", ("mutations_path", mutations_path))
    overlay_conn.execute(
        "INSERT INTO metadata VALUES (?, ?)", ("total_mutations_applied", total_mutations)
    )
    overlay_conn.execute("INSERT INTO metadata VALUES (?, ?)", ("sprint", "1"))

    # Write mutation_log FIRST to satisfy FK constraints
    for mutation in overlay.mutations_applied:
        insert_sql = (
            "INSERT INTO mutation_log "
            "(mutation_id, timestamp, action, parameters_json, metadata_json) "
            "VALUES (?, ?, ?, ?, ?)"
        )
        overlay_conn.execute(
            insert_sql,
            (
                mutation["mutation_id"],
                mutation["timestamp"],
                mutation["action"],
                json.dumps(mutation["parameters"]),
                json.dumps(mutation.get("metadata")) if mutation.get("metadata") else None,
            ),
        )

    # Write semantic_variables (update or insert)
    for binding_id, binding_overlay in overlay.bindings.items():
        # Check if exists (might be composite/virtual binding)
        exists = overlay_conn.execute(
            "SELECT 1 FROM semantic_variables WHERE binding_id = ?", (binding_id,)
        ).fetchone()

        if exists:
            overlay_conn.execute(
                """UPDATE semantic_variables
                   SET label = ?, label_source = ?,
                       actuarial_class = ?, actuarial_class_reasoning = ?,
                       actuarial_class_confidence = ?,
                       is_active = ?, is_composite = ?, superseded_by = ?,
                       reconciliation_required = ?, reconciliation_rationale = ?,
                       label_confidence = ?, classification_confidence = ?,
                       is_orphan = ?
                   WHERE binding_id = ?""",
                (
                    binding_overlay.label,
                    binding_overlay.label_source,
                    binding_overlay.actuarial_class,
                    binding_overlay.actuarial_class_reasoning,
                    binding_overlay.actuarial_class_confidence,
                    binding_overlay.is_active,
                    binding_overlay.is_composite,
                    binding_overlay.superseded_by,
                    _bool_to_sqlite(binding_overlay.reconciliation_required),
                    binding_overlay.reconciliation_rationale,
                    binding_overlay.label_confidence,
                    binding_overlay.classification_confidence,
                    _bool_to_sqlite(binding_overlay.is_orphan),
                    binding_id,
                ),
            )
        else:
            # Insert new virtual/composite binding
            overlay_conn.execute(
                """INSERT INTO semantic_variables
                   (binding_id, label, label_source, actuarial_class,
                    actuarial_class_reasoning, actuarial_class_confidence,
                    is_active, is_composite, superseded_by,
                    reconciliation_required, reconciliation_rationale,
                    label_confidence, classification_confidence, is_orphan)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    binding_id,
                    binding_overlay.label,
                    binding_overlay.label_source,
                    binding_overlay.actuarial_class,
                    binding_overlay.actuarial_class_reasoning,
                    binding_overlay.actuarial_class_confidence,
                    binding_overlay.is_active,
                    binding_overlay.is_composite,
                    binding_overlay.superseded_by,
                    _bool_to_sqlite(binding_overlay.reconciliation_required),
                    binding_overlay.reconciliation_rationale,
                    binding_overlay.label_confidence,
                    binding_overlay.classification_confidence,
                    _bool_to_sqlite(binding_overlay.is_orphan),
                ),
            )

        # Write composite_bindings if composite
        if binding_overlay.is_composite and binding_overlay.composite_members:
            for ordinal, member_id in enumerate(binding_overlay.composite_members):
                overlay_conn.execute(
                    """INSERT INTO composite_bindings
                       (composite_id, ir_binding_id, ordinal)
                       VALUES (?, ?, ?)""",
                    (binding_id, member_id, ordinal),
                )

    overlay_conn.commit()

    # Write to disk
    disk_conn = sqlite3.connect(output_path)
    overlay_conn.backup(disk_conn)
    disk_conn.close()
    overlay_conn.close()


def load_overlay_from_db(overlay_path: str, ir_db_path: str) -> OverlayState:
    """
    Load overlay state from database into memory.

    Args:
        overlay_path: Path to semantic_overlay.db
        ir_db_path: Path to Phase 1 IR (for attachment)

    Returns:
        OverlayState reconstructed from database

    Raises:
        ValueError: If overlay version unsupported
    """
    overlay_conn = sqlite3.connect(overlay_path)

    # Check version (allow v0.2 for backward compatibility)
    version = overlay_conn.execute(
        "SELECT value FROM metadata WHERE key = 'overlay_version'"
    ).fetchone()[0]
    if version not in ["0.2", "0.3"]:
        raise ValueError(f"Unsupported overlay version: {version}. Expected 0.2 or 0.3")

    # Load semantic_variables
    overlay = OverlayState()

    # Determine which columns are available based on version
    # This avoids nested try-except and makes control flow explicit
    has_orphan = False
    has_recon = False

    # Try to load is_orphan column if it exists (v0.3+)
    try:
        rows = overlay_conn.execute(
            """SELECT binding_id, label, label_source,
                      actuarial_class, actuarial_class_reasoning, actuarial_class_confidence,
                      is_active, is_composite, superseded_by,
                      reconciliation_required, reconciliation_rationale,
                      is_orphan
               FROM semantic_variables"""
        ).fetchall()
        has_orphan = True
        has_recon = True
    except sqlite3.OperationalError:
        # is_orphan column doesn't exist (v0.2 or earlier)
        # Try to load reconciliation columns if they exist
        try:
            rows = overlay_conn.execute(
                """SELECT binding_id, label, label_source,
                          actuarial_class, actuarial_class_reasoning, actuarial_class_confidence,
                          is_active, is_composite, superseded_by,
                          reconciliation_required, reconciliation_rationale
                   FROM semantic_variables"""
            ).fetchall()
            has_orphan = False
            has_recon = True
        except sqlite3.OperationalError:
            # Columns don't exist yet (v0.1)
            rows = overlay_conn.execute(
                """SELECT binding_id, label, label_source,
                          actuarial_class, actuarial_class_reasoning, actuarial_class_confidence,
                          is_active, is_composite, superseded_by
                   FROM semantic_variables"""
            ).fetchall()
            has_orphan = False
            has_recon = False

    for row in rows:
        if has_orphan:
            (
                binding_id,
                label,
                label_source,
                actuarial_class,
                actuarial_class_reasoning,
                actuarial_class_confidence,
                is_active,
                is_composite,
                superseded_by,
                recon_required,
                recon_rationale,
                is_orphan,
            ) = row
        elif has_recon:
            (
                binding_id,
                label,
                label_source,
                actuarial_class,
                actuarial_class_reasoning,
                actuarial_class_confidence,
                is_active,
                is_composite,
                superseded_by,
                recon_required,
                recon_rationale,
            ) = row
            is_orphan = False  # Default for v0.2
        else:
            (
                binding_id,
                label,
                label_source,
                actuarial_class,
                actuarial_class_reasoning,
                actuarial_class_confidence,
                is_active,
                is_composite,
                superseded_by,
            ) = row
            recon_required = None
            recon_rationale = None
            is_orphan = False  # Default for v0.1

        binding = BindingOverlay(
            binding_id=binding_id,
            label=label,
            label_source=label_source,
            actuarial_class=actuarial_class,
            actuarial_class_reasoning=actuarial_class_reasoning,
            actuarial_class_confidence=actuarial_class_confidence,
            is_active=bool(is_active),
            is_composite=bool(is_composite),
            superseded_by=superseded_by,
            reconciliation_required=bool(recon_required) if recon_required is not None else None,
            reconciliation_rationale=recon_rationale,
            is_orphan=bool(is_orphan),
        )

        # Load composite members if composite
        if is_composite:
            member_rows = overlay_conn.execute(
                """SELECT ir_binding_id FROM composite_bindings
                   WHERE composite_id = ? ORDER BY ordinal""",
                (binding_id,),
            ).fetchall()
            binding.composite_members = [row[0] for row in member_rows]

        overlay.bindings[binding_id] = binding

    # Load mutations_applied
    select_sql = (
        "SELECT mutation_id, timestamp, action, parameters_json, metadata_json "
        "FROM mutation_log ORDER BY mutation_id"
    )
    mutation_rows = overlay_conn.execute(select_sql).fetchall()

    for mutation_id, timestamp, action, params_json, meta_json in mutation_rows:
        mutation = {
            "mutation_id": mutation_id,
            "timestamp": timestamp,
            "action": action,
            "parameters": json.loads(params_json),
        }
        if meta_json:
            mutation["metadata"] = json.loads(meta_json)
        overlay.mutations_applied.append(mutation)

    overlay_conn.close()
    return overlay


def validate_overlay_integrity(overlay_conn: sqlite3.Connection) -> list[str]:
    """
    Run integrity checks on overlay database.

    Args:
        overlay_conn: Overlay database connection (must have IR attached)

    Returns:
        List of error messages (empty if all checks pass)
    """
    errors = []

    # Check 1: All IR bindings have semantic_variables entry
    result = overlay_conn.execute("""
        SELECT COUNT(*) FROM ir.agent_bindings b
        LEFT JOIN semantic_variables sv ON b.binding_id = sv.binding_id
        WHERE sv.binding_id IS NULL
    """).fetchone()[0]
    if result > 0:
        errors.append(f"Missing semantic_variables entries: {result}")

    # Check 2: All label_source values reference valid mutations
    result = overlay_conn.execute("""
        SELECT COUNT(*) FROM semantic_variables sv
        LEFT JOIN mutation_log ml ON sv.label_source = ml.mutation_id
        WHERE sv.label_source IS NOT NULL AND ml.mutation_id IS NULL
    """).fetchone()[0]
    if result > 0:
        errors.append(f"Orphaned label_source references: {result}")

    # Check 3: Mutation IDs are contiguous
    max_id = overlay_conn.execute("SELECT MAX(mutation_id) FROM mutation_log").fetchone()[0]
    if max_id:
        result = overlay_conn.execute(f"""
            WITH RECURSIVE seq(id) AS (
                SELECT 1
                UNION ALL
                SELECT id + 1 FROM seq WHERE id < {max_id}
            )
            SELECT COUNT(*) FROM seq
            LEFT JOIN mutation_log ml ON seq.id = ml.mutation_id
            WHERE ml.mutation_id IS NULL
        """).fetchone()[0]
        if result > 0:
            errors.append(f"Mutation ID sequence has gaps: {result}")

    # Check 4: Metadata completeness
    required_keys = ["overlay_version", "created_at", "ir_db_path", "ir_workbook_guid"]
    result = overlay_conn.execute(
        f"SELECT COUNT(*) FROM metadata WHERE key IN ({','.join('?' * len(required_keys))})",
        required_keys,
    ).fetchone()[0]
    if result < len(required_keys):
        errors.append(f"Incomplete metadata: expected {len(required_keys)}, got {result}")

    return errors
