# ABOUTME: SQLite schema with deterministic PRAGMAs per ADR-011 and schema validation.
# ABOUTME: Implements meta/user_roots tables, deterministic workbook_guid, validator.
"""SQLite schema creation and validation."""

import hashlib
import json
import sqlite3
from pathlib import Path
from typing import Any

# Deterministic PRAGMAs per ADR-011
# Note: journal_mode=DELETE (default) persists in database header
# synchronous=OFF is per-connection and must be set on each open
DETERMINISTIC_PRAGMAS = [
    "PRAGMA page_size = 4096",
    "PRAGMA journal_mode = DELETE",  # Persists; deterministic with consistent settings
    "PRAGMA synchronous = OFF",  # Per-connection optimization
    "PRAGMA locking_mode = EXCLUSIVE",
    "PRAGMA temp_store = 2",  # CRITICAL: 2=FILE mode - Use disk for temp tables to prevent 17 GB RAM usage
]

# Per-connection PRAGMAs that must be set when re-opening an existing database
# (page_size and journal_mode persist in database header, don't need re-setting)
PER_CONNECTION_PRAGMAS = [
    "PRAGMA synchronous = OFF",
    "PRAGMA locking_mode = EXCLUSIVE",
    "PRAGMA temp_store = 2",  # CRITICAL: 2=FILE mode - Use disk for temp tables to prevent 17 GB RAM usage
]


# Schema version
IR_VERSION = "2.3.0"

# Fixed timestamps for determinism (per ADR-010)
FIXED_TIMESTAMP = "1970-01-01T00:00:00Z"


def _canonicalize_json_not_null(payload: Any, field_name: str) -> str:
    """
    Canonicalize JSON for NOT NULL columns; raise on invalid or None.

    Used for time_index_candidates.reasons_top3_json and
    binding_time_annotations.reasons_top3_json / evidence_flags_json.
    """
    if payload is None:
        raise ValueError(f"{field_name} is required (cannot be None)")
    if isinstance(payload, str):
        try:
            parsed = json.loads(payload)
        except json.JSONDecodeError as exc:
            raise ValueError(f"{field_name} is not valid JSON") from exc
    else:
        parsed = payload
    try:
        return json.dumps(parsed, sort_keys=True, separators=(",", ":"))
    except TypeError as exc:
        raise ValueError(f"{field_name} contains non-serializable data") from exc


def compute_workbook_guid(workbook_path: Path) -> str:
    """
    Compute deterministic workbook GUID from workbook bytes.

    Per ADR-010: SHA-256 hash of workbook bytes, lowercase hex, 64 characters.

    Args:
        workbook_path: Path to workbook file.

    Returns:
        Hex64 workbook GUID (64 lowercase hex characters).
    """
    workbook_bytes = workbook_path.read_bytes()
    hash_obj = hashlib.sha256(workbook_bytes)
    return hash_obj.hexdigest()  # Returns lowercase hex string (64 chars)


def create_deterministic_connection(db_path: Path) -> sqlite3.Connection:
    """
    Create SQLite connection with deterministic PRAGMAs.

    Per ADR-011: Set fixed page size, disable journal, etc.

    Args:
        db_path: Path to database file.

    Returns:
        SQLite connection with deterministic settings.
    """
    # Remove existing database to ensure clean slate
    if db_path.exists():
        db_path.unlink()

    # Create parent directory if needed
    db_path.parent.mkdir(parents=True, exist_ok=True)

    # Create connection
    conn = sqlite3.connect(str(db_path))

    # Set deterministic PRAGMAs
    cursor = conn.cursor()
    for pragma in DETERMINISTIC_PRAGMAS:
        cursor.execute(pragma)

    return conn


def open_existing_database(db_path: Path) -> sqlite3.Connection:
    """
    Open existing SQLite database with per-connection deterministic PRAGMAs.

    Per ADR-011: Some PRAGMAs persist in database header (page_size, journal_mode),
    but per-connection PRAGMAs (synchronous, locking_mode, temp_store) must be
    set each time a connection is opened.

    This function is used when re-opening a database that was created by
    create_deterministic_connection(). It does NOT delete existing data.

    Args:
        db_path: Path to existing database file.

    Returns:
        SQLite connection with per-connection deterministic settings.

    Raises:
        ValueError: If database file does not exist.
    """
    if not db_path.exists():
        raise ValueError(f"Database does not exist: {db_path}")

    conn = sqlite3.connect(str(db_path))

    # Set per-connection PRAGMAs (persistent ones already in database header)
    cursor = conn.cursor()
    for pragma in PER_CONNECTION_PRAGMAS:
        cursor.execute(pragma)

    return conn


def create_meta_table(conn: sqlite3.Connection) -> None:
    """Create meta table per schema definition."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS meta (
            ir_version TEXT NOT NULL,
            workbook_guid TEXT NOT NULL,
            file_path TEXT NOT NULL,
            excel_build_info TEXT,
            calculation_mode TEXT NOT NULL,
            imported_at TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)
    conn.commit()


def create_user_roots_table(conn: sqlite3.Connection) -> None:
    """Create user_roots table per schema definition."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS user_roots (
            root_id INTEGER PRIMARY KEY,
            sheet TEXT NOT NULL,
            range_a1 TEXT NOT NULL,
            label_hint TEXT
        )
    """)
    conn.commit()


def create_time_index_candidates_table(conn: sqlite3.Connection) -> None:
    """Create time_index_candidates table for Sprint 6 time axis outputs."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS time_index_candidates (
            sheet TEXT NOT NULL,
            binding_id TEXT NOT NULL,
            rank INTEGER NOT NULL,
            confidence REAL NOT NULL,
            reasons_top3_json TEXT NOT NULL,
            PRIMARY KEY (sheet, rank),
            FOREIGN KEY (binding_id) REFERENCES bindings(binding_id)
        )
    """)
    conn.commit()


def create_binding_time_annotations_table(conn: sqlite3.Connection) -> None:
    """Create binding_time_annotations table for Sprint 6 time dependence outputs."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS binding_time_annotations (
            binding_id TEXT PRIMARY KEY,
            time_index_binding_id TEXT NOT NULL,
            is_time_dependent BOOLEAN NOT NULL,
            confidence REAL NOT NULL,
            reasons_top3_json TEXT NOT NULL,
            evidence_flags_json TEXT NOT NULL,
            FOREIGN KEY (binding_id) REFERENCES bindings(binding_id),
            FOREIGN KEY (time_index_binding_id) REFERENCES bindings(binding_id)
        )
    """)
    conn.commit()


def create_bindings_table(conn: sqlite3.Connection) -> None:
    """Create bindings table per schema definition."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bindings (
            binding_id TEXT PRIMARY KEY,
            debug_label TEXT,
            sheet TEXT NOT NULL,
            address_a1 TEXT NOT NULL,
            top_left_a1 TEXT NOT NULL,
            shape_rows INTEGER NOT NULL,
            shape_cols INTEGER NOT NULL,
            binding_type TEXT NOT NULL DEFAULT 'formula',
            cells_structure_hash TEXT NOT NULL,
            label_candidates_json TEXT NOT NULL,
            relationships_json TEXT NOT NULL,
            extraction_source TEXT,
            spatial_candidates_json TEXT
        )
    """)
    conn.commit()


def create_binding_level_edges_table(conn: sqlite3.Connection) -> None:
    """Create binding_level_edges table per schema definition."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS binding_level_edges (
            edge_id INTEGER PRIMARY KEY,
            from_binding_id TEXT NOT NULL,
            to_binding_id TEXT NOT NULL,
            FOREIGN KEY (from_binding_id) REFERENCES bindings(binding_id),
            FOREIGN KEY (to_binding_id) REFERENCES bindings(binding_id)
        )
    """)
    conn.commit()


def create_structure_hashes_table(conn: sqlite3.Connection) -> None:
    """Create structure_hashes table per schema definition."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS structure_hashes (
            hash_type TEXT NOT NULL,
            hash_key TEXT NOT NULL,
            hash_value TEXT NOT NULL,
            PRIMARY KEY (hash_type, hash_key)
        )
    """)
    conn.commit()


def create_cells_table(conn: sqlite3.Connection) -> None:
    """Create cells table per schema definition."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cells (
            cell_id INTEGER PRIMARY KEY,
            binding_id TEXT NOT NULL,
            cell_address_a1 TEXT NOT NULL,
            formula_a1 TEXT NOT NULL,
            formula_r1c1 TEXT NOT NULL,
            ast_structural TEXT,
            dtype TEXT NOT NULL,
            value_snapshot TEXT,
            evaluated_value TEXT,
            format_tokens_json TEXT NOT NULL,
            udf_calls_json TEXT NOT NULL,
            protection_locked BOOLEAN NOT NULL,
            protection_hidden_formula BOOLEAN NOT NULL,
            spilled_from TEXT,
            resolution_json TEXT NOT NULL,
            ref_kinds_json TEXT NOT NULL,
            broken_refs_json TEXT,
            extras_json TEXT NOT NULL,
            is_orphan BOOLEAN NOT NULL DEFAULT FALSE,
            FOREIGN KEY (binding_id) REFERENCES bindings(binding_id)
        )
    """)
    conn.commit()


def create_cell_level_edges_table(conn: sqlite3.Connection) -> None:
    """Create cell_level_edges table per schema definition."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cell_level_edges (
            edge_id INTEGER PRIMARY KEY,
            from_cell TEXT NOT NULL,
            to_cell TEXT,
            to_external TEXT,
            UNIQUE(from_cell, to_cell, to_external)
        )
    """)
    conn.commit()


def create_range_level_edges_table(conn: sqlite3.Connection) -> None:
    """Create range_level_edges table per schema definition."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS range_level_edges (
            edge_id INTEGER PRIMARY KEY,
            from_cell TEXT NOT NULL,
            to_range TEXT NOT NULL,
            to_external TEXT,
            cell_count INTEGER NOT NULL,
            UNIQUE(from_cell, to_range, to_external)
        )
    """)
    conn.commit()


def create_levels_table(conn: sqlite3.Connection) -> None:
    """Create levels table per schema definition."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS levels (
            level INTEGER NOT NULL,
            binding_id TEXT NOT NULL,
            PRIMARY KEY (level, binding_id),
            FOREIGN KEY (binding_id) REFERENCES bindings(binding_id)
        )
    """)
    conn.commit()


def create_cycles_table(conn: sqlite3.Connection) -> None:
    """Create cycles table per schema definition."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cycles (
            cycle_id INTEGER NOT NULL,
            ord INTEGER NOT NULL,
            binding_id TEXT NOT NULL,
            PRIMARY KEY (cycle_id, ord),
            FOREIGN KEY (binding_id) REFERENCES bindings(binding_id)
        )
    """)
    conn.commit()


def create_consistency_report_table(conn: sqlite3.Connection) -> None:
    """Create consistency_report table per schema definition."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS consistency_report (
            report_id INTEGER PRIMARY KEY,
            report_type TEXT NOT NULL,
            code TEXT NOT NULL,
            where_location TEXT,
            binding_id TEXT,
            details TEXT,
            drivers_json TEXT,
            FOREIGN KEY (binding_id) REFERENCES bindings(binding_id)
        )
    """)
    conn.commit()


def create_udfs_table(conn: sqlite3.Connection) -> None:
    """Create udfs table per Story 10 requirements (IR Spec §10)."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS udfs (
            udf_id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            module TEXT NOT NULL,
            param_count INTEGER NOT NULL,
            param_names_json TEXT NOT NULL,
            declared_volatile BOOLEAN NOT NULL,
            source_text TEXT NOT NULL,
            source_hash TEXT NOT NULL,
            UNIQUE(name, module)
        )
    """)
    conn.commit()


def create_sheet_topology_table(conn: sqlite3.Connection) -> None:
    """Create sheet_topology table per Story 9 requirements."""
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sheet_topology (
            sheet_name TEXT PRIMARY KEY,
            topology_json TEXT NOT NULL,
            bbox_min_row INTEGER NOT NULL,
            bbox_max_row INTEGER NOT NULL,
            bbox_min_col INTEGER NOT NULL,
            bbox_max_col INTEGER NOT NULL
        )
    """)
    conn.commit()


def insert_udfs(conn: sqlite3.Connection, udfs: list[dict[str, Any]]) -> None:
    """
    Insert UDFs into udfs table.

    Args:
        conn: SQLite connection
        udfs: List of UDF dictionaries from UDFMetadata.to_dict()

    Determinism: UDFs sorted by (name, module) before insertion.
    """
    if not udfs:
        return

    # Sort UDFs by name, module for determinism
    sorted_udfs = sorted(udfs, key=lambda u: (u["name"], u["module"]))

    cursor = conn.cursor()
    for udf in sorted_udfs:
        cursor.execute(
            """
            INSERT INTO udfs (
                name, module, param_count, param_names_json,
                declared_volatile, source_text, source_hash
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
            (
                udf["name"],
                udf["module"],
                udf["param_count"],
                udf["param_names_json"],
                udf["declared_volatile"],
                udf["source_text"],
                udf["source_hash"],
            ),
        )

    conn.commit()


def create_indexes(conn: sqlite3.Connection) -> None:
    """Create indexes for query optimization per schema definition."""
    cursor = conn.cursor()

    # Bindings indexes
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_bindings_sheet ON bindings(sheet)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_bindings_top_left ON bindings(top_left_a1)")

    # Cells indexes
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_cells_binding_id ON cells(binding_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_cells_address ON cells(cell_address_a1)")

    # Cell edges indexes
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_cell_edges_from ON cell_level_edges(from_cell)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_cell_edges_to ON cell_level_edges(to_cell)")

    # Range edges indexes
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_range_edges_from ON range_level_edges(from_cell)"
    )
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_range_edges_to ON range_level_edges(to_range)")

    # Binding edges indexes
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_binding_edges_from ON binding_level_edges(from_binding_id)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_binding_edges_to ON binding_level_edges(to_binding_id)"
    )

    # Levels indexes
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_levels_binding_id ON levels(binding_id)")

    # Cycles indexes
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_cycles_binding_id ON cycles(binding_id)")

    # Consistency report indexes
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_consistency_binding_id ON consistency_report(binding_id)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_consistency_type ON consistency_report(report_type)"
    )

    # UDFs indexes
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_udfs_name ON udfs(name)")

    # Sprint 6 time axis indexes
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_time_index_candidates_sheet ON time_index_candidates(sheet)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_time_index_candidates_binding "
        "ON time_index_candidates(binding_id)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_binding_time_annotations_time_index "
        "ON binding_time_annotations(time_index_binding_id)"
    )

    conn.commit()


def create_schema_tables(conn: sqlite3.Connection) -> None:
    """
    Create all schema tables per design doc.

    Creates meta, user_roots, bindings, binding_level_edges, structure_hashes,
    cells, cell_level_edges, levels, cycles, consistency_report, udfs,
    sheet_topology, time_index_candidates, and binding_time_annotations tables.
    Also creates indexes.

    Args:
        conn: SQLite connection.
    """
    create_meta_table(conn)
    create_user_roots_table(conn)
    create_bindings_table(conn)
    create_binding_level_edges_table(conn)
    create_structure_hashes_table(conn)
    create_cells_table(conn)
    create_cell_level_edges_table(conn)
    create_range_level_edges_table(conn)
    create_levels_table(conn)
    create_cycles_table(conn)
    create_consistency_report_table(conn)
    create_udfs_table(conn)
    create_sheet_topology_table(conn)
    create_time_index_candidates_table(conn)
    create_binding_time_annotations_table(conn)
    create_indexes(conn)


def insert_meta_row(
    conn: sqlite3.Connection,
    workbook_guid: str,
    file_path: str,
    calculation_mode: str = "unknown",
) -> None:
    """
    Insert meta row with deterministic values per ADR-010.

    Args:
        conn: SQLite connection.
        workbook_guid: SHA-256 hex64 workbook GUID.
        file_path: Workbook file path (basename only per ADR-010).
        calculation_mode: "auto" | "manual" | "unknown" (default: "unknown").
    """
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO meta (
            ir_version,
            workbook_guid,
            file_path,
            excel_build_info,
            calculation_mode,
            imported_at,
            created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            IR_VERSION,
            workbook_guid,
            file_path,
            None,  # excel_build_info = NULL (deterministic)
            calculation_mode,
            FIXED_TIMESTAMP,  # imported_at = fixed constant
            FIXED_TIMESTAMP,  # created_at = fixed constant
        ),
    )
    conn.commit()


def insert_user_root(
    conn: sqlite3.Connection,
    sheet: str,
    range_a1: str,
    label_hint: str | None = None,
) -> None:
    """
    Insert single user root per ADR-009.

    Args:
        conn: SQLite connection.
        sheet: Sheet name.
        range_a1: A1 range notation.
        label_hint: Optional label hint (metadata only).
    """
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO user_roots (root_id, sheet, range_a1, label_hint)
        VALUES (?, ?, ?, ?)
        """,
        (1, sheet, range_a1, label_hint),  # root_id = 1 (single root)
    )
    conn.commit()


def insert_time_index_candidates(
    conn: sqlite3.Connection,
    candidates: list[dict[str, Any]],
) -> None:
    """
    Insert time index candidates with deterministic ordering.
    Clears existing candidates for the affected sheets before insertion.

    Args:
        conn: SQLite connection.
        candidates: List of dicts with sheet, binding_id, rank, confidence, reasons_top3_json.
    """
    if not candidates:
        return

    cursor = conn.cursor()
    use_sheet_name = _time_index_candidates_has_sheet_column(conn)

    # Identify affected sheets to clear existing entries
    sheets = {c["sheet"] for c in candidates}
    for sheet in sorted(sheets):
        if use_sheet_name:
            cursor.execute("DELETE FROM time_index_candidates WHERE sheet = ?", (sheet,))
        else:
            # Resolve sheet_id
            cursor.execute("SELECT sheet_id FROM sheets WHERE sheet_name = ?", (sheet,))
            row = cursor.fetchone()
            if row:
                cursor.execute("DELETE FROM time_index_candidates WHERE sheet_id = ?", (row[0],))

    sorted_candidates = sorted(candidates, key=lambda c: (c["sheet"], c["rank"]))
    for candidate in sorted_candidates:
        reasons_json = _canonicalize_json_not_null(
            candidate["reasons_top3_json"], "reasons_top3_json"
        )
        if use_sheet_name:
            cursor.execute(
                """
                INSERT INTO time_index_candidates (
                    sheet, binding_id, rank, confidence, reasons_top3_json
                ) VALUES (?, ?, ?, ?, ?)
                """,
                (
                    candidate["sheet"],
                    candidate["binding_id"],
                    candidate["rank"],
                    candidate["confidence"],
                    reasons_json,
                ),
            )
        else:
            cursor.execute(
                "SELECT sheet_id FROM sheets WHERE sheet_name = ?",
                (candidate["sheet"],),
            )
            row = cursor.fetchone()
            if row is None:
                raise ValueError(f"Unknown sheet in time_index_candidates: {candidate['sheet']}")
            sheet_id = row[0]
            cursor.execute(
                """
                INSERT INTO time_index_candidates (
                    sheet_id, binding_id, rank, confidence, reasons_top3_json
                ) VALUES (?, ?, ?, ?, ?)
                """,
                (
                    sheet_id,
                    candidate["binding_id"],
                    candidate["rank"],
                    candidate["confidence"],
                    reasons_json,
                ),
            )
    conn.commit()


def _time_index_candidates_has_sheet_column(conn: sqlite3.Connection) -> bool:
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(time_index_candidates)")
    return any(row[1] == "sheet" for row in cursor.fetchall())


def insert_binding_time_annotations(
    conn: sqlite3.Connection,
    annotations: list[dict[str, Any]],
) -> None:
    """
    Insert binding time annotations with deterministic ordering.

    Args:
        conn: SQLite connection.
        annotations: List of dicts with binding_id, time_index_binding_id, is_time_dependent,
            confidence, reasons_top3_json, evidence_flags_json.
    """
    if not annotations:
        return

    sorted_annotations = sorted(annotations, key=lambda a: a["binding_id"])
    cursor = conn.cursor()
    for annotation in sorted_annotations:
        cursor.execute(
            """
            INSERT OR REPLACE INTO binding_time_annotations (
                binding_id, time_index_binding_id, is_time_dependent, confidence,
                reasons_top3_json, evidence_flags_json
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                annotation["binding_id"],
                annotation["time_index_binding_id"],
                annotation["is_time_dependent"],
                annotation["confidence"],
                _canonicalize_json_not_null(annotation["reasons_top3_json"], "reasons_top3_json"),
                _canonicalize_json_not_null(
                    annotation["evidence_flags_json"], "evidence_flags_json"
                ),
            ),
        )
    conn.commit()


def fetch_time_index_candidates(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    """
    Fetch time index candidates in deterministic order (sheet, rank).

    Args:
        conn: SQLite connection.

    Returns:
        List of dicts with sheet, binding_id, rank, confidence, reasons_top3_json.
    """
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT sheet, binding_id, rank, confidence, reasons_top3_json
        FROM time_index_candidates
        ORDER BY sheet, rank
        """
    )
    return [
        {
            "sheet": row[0],
            "binding_id": row[1],
            "rank": row[2],
            "confidence": row[3],
            "reasons_top3_json": row[4],
        }
        for row in cursor.fetchall()
    ]


def fetch_binding_time_annotations(
    conn: sqlite3.Connection,
) -> list[dict[str, Any]]:
    """
    Fetch binding time annotations in deterministic order (binding_id).

    Args:
        conn: SQLite connection.

    Returns:
        List of dicts with binding_id, time_index_binding_id, is_time_dependent,
        confidence, reasons_top3_json, evidence_flags_json.
    """
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT binding_id, time_index_binding_id, is_time_dependent, confidence,
               reasons_top3_json, evidence_flags_json
        FROM binding_time_annotations
        ORDER BY binding_id
        """
    )
    return [
        {
            "binding_id": row[0],
            "time_index_binding_id": row[1],
            "is_time_dependent": bool(row[2]),
            "confidence": row[3],
            "reasons_top3_json": row[4],
            "evidence_flags_json": row[5],
        }
        for row in cursor.fetchall()
    ]


def create_database(
    workbook_path: Path,
    roots_data: dict[str, Any],
    output_path: Path,
) -> None:
    """
    Create IR database with meta and user_roots tables.

    Args:
        workbook_path: Path to Excel workbook.
        roots_data: Roots JSON data (validated).
        output_path: Output database path.
    """
    # Compute deterministic workbook GUID
    workbook_guid = compute_workbook_guid(workbook_path)

    # Create database with deterministic settings
    conn = create_deterministic_connection(output_path)

    try:
        # Create schema tables
        create_schema_tables(conn)

        # Insert meta row
        # Per ADR-010: file_path = basename only
        file_path = workbook_path.name
        insert_meta_row(conn, workbook_guid, file_path)

        # Insert user root
        user_root = roots_data["user_root"]
        insert_user_root(
            conn,
            user_root["sheet"],
            user_root["range"],
            user_root.get("label_hint"),
        )

    finally:
        conn.close()


def insert_cells(
    conn: sqlite3.Connection,
    cells: list[dict[str, Any]],
) -> None:
    """
    Insert cells into cells table.

    Args:
        conn: SQLite connection.
        cells: List of cell dictionaries from cell_record_to_dict().
    """
    cursor = conn.cursor()

    # Sort cells by address for deterministic insertion order
    sorted_cells = sorted(cells, key=lambda c: c["cell_address_a1"])

    for cell in sorted_cells:
        cursor.execute(
            """
            INSERT INTO cells (
                binding_id, cell_address_a1, formula_a1, formula_r1c1,
                ast_structural, dtype, value_snapshot, evaluated_value, format_tokens_json,
                udf_calls_json, protection_locked, protection_hidden_formula,
                spilled_from, resolution_json, ref_kinds_json,                 broken_refs_json, extras_json, is_orphan
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                cell["binding_id"],
                cell["cell_address_a1"],
                cell["formula_a1"],
                cell["formula_r1c1"],
                cell["ast_structural"],
                cell["dtype"],
                cell["value_snapshot"],
                cell.get("evaluated_value"),
                cell["format_tokens_json"],
                cell["udf_calls_json"],
                cell["protection_locked"],
                cell["protection_hidden_formula"],
                cell["spilled_from"],
                cell["resolution_json"],
                cell["ref_kinds_json"],
                cell.get("broken_refs_json"),
                cell["extras_json"],
                cell.get("is_orphan", False),
            ),
        )

    conn.commit()


def insert_cell_level_edges(
    conn: sqlite3.Connection,
    edges: list[tuple[str, str | None, str | None]],
) -> None:
    """
    Insert cell-level edges into cell_level_edges table.

    Args:
        conn: SQLite connection.
        edges: List of (from_cell, to_cell, to_external) tuples.
              Exactly one of to_cell or to_external must be non-None.
    """
    cursor = conn.cursor()

    # Sort edges for deterministic insertion order
    sorted_edges = sorted(edges, key=lambda e: (e[0], e[1] or "", e[2] or ""))

    # Use executemany for batch insert (much faster than individual inserts)
    cursor.executemany(
        """
        INSERT OR IGNORE INTO cell_level_edges (from_cell, to_cell, to_external)
        VALUES (?, ?, ?)
        """,
        sorted_edges,
    )

    conn.commit()


def insert_range_level_edges(
    conn: sqlite3.Connection,
    edges: list[tuple[str, str, str | None, int]],
) -> None:
    """
    Insert range-level edges into range_level_edges table.

    Args:
        conn: SQLite connection.
        edges: List of (from_cell, to_range, to_external, cell_count) tuples.
    """
    cursor = conn.cursor()
    sorted_edges = sorted(edges, key=lambda e: (e[0], e[1], e[2] or ""))

    # Use executemany for batch insert (much faster than individual inserts)
    cursor.executemany(
        """
        INSERT OR IGNORE INTO range_level_edges 
        (from_cell, to_range, to_external, cell_count)
        VALUES (?, ?, ?, ?)
    """,
        sorted_edges,
    )
    conn.commit()


def insert_binding_level_edges(
    conn: sqlite3.Connection,
    edges: list[tuple[str, str]],
) -> None:
    """
    Insert binding-level edges into binding_level_edges table.

    Args:
        conn: SQLite connection.
        edges: List of (from_binding_id, to_binding_id) tuples.
    """
    cursor = conn.cursor()

    # Sort edges for deterministic insertion order
    sorted_edges = sorted(edges, key=lambda e: (e[0], e[1]))

    # Use executemany for batch insert (much faster than individual inserts)
    cursor.executemany(
        """
        INSERT INTO binding_level_edges (from_binding_id, to_binding_id)
        VALUES (?, ?)
        """,
        sorted_edges,
    )

    conn.commit()


def insert_bindings(
    conn: sqlite3.Connection,
    bindings: list[dict[str, Any]],
) -> None:
    """
    Insert bindings into bindings table.

    Args:
        conn: SQLite connection.
        bindings: List of binding dictionaries with all required fields.
    """
    cursor = conn.cursor()

    # Sort bindings by ID for deterministic insertion order
    sorted_bindings = sorted(bindings, key=lambda b: b["binding_id"])

    for binding in sorted_bindings:
        cursor.execute(
            """
            INSERT INTO bindings (
                binding_id, debug_label, sheet, address_a1, top_left_a1,
                shape_rows, shape_cols, binding_type, cells_structure_hash,
                label_candidates_json, relationships_json,
                extraction_source, spatial_candidates_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                binding["binding_id"],
                binding.get("debug_label"),
                binding["sheet"],
                binding["address_a1"],
                binding["top_left_a1"],
                binding["shape_rows"],
                binding["shape_cols"],
                binding.get("binding_type", "formula"),
                binding["cells_structure_hash"],
                binding.get("label_candidates_json", "{}"),
                binding.get("relationships_json", "{}"),
                binding.get("extraction_source"),
                binding.get("spatial_candidates_json", "{}"),
            ),
        )

    conn.commit()


def insert_structure_hashes(
    conn: sqlite3.Connection,
    hashes: list[dict[str, str]],
) -> None:
    """
    Insert structure hashes into structure_hashes table.

    Args:
        conn: SQLite connection.
        hashes: List of hash dictionaries with hash_type, hash_key, hash_value.
    """
    cursor = conn.cursor()

    # Sort hashes for deterministic insertion order
    sorted_hashes = sorted(hashes, key=lambda h: (h["hash_type"], h["hash_key"]))

    for hash_entry in sorted_hashes:
        cursor.execute(
            """
            INSERT INTO structure_hashes (hash_type, hash_key, hash_value)
            VALUES (?, ?, ?)
            """,
            (
                hash_entry["hash_type"],
                hash_entry["hash_key"],
                hash_entry["hash_value"],
            ),
        )

    conn.commit()


def insert_levels(
    conn: sqlite3.Connection,
    levels: list[tuple[int, str]],
) -> None:
    """
    Insert topological levels into levels table.

    Args:
        conn: SQLite connection.
        levels: List of (level, binding_id) tuples sorted by level, then binding_id.
    """
    cursor = conn.cursor()

    # Sort levels for deterministic insertion order (should already be sorted)
    sorted_levels = sorted(levels, key=lambda lvl: (lvl[0], lvl[1]))

    for level, binding_id in sorted_levels:
        cursor.execute(
            """
            INSERT INTO levels (level, binding_id)
            VALUES (?, ?)
            """,
            (level, binding_id),
        )

    conn.commit()


def insert_cycles(
    conn: sqlite3.Connection,
    cycles: list[tuple[int, int, str]],
) -> None:
    """
    Insert cycles into cycles table.

    Args:
        conn: SQLite connection.
        cycles: List of (cycle_id, ord, binding_id) tuples sorted by cycle_id, then ord.
    """
    cursor = conn.cursor()

    # Sort cycles for deterministic insertion order (should already be sorted)
    sorted_cycles = sorted(cycles, key=lambda c: (c[0], c[1]))

    for cycle_id, ord_num, binding_id in sorted_cycles:
        cursor.execute(
            """
            INSERT INTO cycles (cycle_id, ord, binding_id)
            VALUES (?, ?, ?)
            """,
            (cycle_id, ord_num, binding_id),
        )

    conn.commit()


def validate_schema(db_path: Path) -> dict[str, Any]:
    """
    Validate database schema against design spec.

    Comprehensive validation including:
    - All required tables exist (meta, user_roots, bindings, binding_level_edges,
      structure_hashes, cells, cell_level_edges, range_level_edges, levels, cycles,
      consistency_report, udfs, sheet_topology)
    - Tables have correct columns with correct types
    - PRAGMAs are set correctly
    - Foreign key constraints exist

    Args:
        db_path: Path to database file.

    Returns:
        Validation result dictionary with:
        - valid: bool (True if all checks pass)
        - errors: List[str] (error messages)
        - warnings: List[str] (warning messages)
        - pragma_check: Dict[str, Any] (PRAGMA values)
    """
    errors: list[str] = []
    warnings: list[str] = []
    pragma_values: dict[str, Any] = {}

    # Define expected schema (column_name: (type, not_null))
    # Note: SQLite type affinity is flexible, we check the declared type
    expected_schema = {
        "meta": {
            "ir_version": ("TEXT", True),
            "workbook_guid": ("TEXT", True),
            "file_path": ("TEXT", True),
            "excel_build_info": ("TEXT", False),
            "calculation_mode": ("TEXT", True),
            "imported_at": ("TEXT", True),
            "created_at": ("TEXT", True),
        },
        "user_roots": {
            "root_id": ("INTEGER", False),  # PRIMARY KEY
            "sheet": ("TEXT", True),
            "range_a1": ("TEXT", True),
            "label_hint": ("TEXT", False),
        },
        "bindings": {
            "binding_id": ("TEXT", False),  # PRIMARY KEY
            "debug_label": ("TEXT", False),
            "sheet": ("TEXT", True),
            "address_a1": ("TEXT", True),
            "top_left_a1": ("TEXT", True),
            "shape_rows": ("INTEGER", True),
            "shape_cols": ("INTEGER", True),
            "binding_type": ("TEXT", True),
            "cells_structure_hash": ("TEXT", True),
            "label_candidates_json": ("TEXT", True),
            "relationships_json": ("TEXT", True),
            "extraction_source": ("TEXT", False),
            "spatial_candidates_json": ("TEXT", False),
        },
        "binding_level_edges": {
            "edge_id": ("INTEGER", False),  # PRIMARY KEY
            "from_binding_id": ("TEXT", True),
            "to_binding_id": ("TEXT", True),
        },
        "structure_hashes": {
            "hash_type": ("TEXT", True),
            "hash_key": ("TEXT", True),
            "hash_value": ("TEXT", True),
        },
        "cells": {
            "cell_id": ("INTEGER", False),  # PRIMARY KEY
            "binding_id": ("TEXT", True),
            "cell_address_a1": ("TEXT", True),
            "formula_a1": ("TEXT", True),
            "formula_r1c1": ("TEXT", True),
            "ast_structural": ("TEXT", False),
            "dtype": ("TEXT", True),
            "value_snapshot": ("TEXT", False),
            "evaluated_value": ("TEXT", False),
            "format_tokens_json": ("TEXT", True),
            "udf_calls_json": ("TEXT", True),
            "protection_locked": ("BOOLEAN", True),
            "protection_hidden_formula": ("BOOLEAN", True),
            "spilled_from": ("TEXT", False),
            "resolution_json": ("TEXT", True),
            "ref_kinds_json": ("TEXT", True),
            "broken_refs_json": ("TEXT", False),
            "extras_json": ("TEXT", True),
            "is_orphan": ("BOOLEAN", True),
        },
        "cell_level_edges": {
            "edge_id": ("INTEGER", False),  # PRIMARY KEY
            "from_cell": ("TEXT", True),
            "to_cell": ("TEXT", False),
            "to_external": ("TEXT", False),
        },
        "range_level_edges": {
            "edge_id": ("INTEGER", False),  # PRIMARY KEY
            "from_cell": ("TEXT", True),
            "to_range": ("TEXT", True),
            "to_external": ("TEXT", False),
            "cell_count": ("INTEGER", True),
        },
        "levels": {
            "level": ("INTEGER", True),
            "binding_id": ("TEXT", True),
        },
        "cycles": {
            "cycle_id": ("INTEGER", True),
            "ord": ("INTEGER", True),
            "binding_id": ("TEXT", True),
        },
        "consistency_report": {
            "report_id": ("INTEGER", False),  # PRIMARY KEY
            "report_type": ("TEXT", True),
            "code": ("TEXT", True),
            "where_location": ("TEXT", False),
            "binding_id": ("TEXT", False),
            "details": ("TEXT", False),
            "drivers_json": ("TEXT", False),
        },
        "udfs": {
            "udf_id": ("INTEGER", False),  # PRIMARY KEY
            "name": ("TEXT", True),
            "module": ("TEXT", True),
            "param_count": ("INTEGER", True),
            "param_names_json": ("TEXT", True),
            "declared_volatile": ("BOOLEAN", True),
            "source_text": ("TEXT", True),
            "source_hash": ("TEXT", True),
        },
        "sheet_topology": {
            "sheet_name": ("TEXT", False),  # PRIMARY KEY
            "topology_json": ("TEXT", True),
            "bbox_min_row": ("INTEGER", True),
            "bbox_max_row": ("INTEGER", True),
            "bbox_min_col": ("INTEGER", True),
            "bbox_max_col": ("INTEGER", True),
        },
        "time_index_candidates": {
            "sheet": ("TEXT", True),
            "binding_id": ("TEXT", True),
            "rank": ("INTEGER", True),
            "confidence": ("REAL", True),
            "reasons_top3_json": ("TEXT", True),
        },
        "binding_time_annotations": {
            "binding_id": ("TEXT", False),  # PRIMARY KEY
            "time_index_binding_id": ("TEXT", True),
            "is_time_dependent": ("BOOLEAN", True),
            "confidence": ("REAL", True),
            "reasons_top3_json": ("TEXT", True),
            "evidence_flags_json": ("TEXT", True),
        },
    }

    try:
        conn = sqlite3.connect(str(db_path))
        cursor = conn.cursor()

        # Get all tables in database
        cursor.execute("""
            SELECT name FROM sqlite_master
            WHERE type='table' AND name NOT LIKE 'sqlite_%'
            ORDER BY name
        """)
        actual_tables = {row[0] for row in cursor.fetchall()}

        # Check all required tables exist
        required_tables = set(expected_schema.keys())
        missing_tables = required_tables - actual_tables
        for table in sorted(missing_tables):
            errors.append(f"Required table '{table}' not found")

        # Validate each table's schema
        for table_name, expected_cols in expected_schema.items():
            if table_name not in actual_tables:
                continue  # Already reported as missing

            cursor.execute(f"PRAGMA table_info({table_name})")
            # PRAGMA table_info returns: (cid, name, type, notnull, dflt_value, pk)
            actual_cols = {row[1]: (row[2].upper(), bool(row[3])) for row in cursor.fetchall()}

            # Check for missing columns
            missing_cols = set(expected_cols.keys()) - set(actual_cols.keys())
            if missing_cols:
                errors.append(f"Table '{table_name}' missing columns: {sorted(missing_cols)}")

            # Check for extra columns (warning only)
            extra_cols = set(actual_cols.keys()) - set(expected_cols.keys())
            if extra_cols:
                warnings.append(
                    f"Table '{table_name}' has unexpected columns: {sorted(extra_cols)}"
                )

            # Check column types and NOT NULL constraints
            for col_name, (expected_type, expected_not_null) in expected_cols.items():
                if col_name not in actual_cols:
                    continue  # Already reported as missing

                actual_type, actual_not_null = actual_cols[col_name]

                # SQLite type affinity: check if types are compatible
                # TEXT, INTEGER, BOOLEAN are the main types we use
                if actual_type and not _types_compatible(actual_type, expected_type):
                    errors.append(
                        f"Table '{table_name}' column '{col_name}': "
                        f"expected type {expected_type}, got {actual_type}"
                    )

                # Check NOT NULL constraint (only if expected)
                if expected_not_null and not actual_not_null:
                    warnings.append(
                        f"Table '{table_name}' column '{col_name}': expected NOT NULL constraint"
                    )

        # Check PRAGMA settings
        cursor.execute("PRAGMA page_size")
        page_size = cursor.fetchone()[0]
        pragma_values["page_size"] = page_size
        if page_size != 4096:
            warnings.append(f"page_size is {page_size}, expected 4096")

        cursor.execute("PRAGMA journal_mode")
        journal_mode = cursor.fetchone()[0]
        pragma_values["journal_mode"] = journal_mode
        if journal_mode.lower() not in ("delete", "truncate", "persist"):
            warnings.append(
                f"journal_mode is '{journal_mode}', expected 'delete' (or truncate/persist)"
            )

        conn.close()

    except sqlite3.Error as e:
        errors.append(f"SQLite error: {e}")
    except Exception as e:
        errors.append(f"Validation error: {e}")

    return {
        "valid": len(errors) == 0,
        "errors": errors,
        "warnings": warnings,
        "pragma_check": pragma_values,
    }


def _types_compatible(actual_type: str, expected_type: str) -> bool:
    """
    Check if SQLite column types are compatible.

    SQLite has flexible type affinity, so we check for compatibility
    rather than exact match.

    Args:
        actual_type: Actual column type from PRAGMA table_info
        expected_type: Expected column type from schema definition

    Returns:
        True if types are compatible
    """
    # Normalize types
    actual = actual_type.upper().strip()
    expected = expected_type.upper().strip()

    # Exact match
    if actual == expected:
        return True

    # Empty type is compatible with anything (SQLite allows this)
    if not actual:
        return True

    # INTEGER affinity
    if expected == "INTEGER":
        return actual in ("INTEGER", "INT", "TINYINT", "SMALLINT", "MEDIUMINT", "BIGINT")

    # TEXT affinity
    if expected == "TEXT":
        return actual in ("TEXT", "VARCHAR", "CHAR", "CLOB")

    # BOOLEAN is stored as INTEGER in SQLite
    if expected == "BOOLEAN":
        return actual in ("BOOLEAN", "INTEGER", "INT")

    # REAL affinity
    if expected == "REAL":
        return actual in ("REAL", "DOUBLE", "FLOAT")

    # BLOB affinity
    if expected == "BLOB":
        return actual == "BLOB"

    return False
