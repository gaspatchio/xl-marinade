# ABOUTME: Orchestrates heap-first loading and DB-driven deduplication for memory-efficient extraction
# ABOUTME: Implements Pattern A: Stream -> Heap -> Dedupe -> Final with VACUUM INTO canonicalization

"""
Bulk Loader

Orchestrates the memory-efficient extraction pipeline:
1. Stream raw data into heap tables (no PK/indexes)
2. Batch insert using executemany
3. Finalize: dedupe and normalize into final tables using deterministic SQL
4. VACUUM INTO to produce canonical artifact

Design reference: §6.2 Principle 3/4, §6.5, §7.1 of memory_efficient_extraction_architecture.md
"""

import hashlib
import os
import sqlite3
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path

from .formula_canonical import compute_canonical_a1
from .memory_budget import MemoryBudgetConfig, MemoryBudgetController
from .sqlite_pragmas import apply_pragmas, verify_pragmas

# Batch size for executemany (configurable)
DEFAULT_BATCH_SIZE = 10_000
MIN_BATCH_SIZE = 1_000
MAX_BATCH_SIZE = 50_000


@dataclass
class LoadStats:
    """Statistics from bulk loading."""

    raw_cells: int = 0
    raw_formulas: int = 0
    raw_json_blobs: int = 0
    raw_edges_internal: int = 0
    raw_edges_range: int = 0
    raw_edges_external: int = 0

    final_cells: int = 0
    final_formulas: int = 0
    final_json_blobs: int = 0
    final_edges_internal: int = 0
    final_edges_range: int = 0
    final_edges_external: int = 0


class BulkLoader:
    """
    Bulk loader for memory-efficient IR extraction.

    Implements heap-first loading pattern:
    - Write to raw_* heap tables with no constraints
    - Batch insert using executemany
    - Finalize using deterministic SQL (explicit ORDER BY)
    - VACUUM INTO to produce canonical artifact
    """

    def __init__(
        self,
        build_db_path: str,
        batch_size: int = DEFAULT_BATCH_SIZE,
        allow_non_canonical: bool = False,
        memory_budget_config: MemoryBudgetConfig | None = None,
    ):
        """
        Initialize bulk loader.

        Args:
            build_db_path: Path to build database (temporary, will be deleted after VACUUM INTO)
            batch_size: Number of rows per executemany batch
            allow_non_canonical: Allow non-canonical builds (for testing only)
            memory_budget_config: Memory budget configuration (uses defaults if None)
        """
        if not (MIN_BATCH_SIZE <= batch_size <= MAX_BATCH_SIZE):
            raise ValueError(
                f"batch_size must be between {MIN_BATCH_SIZE} and {MAX_BATCH_SIZE}, "
                f"got {batch_size}"
            )

        self.build_db_path = Path(build_db_path)
        self.batch_size = batch_size
        self.allow_non_canonical = allow_non_canonical
        self.conn: sqlite3.Connection | None = None
        self.stats = LoadStats()

        # Initialize memory budget controller
        self.memory_budget = MemoryBudgetController(memory_budget_config)

        # Verify SQLite version
        version = sqlite3.sqlite_version_info
        if version < (3, 27, 0):
            raise RuntimeError(
                f"SQLite 3.27.0+ required for VACUUM INTO support, found {sqlite3.sqlite_version}"
            )

    def __enter__(self):
        """Context manager entry."""
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        if self.conn:
            self.conn.close()
            self.conn = None

    def open(self) -> None:
        """Open build database and apply PRAGMAs."""
        if self.conn:
            raise RuntimeError("Database already open")

        # Remove existing build DB
        if self.build_db_path.exists():
            self.build_db_path.unlink()

        # Open connection
        self.conn = sqlite3.connect(str(self.build_db_path))

        # Apply PRAGMAs
        apply_pragmas(self.conn)

        # Verify PRAGMAs
        verify_pragmas(self.conn)

        # Register canonical-A1 computation as a SQL function (Cycle 17 #312-B)
        # so formula finalization can populate formula_canonical_a1 inline.
        self.conn.create_function(
            "compute_canonical_a1", 1, compute_canonical_a1, deterministic=True
        )

    def create_schema(self, schema_sql: str) -> None:
        """
        Create schema from SQL file.

        Args:
            schema_sql: SQL text containing CREATE TABLE statements
        """
        if not self.conn:
            raise RuntimeError("Database not open")

        # Execute schema in a transaction
        with self.conn:
            self.conn.executescript(schema_sql)

    def create_views(self) -> None:
        """Ensure agent views exist after finalization.

        Views are defined in schema.sql and created during create_schema().
        This method is a no-op placeholder to preserve the pipeline contract.
        """
        if not self.conn:
            raise RuntimeError("Database not open")

    def load_sheets(self, sheets: list[tuple[int, str]]) -> None:
        """
        Load sheet catalog.

        Args:
            sheets: List of (sheet_id, sheet_name) tuples in deterministic order
        """
        if not self.conn:
            raise RuntimeError("Database not open")

        with self.conn:
            self.conn.executemany("INSERT INTO sheets (sheet_id, sheet_name) VALUES (?, ?)", sheets)

    def load_raw_cells(self, cells: Iterator[tuple]) -> None:
        """
        Load raw cells into heap table.

        Args:
            cells: Iterator of (cell_id, sheet_id, row, col, a1, formula_r1c1,
                   formula_a1, value_sha256, format_sha256, data_type,
                   is_array_formula, is_spilled, spilled_from_cell_id) tuples
        """
        if not self.conn:
            raise RuntimeError("Database not open")

        batch = []
        count = 0

        with self.conn:
            for cell in cells:
                batch.append(cell)
                count += 1

                if len(batch) >= self.batch_size:
                    self.conn.executemany(
                        """
                        INSERT INTO raw_cells (
                            cell_id, sheet_id, row, col, a1,
                            formula_r1c1, formula_a1, value_sha256, format_sha256,
                            data_type, is_array_formula, is_spilled, spilled_from_cell_id
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        batch,
                    )
                    batch = []

                    # Check memory budget after each batch
                    self.memory_budget.check(count)

            # Insert remaining
            if batch:
                self.conn.executemany(
                    """
                    INSERT INTO raw_cells (
                        cell_id, sheet_id, row, col, a1,
                        formula_r1c1, formula_a1, value_sha256, format_sha256,
                        data_type, is_array_formula, is_spilled, spilled_from_cell_id
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    batch,
                )

        self.stats.raw_cells = count

    def load_raw_formulas(self, formulas: Iterator[tuple]) -> None:
        """
        Load raw formulas into heap table.

        Args:
            formulas: Iterator of either:
                - (formula_r1c1, formula_a1_example)
                - (formula_r1c1, formula_a1_example, source_sheet_id, source_row, source_col)
            The 2-tuple form is accepted for backward compatibility and stores NULL
            source coordinates.
        """
        if not self.conn:
            raise RuntimeError("Database not open")

        def _normalize_formula_row(
            row: tuple,
        ) -> tuple[str, str | None, int | None, int | None, int | None]:
            if len(row) == 2:
                formula_r1c1, formula_a1_example = row
                return str(formula_r1c1), formula_a1_example, None, None, None
            if len(row) == 5:
                formula_r1c1, formula_a1_example, source_sheet_id, source_row, source_col = row
                return (
                    str(formula_r1c1),
                    formula_a1_example,
                    int(source_sheet_id) if source_sheet_id is not None else None,
                    int(source_row) if source_row is not None else None,
                    int(source_col) if source_col is not None else None,
                )
            raise ValueError(
                "Formula row must have 2 or 5 fields: "
                "(formula_r1c1, formula_a1_example[, source_sheet_id, source_row, source_col])"
            )

        batch = []
        count = 0

        with self.conn:
            for formula in formulas:
                batch.append(_normalize_formula_row(formula))
                count += 1

                if len(batch) >= self.batch_size:
                    self.conn.executemany(
                        """
                        INSERT INTO raw_formulas (
                            formula_r1c1, formula_a1_example, source_sheet_id, source_row, source_col
                        ) VALUES (?, ?, ?, ?, ?)
                        """,
                        batch,
                    )
                    batch = []

            # Insert remaining
            if batch:
                self.conn.executemany(
                    """
                    INSERT INTO raw_formulas (
                        formula_r1c1, formula_a1_example, source_sheet_id, source_row, source_col
                    ) VALUES (?, ?, ?, ?, ?)
                    """,
                    batch,
                )

        self.stats.raw_formulas = count

    def load_raw_json_blobs(self, blobs: Iterator[tuple[str, str]]) -> None:
        """
        Load raw JSON blobs into heap table.

        Args:
            blobs: Iterator of (sha256, json) tuples
        """
        if not self.conn:
            raise RuntimeError("Database not open")

        batch = []
        count = 0

        with self.conn:
            for blob in blobs:
                batch.append(blob)
                count += 1

                if len(batch) >= self.batch_size:
                    self.conn.executemany(
                        "INSERT INTO raw_json_blobs (sha256, json) VALUES (?, ?)", batch
                    )
                    batch = []

            # Insert remaining
            if batch:
                self.conn.executemany(
                    "INSERT INTO raw_json_blobs (sha256, json) VALUES (?, ?)", batch
                )

        self.stats.raw_json_blobs = count

    def load_raw_edges_internal(self, edges: Iterator[tuple[int, int]]) -> None:
        """
        Load raw internal edges into heap table.

        Args:
            edges: Iterator of (from_cell_id, to_cell_id) tuples
        """
        if not self.conn:
            raise RuntimeError("Database not open")

        batch = []
        count = 0

        with self.conn:
            for edge in edges:
                batch.append(edge)
                count += 1

                if len(batch) >= self.batch_size:
                    self.conn.executemany(
                        "INSERT INTO raw_edges_internal (from_cell_id, to_cell_id) VALUES (?, ?)",
                        batch,
                    )
                    batch = []

                    # Check memory budget after each batch
                    self.memory_budget.check(count)

            # Insert remaining
            if batch:
                self.conn.executemany(
                    "INSERT INTO raw_edges_internal (from_cell_id, to_cell_id) VALUES (?, ?)", batch
                )

        self.stats.raw_edges_internal = count

    def load_raw_edges_range(self, edges: Iterator[tuple]) -> None:
        """
        Load raw range edges into heap table.

        Args:
            edges: Iterator of (from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2,
                   to_range_a1, cell_count, provenance) tuples
        """
        if not self.conn:
            raise RuntimeError("Database not open")

        batch = []
        count = 0

        with self.conn:
            for edge in edges:
                batch.append(edge)
                count += 1

                if len(batch) >= self.batch_size:
                    self.conn.executemany(
                        """
                        INSERT INTO raw_edges_range (
                            from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2,
                            to_range_a1, cell_count, provenance
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        batch,
                    )
                    batch = []

                    # Check memory budget after each batch
                    self.memory_budget.check(count)

            # Insert remaining
            if batch:
                self.conn.executemany(
                    """
                    INSERT INTO raw_edges_range (
                        from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2,
                        to_range_a1, cell_count, provenance
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    batch,
                )

        self.stats.raw_edges_range = count

    def load_raw_edges_external(self, edges: Iterator[tuple[int, str]]) -> None:
        """
        Load raw external edges into heap table.

        Args:
            edges: Iterator of (from_cell_id, external_ref) tuples
        """
        if not self.conn:
            raise RuntimeError("Database not open")

        batch = []
        count = 0

        with self.conn:
            for edge in edges:
                batch.append(edge)
                count += 1

                if len(batch) >= self.batch_size:
                    self.conn.executemany(
                        "INSERT INTO raw_edges_external (from_cell_id, external_ref) VALUES (?, ?)",
                        batch,
                    )
                    batch = []

            # Insert remaining
            if batch:
                self.conn.executemany(
                    "INSERT INTO raw_edges_external (from_cell_id, external_ref) VALUES (?, ?)",
                    batch,
                )

        self.stats.raw_edges_external = count

    def finalize(self) -> None:
        """
        Finalize tables using normative SQL.

        Executes the deterministic SQL from Design §7.1:
        - Dedupe formulas and JSON blobs
        - Dedupe cells with validation
        - Dedupe edges
        - All with explicit ORDER BY for determinism
        """
        if not self.conn:
            raise RuntimeError("Database not open")

        with self.conn:
            # Validate raw_cells consistency (per design §7.1)
            # For each (sheet_id, row, col), all non-null values must be identical
            cursor = self.conn.execute("""
                WITH cell_groups AS (
                    SELECT 
                        sheet_id, row, col,
                        COUNT(DISTINCT CASE WHEN cell_id IS NOT NULL THEN cell_id END) as cell_id_count,
                        COUNT(DISTINCT CASE WHEN a1 IS NOT NULL THEN a1 END) as a1_count,
                        COUNT(DISTINCT CASE WHEN formula_r1c1 IS NOT NULL THEN formula_r1c1 END) as formula_count,
                        COUNT(DISTINCT CASE WHEN value_sha256 IS NOT NULL THEN value_sha256 END) as value_count,
                        COUNT(DISTINCT CASE WHEN format_sha256 IS NOT NULL THEN format_sha256 END) as format_count,
                        COUNT(DISTINCT CASE WHEN data_type IS NOT NULL THEN data_type END) as type_count,
                        COUNT(DISTINCT CASE WHEN is_array_formula IS NOT NULL THEN is_array_formula END) as array_count,
                        COUNT(DISTINCT CASE WHEN is_spilled IS NOT NULL THEN is_spilled END) as spilled_count,
                        COUNT(DISTINCT CASE WHEN spilled_from_cell_id IS NOT NULL THEN spilled_from_cell_id END) as spilled_from_count
                    FROM raw_cells
                    GROUP BY sheet_id, row, col
                )
                SELECT sheet_id, row, col
                FROM cell_groups
                WHERE cell_id_count > 1 OR a1_count > 1 OR formula_count > 1
                   OR value_count > 1 OR format_count > 1
                   OR type_count > 1 OR array_count > 1 OR spilled_count > 1
                   OR spilled_from_count > 1
                LIMIT 1
            """)

            conflict = cursor.fetchone()
            if conflict:
                sheet_id, row, col = conflict[:3]
                raise RuntimeError(
                    f"Conflicting cell data at (sheet_id={sheet_id}, row={row}, col={col}). "
                    f"Multiple distinct non-null values found for the same cell coordinates."
                )

            # Finalize formulas (deterministic first-occurrence A1 example).
            # Prefer the earliest observed source row for each normalized R1C1
            # formula; fallback to lexical A1 ordering when source coordinates are
            # unavailable (legacy 2-column raw_formulas records).
            self.conn.execute("""
                WITH ranked AS (
                    SELECT
                        formula_r1c1,
                        formula_a1_example,
                        ROW_NUMBER() OVER (
                            PARTITION BY formula_r1c1
                            ORDER BY
                                CASE WHEN source_row IS NULL THEN 1 ELSE 0 END,
                                source_row,
                                CASE WHEN source_col IS NULL THEN 1 ELSE 0 END,
                                source_col,
                                CASE WHEN source_sheet_id IS NULL THEN 1 ELSE 0 END,
                                source_sheet_id,
                                COALESCE(formula_a1_example, '')
                        ) AS rn
                    FROM raw_formulas
                )
                INSERT INTO formulas (formula_r1c1, formula_a1_example, formula_canonical_a1)
                SELECT
                    formula_r1c1,
                    formula_a1_example,
                    compute_canonical_a1(formula_a1_example)
                FROM ranked
                WHERE rn = 1
                ORDER BY formula_r1c1
            """)

            # Finalize JSON blobs (normative SQL from design §7.1)
            self.conn.execute("""
                INSERT INTO json_blobs (sha256, json)
                SELECT sha256, json
                FROM raw_json_blobs
                GROUP BY sha256
                ORDER BY sha256
            """)

            # Finalize cells (normative SQL from design §7.1)
            self.conn.execute("""
                WITH cells_dedup AS (
                    SELECT
                        sheet_id,
                        row,
                        col,
                        MAX(cell_id) AS cell_id,
                        MAX(a1) AS a1,
                        MAX(formula_r1c1) AS formula_r1c1,
                        MAX(formula_a1) AS formula_a1,
                        MAX(value_sha256) AS value_sha256,
                        MAX(format_sha256) AS format_sha256,
                        MAX(data_type) AS data_type,
                        MAX(is_array_formula) AS is_array_formula,
                        MAX(is_spilled) AS is_spilled,
                        MAX(spilled_from_cell_id) AS spilled_from_cell_id
                    FROM raw_cells
                    GROUP BY sheet_id, row, col
                )
                INSERT INTO cells (
                    cell_id, sheet_id, row, col, a1,
                    formula_id, formula_a1, format_blob_id, value_blob_id,
                    data_type, is_array_formula, is_spilled, spilled_from_cell_id
                )
                SELECT
                    cd.cell_id, cd.sheet_id, cd.row, cd.col, cd.a1,
                    f.formula_id, cd.formula_a1, jf.blob_id, jv.blob_id,
                    cd.data_type, cd.is_array_formula, cd.is_spilled, cd.spilled_from_cell_id
                FROM cells_dedup cd
                LEFT JOIN formulas f ON cd.formula_r1c1 = f.formula_r1c1
                LEFT JOIN json_blobs jv ON cd.value_sha256 = jv.sha256
                LEFT JOIN json_blobs jf ON cd.format_sha256 = jf.sha256
                ORDER BY cd.sheet_id, cd.row, cd.col
            """)

            # Finalize internal edges (normative SQL from design §7.1)
            self.conn.execute("""
                INSERT INTO cell_edges_internal (from_cell_id, to_cell_id)
                SELECT DISTINCT from_cell_id, to_cell_id
                FROM raw_edges_internal
                ORDER BY from_cell_id, to_cell_id
            """)

            # Finalize range edges (normative SQL from design §7.1)
            self.conn.execute("""
                INSERT INTO range_edges (
                    from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2, to_range_a1, cell_count, provenance
                )
                SELECT
                    from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2, to_range_a1, cell_count,
                    MAX(provenance)  -- a rect resolved from cache by any source keeps that provenance
                FROM raw_edges_range
                GROUP BY from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2
                ORDER BY from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2
            """)

            # Finalize external edges (normative SQL from design §7.1)
            self.conn.execute("""
                INSERT INTO cell_edges_external (from_cell_id, external_ref)
                SELECT DISTINCT from_cell_id, external_ref
                FROM raw_edges_external
                ORDER BY from_cell_id, external_ref
            """)

        # Update stats
        self.stats.final_cells = self.conn.execute("SELECT COUNT(*) FROM cells").fetchone()[0]
        self.stats.final_formulas = self.conn.execute("SELECT COUNT(*) FROM formulas").fetchone()[0]
        self.stats.final_json_blobs = self.conn.execute(
            "SELECT COUNT(*) FROM json_blobs"
        ).fetchone()[0]
        self.stats.final_edges_internal = self.conn.execute(
            "SELECT COUNT(*) FROM cell_edges_internal"
        ).fetchone()[0]
        self.stats.final_edges_range = self.conn.execute(
            "SELECT COUNT(*) FROM range_edges"
        ).fetchone()[0]
        self.stats.final_edges_external = self.conn.execute(
            "SELECT COUNT(*) FROM cell_edges_external"
        ).fetchone()[0]

    def drop_raw_tables(self) -> None:
        """Drop raw staging tables to reduce final database size."""
        if not self.conn:
            raise RuntimeError("Database not open")

        raw_tables = (
            "raw_edges_internal",
            "raw_edges_range",
            "raw_edges_external",
            "raw_formulas",
            "raw_json_blobs",
            "raw_cells",
        )
        with self.conn:
            for table in raw_tables:
                self.conn.execute(f"DROP TABLE IF EXISTS {table}")

    def vacuum_into(self, output_path: str) -> None:
        """
        Produce canonical artifact using VACUUM INTO.

        Args:
            output_path: Path to final IR database

        This is the ONLY way to produce the canonical artifact (Design §6.5).
        No writes are allowed after this step.
        """
        if not self.conn:
            raise RuntimeError("Database not open")

        output_path = Path(output_path)

        # VACUUM INTO a sibling temp file, then replace the target atomically.
        # Deleting the target first (as this used to) meant a VACUUM that
        # failed part-way — a full disk is the realistic case, and it is
        # exactly when a user re-runs an extraction — destroyed the previous
        # good database and produced no new one. os.replace is atomic on
        # POSIX and on Windows for same-volume paths, so the target is either
        # the old database or the complete new one, never nothing.
        # VACUUM INTO refuses to write an existing file, so clear any temp
        # left behind by an earlier interrupted run.
        tmp_path = output_path.with_name(output_path.name + ".tmp-vacuum")
        if tmp_path.exists():
            tmp_path.unlink()

        # VACUUM INTO (requires SQLite 3.27+)
        # Use parameterized query to prevent SQL injection
        try:
            with self.conn:
                self.conn.execute("VACUUM INTO ?", (str(tmp_path),))
        except BaseException:
            tmp_path.unlink(missing_ok=True)
            raise

        # Close connection (no writes after VACUUM INTO)
        self.conn.close()
        self.conn = None

        # os.replace consumes tmp_path on success, so the unlink is a no-op
        # then. It matters on failure — a Windows PermissionError when another
        # process still holds the output open would otherwise strand a
        # full-size .tmp-vacuum next to the database the caller still has.
        try:
            os.replace(tmp_path, output_path)
        finally:
            tmp_path.unlink(missing_ok=True)

    def get_peak_rss_mb(self) -> float:
        """
        Get peak RSS observed during bulk loading.

        Returns:
            Peak RSS in megabytes from memory budget controller
        """
        return self.memory_budget.get_peak_rss_mb()

    def get_memory_telemetry(self) -> dict:
        """
        Get memory telemetry from memory budget controller.

        Returns:
            Dictionary with peak_rss_mb, check_count, and config
        """
        return self.memory_budget.get_telemetry()

    def compute_logical_checksum(self) -> str:
        r"""
        Compute logical checksum of database content.

        Returns:
            SHA256 hex digest of serialized database content

        Algorithm (per design §7.1):
        - Query all tables in primary-key order
        - Serialize each row as ||-joined column values
        - Rows separated by newline
        - SQL NULL rendered as \N
        - Concatenate all tables and hash
        """
        if not self.conn:
            raise RuntimeError("Database not open")

        hasher = hashlib.sha256()

        # Tables with their primary key columns (in order)
        tables_with_pks = [
            ("sheets", ["sheet_id"]),
            ("formulas", ["formula_id"]),
            ("json_blobs", ["blob_id"]),
            ("cells", ["cell_id"]),
            ("cell_edges_internal", ["from_cell_id", "to_cell_id"]),
            ("range_edges", ["from_cell_id", "to_sheet_id", "to_r1", "to_c1", "to_r2", "to_c2"]),
            ("cell_edges_external", ["from_cell_id", "external_ref"]),
        ]

        for table, pk_columns in tables_with_pks:
            # Query all rows in primary key order
            order_by = ", ".join(pk_columns)
            cursor = self.conn.execute(f"SELECT * FROM {table} ORDER BY {order_by}")

            for row in cursor:
                # Serialize row
                parts = []
                for value in row:
                    if value is None:
                        parts.append(r"\N")
                    else:
                        parts.append(str(value))

                row_str = "|".join(parts) + "\n"
                hasher.update(row_str.encode("utf-8"))

        return hasher.hexdigest()


def compute_file_checksum(db_path: str) -> str:
    """
    Compute SHA256 checksum of database file.

    Args:
        db_path: Path to database file

    Returns:
        SHA256 hex digest
    """
    hasher = hashlib.sha256()

    with open(db_path, "rb") as f:
        while chunk := f.read(8192):
            hasher.update(chunk)

    return hasher.hexdigest()
