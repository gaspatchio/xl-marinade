# ABOUTME: Unit tests for table candidate detection (Sprint 10 Story 7)
# ABOUTME: Tests deterministic grouping of 1D bindings into table candidates

"""
Unit tests for xl_marinade.core.new_arch.table_candidates

Tests cover:
- Deterministic candidate ID generation
- Column vector grouping by (sheet_id, r1, r2)
- Contiguous segment detection
- Minimum threshold enforcement
- Stable membership ordering
"""

import json
import sqlite3
import tempfile
from pathlib import Path

import pytest

from xl_marinade.core.new_arch.table_candidates import (
    _compute_candidate_id,
    _range_to_a1,
    extract_table_candidates,
)


def test_compute_candidate_id_stable():
    """Candidate IDs are stable for same inputs."""
    id1 = _compute_candidate_id(1, 10, 5, 20, 8, "grid")
    id2 = _compute_candidate_id(1, 10, 5, 20, 8, "grid")
    assert id1 == id2
    assert id1.startswith("tc_")
    assert len(id1) == 19  # 'tc_' + 16 hex chars


def test_compute_candidate_id_different_inputs():
    """Candidate IDs differ for different inputs."""
    id1 = _compute_candidate_id(1, 10, 5, 20, 8, "grid")
    id2 = _compute_candidate_id(1, 10, 5, 20, 9, "grid")  # Different c2
    id3 = _compute_candidate_id(1, 10, 5, 20, 8, "vector")  # Different kind

    assert id1 != id2
    assert id1 != id3
    assert id2 != id3


def test_range_to_a1():
    """Range conversion to A1 notation."""
    assert _range_to_a1(1, 1, 1, 1) == "A1"
    assert _range_to_a1(1, 1, 10, 5) == "A1:E10"
    assert _range_to_a1(5, 2, 10, 2) == "B5:B10"
    assert _range_to_a1(1, 27, 1, 27) == "AA1"


def test_extract_table_candidates_empty_db():
    """No candidates when no bindings exist."""
    with tempfile.TemporaryDirectory() as tmpdir:
        db_path = Path(tmpdir) / "test.db"
        conn = sqlite3.connect(str(db_path))

        # Create minimal schema
        conn.execute("CREATE TABLE sheets (sheet_id INTEGER PRIMARY KEY, sheet_name TEXT)")
        conn.execute(
            "CREATE TABLE cells (cell_id INTEGER PRIMARY KEY, sheet_id INTEGER, row INTEGER, col INTEGER)"
        )
        conn.execute("""
            CREATE TABLE bindings (
                binding_id TEXT PRIMARY KEY,
                sheet_id INTEGER,
                top_left_cell_id INTEGER,
                shape_rows INTEGER,
                shape_cols INTEGER,
                classification TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE table_candidates (
                candidate_id TEXT PRIMARY KEY,
                sheet_id INTEGER,
                kind TEXT,
                r1 INTEGER, c1 INTEGER, r2 INTEGER, c2 INTEGER,
                range_a1 TEXT,
                confidence REAL,
                reasons_top3_json TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE table_candidate_members (
                candidate_id TEXT,
                ordinal INTEGER,
                binding_id TEXT,
                role_hint TEXT,
                PRIMARY KEY (candidate_id, ordinal)
            )
        """)
        conn.commit()

        # Run extraction
        extract_table_candidates(conn=conn)

        # Verify no candidates
        count = conn.execute("SELECT COUNT(*) FROM table_candidates").fetchone()[0]
        assert count == 0

        conn.close()


def test_extract_table_candidates_single_grid():
    """Detect grid candidate from 3 adjacent column vectors."""
    with tempfile.TemporaryDirectory() as tmpdir:
        db_path = Path(tmpdir) / "test.db"
        conn = sqlite3.connect(str(db_path))

        # Create schema
        conn.execute("CREATE TABLE sheets (sheet_id INTEGER PRIMARY KEY, sheet_name TEXT)")
        conn.execute("INSERT INTO sheets VALUES (1, 'Sheet1')")

        conn.execute("""
            CREATE TABLE cells (
                cell_id INTEGER PRIMARY KEY,
                sheet_id INTEGER,
                row INTEGER,
                col INTEGER
            )
        """)

        # Create 3 cells at (10, 1), (10, 2), (10, 3) as top-left of bindings
        conn.execute("INSERT INTO cells VALUES (1, 1, 10, 1)")
        conn.execute("INSERT INTO cells VALUES (2, 1, 10, 2)")
        conn.execute("INSERT INTO cells VALUES (3, 1, 10, 3)")

        conn.execute("""
            CREATE TABLE bindings (
                binding_id TEXT PRIMARY KEY,
                sheet_id INTEGER,
                top_left_cell_id INTEGER,
                shape_rows INTEGER,
                shape_cols INTEGER,
                classification TEXT
            )
        """)

        # Create 3 adjacent column vectors (10 rows each, columns 1-3)
        conn.execute("INSERT INTO bindings VALUES ('b1', 1, 1, 10, 1, NULL)")
        conn.execute("INSERT INTO bindings VALUES ('b2', 1, 2, 10, 1, NULL)")
        conn.execute("INSERT INTO bindings VALUES ('b3', 1, 3, 10, 1, NULL)")

        conn.execute("""
            CREATE TABLE table_candidates (
                candidate_id TEXT PRIMARY KEY,
                sheet_id INTEGER,
                kind TEXT,
                r1 INTEGER, c1 INTEGER, r2 INTEGER, c2 INTEGER,
                range_a1 TEXT,
                confidence REAL,
                reasons_top3_json TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE table_candidate_members (
                candidate_id TEXT,
                ordinal INTEGER,
                binding_id TEXT,
                role_hint TEXT,
                PRIMARY KEY (candidate_id, ordinal)
            )
        """)
        conn.commit()

        # Run extraction
        extract_table_candidates(conn=conn)

        # Verify 1 grid candidate
        candidates = conn.execute("""
            SELECT candidate_id, kind, r1, c1, r2, c2, range_a1, confidence
            FROM table_candidates
        """).fetchall()

        assert len(candidates) == 1
        candidate_id, kind, r1, c1, r2, c2, range_a1, confidence = candidates[0]

        assert kind == "grid"
        assert r1 == 10
        assert c1 == 1
        assert r2 == 19  # 10 + 10 - 1
        assert c2 == 3
        assert range_a1 == "A10:C19"
        assert confidence > 0.7

        # Verify membership
        members = conn.execute(
            """
            SELECT ordinal, binding_id, role_hint
            FROM table_candidate_members
            WHERE candidate_id = ?
            ORDER BY ordinal
        """,
            (candidate_id,),
        ).fetchall()

        assert len(members) == 3
        assert members[0] == (0, "b1", "values")
        assert members[1] == (1, "b2", "values")
        assert members[2] == (2, "b3", "values")

        conn.close()


def test_extract_table_candidates_with_gap():
    """Gap in columns creates separate candidates."""
    with tempfile.TemporaryDirectory() as tmpdir:
        db_path = Path(tmpdir) / "test.db"
        conn = sqlite3.connect(str(db_path))

        # Create schema
        conn.execute("CREATE TABLE sheets (sheet_id INTEGER PRIMARY KEY, sheet_name TEXT)")
        conn.execute("INSERT INTO sheets VALUES (1, 'Sheet1')")

        conn.execute("""
            CREATE TABLE cells (
                cell_id INTEGER PRIMARY KEY,
                sheet_id INTEGER,
                row INTEGER,
                col INTEGER
            )
        """)

        # Create cells at columns 1, 2, 4, 5 (gap at column 3)
        conn.execute("INSERT INTO cells VALUES (1, 1, 10, 1)")
        conn.execute("INSERT INTO cells VALUES (2, 1, 10, 2)")
        conn.execute("INSERT INTO cells VALUES (3, 1, 10, 4)")
        conn.execute("INSERT INTO cells VALUES (4, 1, 10, 5)")

        conn.execute("""
            CREATE TABLE bindings (
                binding_id TEXT PRIMARY KEY,
                sheet_id INTEGER,
                top_left_cell_id INTEGER,
                shape_rows INTEGER,
                shape_cols INTEGER,
                classification TEXT
            )
        """)

        # Create 4 column vectors with gap
        conn.execute("INSERT INTO bindings VALUES ('b1', 1, 1, 10, 1, NULL)")
        conn.execute("INSERT INTO bindings VALUES ('b2', 1, 2, 10, 1, NULL)")
        conn.execute("INSERT INTO bindings VALUES ('b3', 1, 3, 10, 1, NULL)")  # Gap
        conn.execute("INSERT INTO bindings VALUES ('b4', 1, 4, 10, 1, NULL)")

        conn.execute("""
            CREATE TABLE table_candidates (
                candidate_id TEXT PRIMARY KEY,
                sheet_id INTEGER,
                kind TEXT,
                r1 INTEGER, c1 INTEGER, r2 INTEGER, c2 INTEGER,
                range_a1 TEXT,
                confidence REAL,
                reasons_top3_json TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE table_candidate_members (
                candidate_id TEXT,
                ordinal INTEGER,
                binding_id TEXT,
                role_hint TEXT,
                PRIMARY KEY (candidate_id, ordinal)
            )
        """)
        conn.commit()

        # Run extraction
        extract_table_candidates(conn=conn)

        # Verify 2 grid candidates (columns 1-2 and columns 4-5)
        candidates = conn.execute("""
            SELECT candidate_id, kind, c1, c2
            FROM table_candidates
            ORDER BY c1
        """).fetchall()

        assert len(candidates) == 2
        assert candidates[0][1] == "grid"
        assert candidates[0][2] == 1  # c1
        assert candidates[0][3] == 2  # c2

        assert candidates[1][1] == "grid"
        assert candidates[1][2] == 4  # c1
        assert candidates[1][3] == 5  # c2

        conn.close()


def test_extract_table_candidates_vector_classification():
    """Single column with classification becomes vector candidate."""
    with tempfile.TemporaryDirectory() as tmpdir:
        db_path = Path(tmpdir) / "test.db"
        conn = sqlite3.connect(str(db_path))

        # Create schema
        conn.execute("CREATE TABLE sheets (sheet_id INTEGER PRIMARY KEY, sheet_name TEXT)")
        conn.execute("INSERT INTO sheets VALUES (1, 'Sheet1')")

        conn.execute("""
            CREATE TABLE cells (
                cell_id INTEGER PRIMARY KEY,
                sheet_id INTEGER,
                row INTEGER,
                col INTEGER
            )
        """)
        conn.execute("INSERT INTO cells VALUES (1, 1, 10, 1)")

        conn.execute("""
            CREATE TABLE bindings (
                binding_id TEXT PRIMARY KEY,
                sheet_id INTEGER,
                top_left_cell_id INTEGER,
                shape_rows INTEGER,
                shape_cols INTEGER,
                classification TEXT
            )
        """)

        # Single column vector with classification (only 3 rows, below MIN_LENGTH_FOR_VECTOR)
        conn.execute("INSERT INTO bindings VALUES ('b1', 1, 1, 3, 1, 'input')")

        conn.execute("""
            CREATE TABLE table_candidates (
                candidate_id TEXT PRIMARY KEY,
                sheet_id INTEGER,
                kind TEXT,
                r1 INTEGER, c1 INTEGER, r2 INTEGER, c2 INTEGER,
                range_a1 TEXT,
                confidence REAL,
                reasons_top3_json TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE table_candidate_members (
                candidate_id TEXT,
                ordinal INTEGER,
                binding_id TEXT,
                role_hint TEXT,
                PRIMARY KEY (candidate_id, ordinal)
            )
        """)
        conn.commit()

        # Run extraction
        extract_table_candidates(conn=conn)

        # Verify 1 vector candidate
        candidates = conn.execute("""
            SELECT kind, reasons_top3_json
            FROM table_candidates
        """).fetchall()

        assert len(candidates) == 1
        kind, reasons_json = candidates[0]

        assert kind == "vector"
        reasons = json.loads(reasons_json)
        assert any("classification_input" in r for r in reasons)

        conn.close()


def test_extract_table_candidates_determinism():
    """Extraction is deterministic (same inputs → same outputs)."""
    with tempfile.TemporaryDirectory() as tmpdir:
        db_path = Path(tmpdir) / "test.db"

        def create_and_extract():
            conn = sqlite3.connect(str(db_path))

            # Create schema
            conn.execute("CREATE TABLE sheets (sheet_id INTEGER PRIMARY KEY, sheet_name TEXT)")
            conn.execute("INSERT INTO sheets VALUES (1, 'Sheet1')")

            conn.execute("""
                CREATE TABLE cells (
                    cell_id INTEGER PRIMARY KEY,
                    sheet_id INTEGER,
                    row INTEGER,
                    col INTEGER
                )
            """)
            conn.execute("INSERT INTO cells VALUES (1, 1, 10, 1)")
            conn.execute("INSERT INTO cells VALUES (2, 1, 10, 2)")

            conn.execute("""
                CREATE TABLE bindings (
                    binding_id TEXT PRIMARY KEY,
                    sheet_id INTEGER,
                    top_left_cell_id INTEGER,
                    shape_rows INTEGER,
                    shape_cols INTEGER,
                    classification TEXT
                )
            """)
            conn.execute("INSERT INTO bindings VALUES ('b1', 1, 1, 10, 1, NULL)")
            conn.execute("INSERT INTO bindings VALUES ('b2', 1, 2, 10, 1, NULL)")

            conn.execute("""
                CREATE TABLE table_candidates (
                    candidate_id TEXT PRIMARY KEY,
                    sheet_id INTEGER,
                    kind TEXT,
                    r1 INTEGER, c1 INTEGER, r2 INTEGER, c2 INTEGER,
                    range_a1 TEXT,
                    confidence REAL,
                    reasons_top3_json TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE table_candidate_members (
                    candidate_id TEXT,
                    ordinal INTEGER,
                    binding_id TEXT,
                    role_hint TEXT,
                    PRIMARY KEY (candidate_id, ordinal)
                )
            """)
            conn.commit()

            # Run extraction
            extract_table_candidates(conn=conn)

            # Get results
            candidate_id = conn.execute("SELECT candidate_id FROM table_candidates").fetchone()[0]
            members = conn.execute(
                """
                SELECT GROUP_CONCAT(binding_id, ',')
                FROM table_candidate_members
                WHERE candidate_id = ?
                ORDER BY ordinal
            """,
                (candidate_id,),
            ).fetchone()[0]

            conn.close()
            return candidate_id, members

        # Run twice
        id1, members1 = create_and_extract()
        db_path.unlink()  # Remove DB between runs
        id2, members2 = create_and_extract()

        # Verify determinism
        assert id1 == id2
        assert members1 == members2


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
