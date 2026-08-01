# ABOUTME: IR schema compatibility helpers for querying dependency edges deterministically.
# ABOUTME: Supports multiple IR schemas (agent_* fast, new_arch fast, legacy).

from __future__ import annotations

import sqlite3
from dataclasses import dataclass


@dataclass(frozen=True)
class DependencyEdgesSpec:
    table: str
    from_col: str
    to_col: str


def _table_exists(conn: sqlite3.Connection, schema: str | None, table: str) -> bool:
    try:
        if schema:
            row = conn.execute(
                f"SELECT 1 FROM {schema}.sqlite_master WHERE type='table' AND name=? LIMIT 1",
                (table,),
            ).fetchone()
        else:
            row = conn.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1",
                (table,),
            ).fetchone()
        return row is not None
    except sqlite3.OperationalError:
        return False


def _resolve_edge_columns(
    conn: sqlite3.Connection, schema: str | None, table: str
) -> tuple[str, str] | None:
    try:
        pragma = f"PRAGMA {schema}.table_info({table})" if schema else f"PRAGMA table_info({table})"
        cols = {row[1] for row in conn.execute(pragma)}
    except sqlite3.OperationalError:
        return None

    candidates = [
        ("from_binding", "to_binding"),
        ("from_binding_id", "to_binding_id"),
        ("from_binding", "to_binding_id"),
        ("from_binding_id", "to_binding"),
    ]
    for from_col, to_col in candidates:
        if from_col in cols and to_col in cols:
            return from_col, to_col
    return None


def detect_dependency_edges(
    conn: sqlite3.Connection, schema: str | None = None
) -> DependencyEdgesSpec | None:
    """
    Detect the best available dependency edge table and columns.

    Supports:
    - agent fast schema: agent_binding_dependencies(from_binding,to_binding)
    - new_arch fast schema: binding_edges(from_binding_id,to_binding_id)
    - legacy IR schema: binding_level_edges(from_binding_id,to_binding_id)

    Args:
        conn: SQLite connection (optionally with an attached database).
        schema: Optional schema name (e.g. "ir" when IR is attached to overlay).

    Returns:
        DependencyEdgesSpec or None if no known edge table exists.
    """
    for table in ("agent_binding_dependencies", "binding_edges", "binding_level_edges"):
        if _table_exists(conn, schema, table):
            cols = _resolve_edge_columns(conn, schema, table)
            if cols:
                from_col, to_col = cols
                return DependencyEdgesSpec(table=table, from_col=from_col, to_col=to_col)

    return None
