# ABOUTME: PRAGMA policy enforcement for memory-efficient IR extraction with bounded memory and deterministic output.
# ABOUTME: Applies and verifies mandatory SQLite PRAGMAs before any heavy inserts.

import sqlite3
from typing import Any

# Mandatory PRAGMA configuration for canonical builds
MANDATORY_PRAGMAS: dict[str, Any] = {
    "page_size": 4096,
    "journal_mode": "OFF",
    "synchronous": "NORMAL",
    "temp_store": "FILE",
    "cache_size": -262144,  # 256 MB (negative = KB)
    "locking_mode": "EXCLUSIVE",
    "mmap_size": 0,
    "auto_vacuum": "NONE",
}


def apply_pragmas(conn: sqlite3.Connection) -> None:
    """
    Apply mandatory PRAGMAs to the given SQLite connection.

    Must be called before any table creation or heavy inserts.

    Args:
        conn: SQLite connection to configure

    Raises:
        RuntimeError: If any PRAGMA fails to set or verify

    Example:
        >>> import sqlite3
        >>> from xl_marinade.core.new_arch import apply_pragmas
        >>> conn = sqlite3.connect("ir.db")
        >>> apply_pragmas(conn)  # Must be called before CREATE TABLE
        >>> conn.execute("CREATE TABLE cells (...)")
    """
    cursor = conn.cursor()

    for pragma, expected_value in MANDATORY_PRAGMAS.items():
        # Set the PRAGMA
        cursor.execute(f"PRAGMA {pragma} = {_format_pragma_value(expected_value)}")

        # Verify it was set correctly
        cursor.execute(f"PRAGMA {pragma}")
        actual_value = cursor.fetchone()[0]

        # Normalize for comparison
        normalized_expected = _normalize_pragma_value(pragma, expected_value)
        normalized_actual = _normalize_pragma_value(pragma, actual_value)

        if normalized_actual != normalized_expected:
            raise RuntimeError(
                f"PRAGMA {pragma} verification failed: "
                f"expected {normalized_expected}, got {normalized_actual}"
            )


def verify_pragmas(conn: sqlite3.Connection) -> None:
    """
    Verify that all mandatory PRAGMAs are set correctly.

    Args:
        conn: SQLite connection to verify

    Raises:
        RuntimeError: If any PRAGMA does not match expected value

    Example:
        >>> import sqlite3
        >>> from xl_marinade.core.new_arch import apply_pragmas, verify_pragmas
        >>> conn = sqlite3.connect("ir.db")
        >>> apply_pragmas(conn)
        >>> verify_pragmas(conn)  # Raises if any PRAGMA mismatch
    """
    cursor = conn.cursor()

    for pragma, expected_value in MANDATORY_PRAGMAS.items():
        cursor.execute(f"PRAGMA {pragma}")
        actual_value = cursor.fetchone()[0]

        normalized_expected = _normalize_pragma_value(pragma, expected_value)
        normalized_actual = _normalize_pragma_value(pragma, actual_value)

        if normalized_actual != normalized_expected:
            raise RuntimeError(
                f"PRAGMA {pragma} mismatch: expected {normalized_expected}, got {normalized_actual}"
            )


def get_pragma_config(conn: sqlite3.Connection) -> dict[str, Any]:
    """
    Get current PRAGMA configuration from the connection.

    Returns normalized values for all mandatory PRAGMAs without raising errors.
    Useful for debugging and introspection.

    Args:
        conn: SQLite connection to query

    Returns:
        Dictionary mapping PRAGMA names to their current normalized values

    Example:
        >>> import sqlite3
        >>> from xl_marinade.core.new_arch import apply_pragmas, get_pragma_config
        >>> conn = sqlite3.connect("ir.db")
        >>> apply_pragmas(conn)
        >>> config = get_pragma_config(conn)
        >>> config['page_size']
        4096
        >>> config['temp_store']
        1
    """
    cursor = conn.cursor()
    config = {}

    for pragma in MANDATORY_PRAGMAS:
        cursor.execute(f"PRAGMA {pragma}")
        actual_value = cursor.fetchone()[0]
        config[pragma] = _normalize_pragma_value(pragma, actual_value)

    return config


def _format_pragma_value(value: Any) -> str:
    """Format a PRAGMA value for SQL execution."""
    if isinstance(value, str):
        return f"'{value}'"
    return str(value)


def _normalize_pragma_value(pragma: str, value: Any) -> Any:
    """
    Normalize a PRAGMA value for comparison.

    Some PRAGMAs return values in different formats than they accept.
    """
    # journal_mode returns uppercase
    if pragma == "journal_mode" and isinstance(value, str):
        return value.upper()

    # synchronous can be returned as integer (0=OFF, 1=NORMAL, 2=FULL)
    if pragma == "synchronous":
        if value == "NORMAL" or value == 1:
            return 1
        if value == "OFF" or value == 0:
            return 0
        if value == "FULL" or value == 2:
            return 2

    # auto_vacuum can be returned as integer (0=NONE, 1=FULL, 2=INCREMENTAL)
    if pragma == "auto_vacuum":
        if value == "NONE" or value == 0:
            return 0
        if value == "FULL" or value == 1:
            return 1
        if value == "INCREMENTAL" or value == 2:
            return 2

    # temp_store can be returned as integer (0=DEFAULT, 1=FILE, 2=MEMORY)
    if pragma == "temp_store":
        if value == "FILE" or value == 1:
            return 1
        if value == "MEMORY" or value == 2:
            return 2
        if value == "DEFAULT" or value == 0:
            return 0

    # locking_mode returns uppercase
    if pragma == "locking_mode" and isinstance(value, str):
        return value.upper()

    return value
