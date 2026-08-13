# SPDX-FileCopyrightText: 2026 Klaas Stijnen
#
# SPDX-License-Identifier: MIT
"""Safe SQLite URI construction for read-only database opens.

Interpolating a filesystem path into ``f"file:{path}?mode=ro"`` is wrong: the
text after ``file:`` is a URI, not a path, so SQLite reinterprets several
characters that are perfectly legal in a filename.

* ``#`` starts a URI fragment — the path is truncated there AND the trailing
  ``?mode=ro`` is swallowed with it, so the wrong file is opened *and* the
  read-only guarantee is silently lost (SQLite then creates the truncated
  path as a new, empty database).
* ``%HH`` percent-decodes — ``50%20off`` resolves to ``50 off``, opening a
  different file or failing with a confusing "unable to open database file".
* ``?`` starts the query string early, mangling both path and parameters.

``Path.as_uri()`` percent-encodes all of these, so it round-trips through
SQLite's URI parser back to the original path. Windows paths are covered too:
``as_uri()`` produces ``file:///C:/dir/db.sqlite`` from a ``C:\\dir`` path.
"""

import sqlite3
from pathlib import Path


def read_only_uri(path: str | Path) -> str:
    """Return a SQLite URI opening ``path`` read-only, safe for any filename."""
    return Path(path).resolve().as_uri() + "?mode=ro"


def connect_read_only(path: str | Path) -> sqlite3.Connection:
    """Open ``path`` read-only via a correctly-encoded SQLite URI."""
    return sqlite3.connect(read_only_uri(path), uri=True)


def attach_read_only(conn: sqlite3.Connection, path: str | Path, alias: str) -> None:
    """ATTACH ``path`` read-only under ``alias``.

    The filename is bound as a parameter rather than interpolated into SQL:
    an apostrophe in a directory name (``C:\\Users\\O'Brien\\...``) would
    otherwise terminate the string literal and raise a syntax error. The
    connection must have been opened with ``uri=True`` for SQLite to parse the
    bound value as a URI — every caller here uses ``connect_read_only`` or
    passes ``uri=True`` explicitly.
    """
    conn.execute(f"ATTACH DATABASE ? AS {alias}", (read_only_uri(path),))
