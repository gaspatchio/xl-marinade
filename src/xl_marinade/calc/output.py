"""Deterministic writers for calc outputs."""

from __future__ import annotations

import csv
from pathlib import Path

from xl_marinade.calc.graph import StoredValue
from xl_marinade.calc.models import CellRef, Reason
from xl_marinade.calc.rows import Row

HEADER = ["policy", "variable", "period", "value", "status", "reasons"]
_NO_ROOT = CellRef("", 0, 0)


def format_value(value: StoredValue) -> str:
    """CSV text for a value: ``repr`` floats, ``TRUE``/``FALSE``, text as-is, empty for None."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        return repr(value)
    return str(value)


def format_reasons(reasons: list[Reason]) -> str:
    """``kind@root:detail`` joined by ``; `` in a stable order (``kind:detail`` with no root)."""
    ordered = sorted(reasons, key=lambda r: (r.kind, r.root or _NO_ROOT, r.detail))
    return "; ".join(
        f"{r.kind}@{r.root.a1()}:{r.detail}" if r.root else f"{r.kind}:{r.detail}" for r in ordered
    )


def write_rows(path: Path, rows: list[Row]) -> None:
    """Write ``rows.csv`` (UTF-8, ``\\n`` line endings)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as fh:
        writer = csv.writer(fh, lineterminator="\n")
        writer.writerow(HEADER)
        for r in rows:
            writer.writerow(
                [
                    r.policy,
                    r.variable,
                    "" if r.period is None else r.period,
                    format_value(r.value),
                    r.status.value,
                    format_reasons(r.reasons),
                ]
            )
