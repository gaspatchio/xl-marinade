"""Deterministic writers for calc outputs."""

from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from xl_marinade.calc.engine import EngineValue
from xl_marinade.calc.graph import StoredValue
from xl_marinade.calc.models import CellRef, Reason, Status
from xl_marinade.calc.rows import Row

HEADER = ["policy", "variable", "period", "value", "status", "reasons"]
RESULTS_HEADER = ["policy_id", "output", "value", "status", "reasons"]
PATHS_HEADER = ["policy_id", "cell", "family", "kind", "ordinal", "outcome", "exercised"]
_NO_ROOT = CellRef("", 0, 0)


@dataclass(slots=True, frozen=True)
class OutputValue:
    """One recalculated output cell for one policy."""

    policy_id: str
    cell: CellRef
    value: EngineValue
    status: Status
    reasons: list[Reason]


@dataclass(slots=True, frozen=True)
class PathRecord:
    """One branch or lookup-edge probe reading for one policy.

    ``ordinal`` tells apart several probes of one kind in the same cell.
    """

    policy_id: str
    cell: CellRef
    family: str
    kind: str
    outcome: str
    exercised: bool
    ordinal: int = 0


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


def _csv(path: Path, header: list[str], rows: list[list[object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as fh:
        writer = csv.writer(fh, lineterminator="\n")
        writer.writerow(header)
        writer.writerows(rows)


def _json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8", newline="\n")


def write_results(path: Path, values: list[OutputValue]) -> None:
    """Write ``results.csv`` in the given order (UTF-8, ``\\n`` line endings)."""
    _csv(
        path,
        RESULTS_HEADER,
        [
            [
                v.policy_id,
                v.cell.a1(),
                format_value(v.value),
                v.status.value,
                format_reasons(v.reasons),
            ]
            for v in values
        ],
    )


def write_paths(path: Path, records: list[PathRecord]) -> None:
    """Write ``paths.csv`` in the given order; ``exercised`` is ``true``/``false``."""
    _csv(
        path,
        PATHS_HEADER,
        [
            [
                r.policy_id,
                r.cell.a1(),
                r.family,
                r.kind,
                r.ordinal,
                r.outcome,
                "true" if r.exercised else "false",
            ]
            for r in records
        ],
    )


def write_trust_report(path: Path, report: dict[str, Any]) -> None:
    """Write ``trust_report.json`` (keys in the given order, no timestamps)."""
    _json(path, report)


def write_run(path: Path, record: dict[str, Any]) -> None:
    """Write ``run.json`` (keys in the given order, no timestamps)."""
    _json(path, record)
