"""Deterministic writers for calc outputs."""

from __future__ import annotations

import csv
import json
import os
from dataclasses import dataclass
from pathlib import Path
from types import TracebackType
from typing import IO, Any, Self

from loguru import logger

from xl_marinade.calc.engine import EngineValue
from xl_marinade.calc.graph import StoredValue
from xl_marinade.calc.models import CellRef, Reason, Status
from xl_marinade.calc.rows import Row
from xl_marinade.errors import CalcInputError

HEADER = ["policy", "variable", "period", "value", "status", "reasons"]
RESULTS_HEADER = ["policy_id", "output", "value", "status", "reasons"]
CALC_OUTPUTS = ("results.csv", "rows.csv", "paths.csv", "trust_report.json", "run.json")
VALUES_OUTPUTS = ("rows.csv", "run.json")
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


def _row_fields(r: Row) -> list[object]:
    return [
        r.policy,
        r.variable,
        "" if r.period is None else r.period,
        format_value(r.value),
        r.status.value,
        format_reasons(r.reasons),
    ]


def _result_fields(v: OutputValue) -> list[object]:
    return [
        v.policy_id,
        v.cell.a1(),
        format_value(v.value),
        v.status.value,
        format_reasons(v.reasons),
    ]


def _path_fields(r: PathRecord) -> list[object]:
    return [
        r.policy_id,
        r.cell.a1(),
        r.family,
        r.kind,
        r.ordinal,
        r.outcome,
        "true" if r.exercised else "false",
    ]


def write_rows(path: Path, rows: list[Row]) -> None:
    """Write ``rows.csv`` (UTF-8, ``\\n`` line endings)."""
    _csv(path, HEADER, [_row_fields(r) for r in rows])


def _entry(path: Path) -> Path:
    """The directory entry ``path`` names: its parent resolved, its own name kept, so a
    symlink is the link and not its target."""
    return path.absolute().parent.resolve() / path.name


def guard_out_dir(out_dir: Path, names: tuple[str, ...], inputs: list[Path | None]) -> None:
    """Refuse, before anything is cleared or written, an ``out_dir`` that is not a
    directory and an input the run would clear from it.

    An input is cleared when it is one of ``names`` in ``out_dir``: the same directory
    entry, or the same file through another path (a hard link, a symlink to it, another
    spelling on a case-insensitive file system).

    Raises:
        CalcInputError: naming ``out_dir``, or the input and the output file.
    """
    for parent in (out_dir, *out_dir.parents):
        if parent.exists():
            if not parent.is_dir():
                raise CalcInputError(
                    f"-o {out_dir}: {parent} is not a directory; give a directory for the outputs"
                )
            break
    given = [path for path in inputs if path is not None]
    for name in names:
        output = out_dir / name
        if not (output.exists() or output.is_symlink()):
            continue
        for path in given:
            same_entry = _entry(path) == _entry(output)
            same_file = not output.is_symlink() and path.exists() and os.path.samefile(path, output)
            if same_entry or same_file:
                raise CalcInputError(
                    f"input {path} is the output file {output}, which the run clears from "
                    f"-o {out_dir} before writing; move the input or give another -o"
                )


def clear_outputs(out_dir: Path, names: tuple[str, ...]) -> None:
    """Delete the named output files a previous run left in ``out_dir`` (nothing else)."""
    for name in names:
        path = out_dir / name
        if path.is_file() or path.is_symlink():
            path.unlink()


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
    _csv(path, RESULTS_HEADER, [_result_fields(v) for v in values])


def write_paths(path: Path, records: list[PathRecord]) -> None:
    """Write ``paths.csv`` in the given order; ``exercised`` is ``true``/``false``."""
    _csv(path, PATHS_HEADER, [_path_fields(r) for r in records])


class _Stream:
    """One CSV file written row by row."""

    def __init__(self, path: Path, header: list[str]) -> None:
        self.path = path
        path.parent.mkdir(parents=True, exist_ok=True)
        self._fh: IO[str] = path.open("w", encoding="utf-8", newline="\n")
        try:
            self._writer = csv.writer(self._fh, lineterminator="\n")
            self._writer.writerow(header)
            self._fh.flush()
        except BaseException:
            self.discard()
            raise

    def write(self, rows: list[list[object]]) -> None:
        self._writer.writerows(rows)

    def flush(self) -> None:
        self._fh.flush()

    def close(self) -> None:
        self._fh.close()

    def discard(self) -> None:
        """Close and delete the file. A failing close is logged, not raised: the file
        goes either way, and the error that led here is the one worth reporting."""
        try:
            self.close()
        except OSError as exc:
            logger.warning(f"closing {self.path} before deleting it failed: {exc}")
        finally:
            self.path.unlink(missing_ok=True)


class PolicyWriter:
    """Streams ``results.csv``, ``paths.csv`` and (optionally) ``rows.csv`` one policy at
    a time, so a run holds no more than one policy's outputs in memory.

    The files are byte-identical to :func:`write_results`, :func:`write_paths` and
    :func:`write_rows` given the same records in the same order. Leaving the ``with``
    block on an exception deletes all three files: a run that stopped part-way must not
    leave results that look complete.
    """

    def __init__(self, results: Path, paths: Path, rows: Path | None) -> None:
        self._streams: list[_Stream] = []
        try:
            self._results = self._open(results, RESULTS_HEADER)
            self._paths = self._open(paths, PATHS_HEADER)
            self._rows = self._open(rows, HEADER) if rows is not None else None
        except BaseException:
            self._discard()
            raise

    def _open(self, path: Path, header: list[str]) -> _Stream:
        stream = _Stream(path, header)
        self._streams.append(stream)
        return stream

    def write_policy(
        self, values: list[OutputValue], paths: list[PathRecord], rows: list[Row]
    ) -> None:
        """Append one policy's records and flush them to disk."""
        self._results.write([_result_fields(v) for v in values])
        self._paths.write([_path_fields(r) for r in paths])
        if self._rows is not None:
            self._rows.write([_row_fields(r) for r in rows])
        for stream in self._streams:
            stream.flush()

    def _discard(self) -> None:
        for stream in self._streams:
            stream.discard()

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        if exc_type is not None:
            self._discard()
            return
        for stream in self._streams:
            stream.close()


def write_trust_report(path: Path, report: dict[str, Any]) -> None:
    """Write ``trust_report.json`` (keys in the given order, no timestamps)."""
    _json(path, report)


def write_run(path: Path, record: dict[str, Any]) -> None:
    """Write ``run.json`` (keys in the given order, no timestamps)."""
    _json(path, record)
