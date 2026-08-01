# ABOUTME: Deterministic time axis candidate detection and ranking for Sprint 6.
# ABOUTME: Loads 1D binding sequences, scores candidates, and emits top ranks per sheet.
from __future__ import annotations

import json
import sqlite3
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path

from xl_marinade.core.ref_converter import parse_cell_address

_EXCEL_BASE_DATE = date(1899, 12, 30)
_MIN_SEQUENCE_LENGTH = 3
_REL_TOL = 1e-6
_ABS_TOL = 1e-6


@dataclass(frozen=True)
class SequenceBinding:
    """Represents a 1D binding sequence used for time axis detection."""

    sheet: str
    binding_id: str
    top_left_row: int
    top_left_col: int
    axis: str  # "row" or "col"
    values: list[float]


@dataclass(frozen=True)
class _ScoredCandidate:
    sheet: str
    binding_id: str
    confidence: float
    reasons_top3: list[str]
    length: int
    top_left_row: int
    top_left_col: int


def infer_time_index_candidates(db_path: Path) -> list[dict[str, object]]:
    """
    Infer time index candidates from an existing ir.db file.

    Returns:
        List of dicts with sheet, binding_id, rank, confidence, reasons_top3_json.
    """
    conn = sqlite3.connect(str(db_path))
    try:
        return infer_time_index_candidates_from_conn(conn)
    finally:
        conn.close()


def infer_time_index_candidates_from_conn(conn: sqlite3.Connection) -> list[dict[str, object]]:
    """
    Infer time index candidates from an open SQLite connection.

    This is required for fast-pipeline post-processing where the build connection
    may use exclusive locking (a second connection would fail with 'database is locked').
    """
    sequence_bindings = _load_sequence_bindings_from_conn(conn)
    return infer_time_index_candidates_for_bindings(sequence_bindings)


def infer_time_index_candidates_for_bindings(
    bindings: Iterable[SequenceBinding],
) -> list[dict[str, object]]:
    """Infer and rank candidates from a list of SequenceBinding objects."""
    by_sheet: dict[str, list[_ScoredCandidate]] = {}
    for binding in bindings:
        candidate = _score_binding(binding)
        if candidate is None:
            continue
        by_sheet.setdefault(candidate.sheet, []).append(candidate)

    results: list[dict[str, object]] = []
    for sheet, candidates in sorted(by_sheet.items()):
        ranked = _rank_candidates(candidates)
        for rank, candidate in enumerate(ranked, start=1):
            results.append(
                {
                    "sheet": sheet,
                    "binding_id": candidate.binding_id,
                    "rank": rank,
                    "confidence": candidate.confidence,
                    "reasons_top3_json": candidate.reasons_top3,
                }
            )
    return results


def persist_time_index_candidates(db_path: Path, candidates: list[dict[str, object]]) -> None:
    """Persist time index candidates into ir.db using deterministic ordering."""
    from xl_marinade.core import schema as ir_schema

    conn = ir_schema.open_existing_database(db_path)
    try:
        ir_schema.insert_time_index_candidates(conn, candidates)
    finally:
        conn.close()


def _load_sequence_bindings(db_path: Path) -> list[SequenceBinding]:
    conn = sqlite3.connect(str(db_path))
    try:
        return _load_sequence_bindings_from_conn(conn)
    finally:
        conn.close()


def _load_sequence_bindings_from_conn(conn: sqlite3.Connection) -> list[SequenceBinding]:
    is_legacy = _legacy_cells_available(conn)
    bindings = _fetch_1d_bindings(conn)
    sequence_bindings: list[SequenceBinding] = []
    for binding in bindings:
        values = _load_binding_values(conn, binding, is_legacy=is_legacy)
        if values is None or len(values) < _MIN_SEQUENCE_LENGTH:
            continue
        sequence_bindings.append(
            SequenceBinding(
                sheet=binding["sheet"],
                binding_id=binding["binding_id"],
                top_left_row=binding["top_left_row"],
                top_left_col=binding["top_left_col"],
                axis=binding["axis"],
                values=values,
            )
        )
    return sequence_bindings


def _legacy_cells_available(conn: sqlite3.Connection) -> bool:
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(cells)")
    return any(row[1] == "cell_address_a1" for row in cursor.fetchall())


def _fetch_1d_bindings(conn: sqlite3.Connection) -> list[dict[str, object]]:
    if _binding_has_sheet_column(conn):
        rows = _fetch_1d_bindings_legacy(conn)
    else:
        rows = _fetch_1d_bindings_fast(conn)
    bindings: list[dict[str, object]] = []
    for binding_id, sheet, address, shape_rows, shape_cols in rows:
        if shape_rows == 1 and shape_cols >= _MIN_SEQUENCE_LENGTH:
            axis = "row"
        elif shape_cols == 1 and shape_rows >= _MIN_SEQUENCE_LENGTH:
            axis = "col"
        else:
            continue
        parsed = parse_cell_address(address)
        bindings.append(
            {
                "binding_id": binding_id,
                "sheet": sheet,
                "axis": axis,
                "top_left_row": int(parsed["row"]),
                "top_left_col": int(parsed["col"]),
            }
        )
    return bindings


def _binding_has_sheet_column(conn: sqlite3.Connection) -> bool:
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(bindings)")
    columns = {row[1] for row in cursor.fetchall()}
    return "sheet" in columns


def _fetch_1d_bindings_legacy(
    conn: sqlite3.Connection,
) -> list[tuple[str, str, str, int, int]]:
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT binding_id, sheet, top_left_a1, shape_rows, shape_cols
        FROM bindings
        ORDER BY binding_id
        """
    )
    return cursor.fetchall()


def _fetch_1d_bindings_fast(
    conn: sqlite3.Connection,
) -> list[tuple[str, str, str, int, int]]:
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT binding_id, sheet, address, shape_rows, shape_cols
        FROM agent_bindings
        ORDER BY binding_id
        """
    )
    return cursor.fetchall()


def _load_binding_values(
    conn: sqlite3.Connection,
    binding: dict[str, object],
    *,
    is_legacy: bool,
) -> list[float] | None:
    if is_legacy:
        rows = _fetch_binding_cells_legacy(conn, binding["binding_id"])
    else:
        rows = _fetch_binding_cells_fast(conn, binding["binding_id"])
    if not rows:
        return None
    values: list[float] = []
    for cell_address, evaluated_value, value_snapshot in rows:
        value = _select_numeric_value(evaluated_value, value_snapshot)
        if value is None:
            return None
        values.append(value)
    return values


def _fetch_binding_cells_legacy(
    conn: sqlite3.Connection,
    binding_id: str,
) -> list[tuple[str, object | None, object | None]]:
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT cell_address_a1, evaluated_value, value_snapshot
        FROM cells
        WHERE binding_id = ?
        """,
        (binding_id,),
    )
    rows = cursor.fetchall()
    return _sort_cells(rows, address_index=0)


def _fetch_binding_cells_fast(
    conn: sqlite3.Connection,
    binding_id: str,
) -> list[tuple[str, object | None, object | None]]:
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT ac.cell_address, ac.value, NULL
        FROM agent_cells ac
        JOIN cell_to_binding ctb ON ac.cell_id = ctb.cell_id
        WHERE ctb.binding_id = ?
        """,
        (binding_id,),
    )
    rows = cursor.fetchall()
    return _sort_cells(rows, address_index=0)


def _sort_cells(
    rows: list[tuple[object, object | None, object | None]],
    *,
    address_index: int,
) -> list[tuple[str, object | None, object | None]]:
    def sort_key(row: tuple[object, object | None, object | None]) -> tuple[int, int]:
        parsed = parse_cell_address(str(row[address_index]))
        return (int(parsed["row"]), int(parsed["col"]))

    return sorted(rows, key=sort_key)


def _select_numeric_value(
    evaluated_value: object | None,
    value_snapshot: object | None,
) -> float | None:
    numeric = _parse_numeric_value(evaluated_value)
    if numeric is not None:
        return numeric
    return _parse_numeric_value(value_snapshot)


def _parse_numeric_value(value: object | None) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, (bytes, bytearray)):
        return None
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
        except json.JSONDecodeError:
            try:
                return float(value)
            except ValueError:
                return None
        return _parse_numeric_value(parsed)
    return None


def _score_binding(binding: SequenceBinding) -> _ScoredCandidate | None:
    if not _is_monotone_increasing(binding.values):
        return None
    constant_step = _is_constant_step(binding.values)
    calendar_month_step = _is_calendar_month_step(binding.values)

    length_score = min(len(binding.values) / 12.0, 1.0)
    position_score = 1.0 / (1.0 + min(binding.top_left_row, binding.top_left_col))

    score = 0.05 + (0.1 * length_score) + (0.1 * position_score)
    if constant_step:
        score += 0.45
    if calendar_month_step:
        score += 0.45
    confidence = min(score, 1.0)

    reasons = _rank_reasons(
        binding=binding,
        length_score=length_score,
        position_score=position_score,
        constant_step=constant_step,
        calendar_month_step=calendar_month_step,
    )

    return _ScoredCandidate(
        sheet=binding.sheet,
        binding_id=binding.binding_id,
        confidence=confidence,
        reasons_top3=reasons,
        length=len(binding.values),
        top_left_row=binding.top_left_row,
        top_left_col=binding.top_left_col,
    )


def _rank_candidates(candidates: list[_ScoredCandidate]) -> list[_ScoredCandidate]:
    ranked = sorted(
        candidates,
        key=lambda c: (
            -c.confidence,
            -c.length,
            c.top_left_row,
            c.top_left_col,
            c.binding_id,
        ),
    )
    return ranked[:3]


def _rank_reasons(
    *,
    binding: SequenceBinding,
    length_score: float,
    position_score: float,
    constant_step: bool,
    calendar_month_step: bool,
) -> list[str]:
    reason_scores: dict[str, float] = {
        "monotone_increasing": 0.05,
        f"long_run_len={len(binding.values)}": length_score,
        f"position_bias_row={binding.top_left_row}_col={binding.top_left_col}": position_score,
    }
    if constant_step:
        reason_scores["constant_step"] = 0.45
    if calendar_month_step:
        reason_scores["calendar_month_step"] = 0.45
    ordered = sorted(
        reason_scores.items(),
        key=lambda item: (-item[1], item[0]),
    )
    return [reason for reason, _ in ordered[:3]]


def _is_monotone_increasing(values: list[float]) -> bool:
    if len(values) < 2:
        return False
    for prev, current in zip(values, values[1:], strict=False):
        if current <= prev + _ABS_TOL:
            return False
    return True


def _is_constant_step(values: list[float]) -> bool:
    if len(values) < _MIN_SEQUENCE_LENGTH:
        return False
    diffs = [b - a for a, b in zip(values, values[1:], strict=False)]
    if any(diff <= 0 for diff in diffs):
        return False
    step = diffs[0]
    for diff in diffs[1:]:
        if abs(diff - step) > max(_ABS_TOL, abs(step) * _REL_TOL):
            return False
    return True


def _is_calendar_month_step(values: list[float]) -> bool:
    if len(values) < _MIN_SEQUENCE_LENGTH:
        return False
    serials: list[int] = []
    for value in values:
        rounded = round(value)
        if abs(value - rounded) > 1e-4:
            return False
        serials.append(int(rounded))
    try:
        dates = [_excel_serial_to_date(serial) for serial in serials]
    except (OverflowError, ValueError):
        # Serial out of range for Python's date (e.g. an arbitrary large number
        # that happens to look integer-like but isn't a real Excel serial).
        # This binding simply isn't a calendar month sequence.
        return False

    if _check_sequence_pairwise_sticky(dates):
        return True
    if _check_sequence_pairwise_simple(dates):
        return True
    if _check_relative_strategy(dates, sticky=True):
        return True
    if _check_relative_strategy(dates, sticky=False):
        return True
    return False


def _check_sequence_pairwise_sticky(dates: list[date]) -> bool:
    for prev, current in zip(dates, dates[1:], strict=False):
        if current != _add_month_sticky(prev):
            return False
    return True


def _check_sequence_pairwise_simple(dates: list[date]) -> bool:
    for prev, current in zip(dates, dates[1:], strict=False):
        if current != _add_month_simple(prev):
            return False
    return True


def _check_relative_strategy(dates: list[date], sticky: bool) -> bool:
    start = dates[0]
    target_day = start.day
    is_start_eom = start.day == _last_day_of_month(start.year, start.month)

    for i, current in enumerate(dates[1:], start=1):
        year = start.year + (start.month + i - 1) // 12
        month = (start.month + i - 1) % 12 + 1
        last = _last_day_of_month(year, month)

        if sticky and is_start_eom:
            day = last
        else:
            day = min(target_day, last)

        if current != date(year, month, day):
            return False
    return True


def _excel_serial_to_date(serial: int) -> date:
    return _EXCEL_BASE_DATE + timedelta(days=serial)


def _add_month_sticky(value: date) -> date:
    year = value.year + (value.month // 12)
    month = value.month % 12 + 1
    last_day = _last_day_of_month(year, month)
    if value.day == _last_day_of_month(value.year, value.month):
        day = last_day
    else:
        day = min(value.day, last_day)
    return date(year, month, day)


def _add_month_simple(value: date) -> date:
    year = value.year + (value.month // 12)
    month = value.month % 12 + 1
    last = _last_day_of_month(year, month)
    day = min(value.day, last)
    return date(year, month, day)


def _last_day_of_month(year: int, month: int) -> int:
    if month == 12:
        next_month = date(year + 1, 1, 1)
    else:
        next_month = date(year, month + 1, 1)
    return (next_month - timedelta(days=1)).day
