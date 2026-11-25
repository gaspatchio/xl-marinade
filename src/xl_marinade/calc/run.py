"""``marinade calc``: recalculate a workbook per policy and say how far to trust each value.

A run reads the workbook settings (refusing the whole workbook when a setting changes
what a recalculation means), loads the formula graph, scans for refused patterns, loads
a cache-free engine copy with probe cells, runs the trust run on the stored inputs and
any ``--check-against`` input sets, then recalculates every policy and assigns each
output a status from its root causes. Temporary files live in a temporary directory
that is always deleted; only ``out_dir`` is written.
"""

from __future__ import annotations

import hashlib
import math
import re
import tempfile
from collections.abc import Container, Iterable
from dataclasses import asdict, dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, NoReturn

from xl_marinade.calc.config import (
    load_check_against,
    load_input_map,
    load_policies,
    load_variables,
    parse_outputs,
)
from xl_marinade.calc.engine import Engine, EngineValue, engine_version, quiet_native_output
from xl_marinade.calc.graph import ListItems, StoredValue, WorkbookGraph, open_graph
from xl_marinade.calc.models import (
    ArrayRange,
    CellRef,
    CheckAgainst,
    Policy,
    Reason,
    ReasonKind,
    Status,
    VariableSpec,
    WorkbookSettings,
    _col_to_number,
)
from xl_marinade.calc.output import (
    CALC_OUTPUTS,
    OutputValue,
    PathRecord,
    PolicyWriter,
    clear_outputs,
    guard_out_dir,
    write_run,
    write_trust_report,
)
from xl_marinade.calc.patterns import (
    NOT_HELD,
    RELATIVE_NAME,
    input_dependent_lookups,
    scan_patterns,
)
from xl_marinade.calc.prepare import PreparedCopy, count_external_refs, prepare_engine_copy
from xl_marinade.calc.probes import (
    ProbeKind,
    ProbePlan,
    ProbeReadings,
    engine_error_code,
    place_probes,
    plan_probes,
    probe_kinds,
    reach_sheets,
    read_probes,
)
from xl_marinade.calc.reads import HiddenReads, field_read, hidden_reads
from xl_marinade.calc.rows import VariableCells, expand_variables, rows_for, variables_record
from xl_marinade.calc.settings import read_settings, refusal_reasons
from xl_marinade.calc.status import (
    MAX_REASONS_PER_KIND,
    NAME_OFF_PATH,
    Exercised,
    ExercisedBuilder,
    LandingCheck,
    PathKey,
    PathSignature,
    StaticLabels,
    assign,
    engine_error_detail,
    path_keys,
    path_signature,
    policy_roots,
    propagate_static,
    reason_key,
    same_path,
    unplaced_roots,
)
from xl_marinade.calc.trust import TrustResult, run_trust, same
from xl_marinade.errors import CalcError, CalcInputError, WorkbookRefused

_ISO_DATE = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}")
# Excel's number entry, in ASCII: optional sign, digits with an optional fraction (or a
# bare fraction), optional exponent. No "_" separators, no inf/nan, no other scripts.
_PLAIN_NUMBER = re.compile(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?")
_EXCEL_MAX = 9.99999999999999e307  # the largest number Excel accepts typed into a cell
_EXCEL_MIN = 2.2251e-308  # the smallest positive number Excel holds
_TIME = re.compile(r"[0-9]{1,2}:[0-9]{1,2}(?::[0-9]{1,2}(?:\.[0-9]+)?)?(?:\s*[AaPp][Mm])?")
_DATE_TIME = re.compile(
    r"[0-9]{4}-[0-9]{1,2}-[0-9]{1,2}[Tt ]+[0-9]{1,2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]+)?)?Z?"
)
# Day, month and year in any order or with two parts (``30/06/2024``, ``1/2``, ``2024-6``).
_OTHER_DATE = re.compile(r"([0-9]{1,4})[/.-]([0-9]{1,4})(?:[/.-]([0-9]{1,4}))?")
_MONTHS = (
    "jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april"
    "|june|july|august|september|october|november|december"
)
# A month name with day and/or year numbers around it: 30-Jun-2024, June 30, 2024,
# Jun 2024, Jun-24. The month alone, or with other words, is text.
_MONTH_DATE = re.compile(
    rf"(?:[0-9]{{1,4}}[ ./,-]+)?(?:{_MONTHS})\.?(?:[ ./,-]+[0-9]{{1,4}}){{0,2}}",
    re.IGNORECASE,
)
_FRACTION = re.compile(r"[+-]?(?:[0-9]+ +)?[0-9]+/[0-9]+")  # 1 1/2 is 1.5 in Excel
_DOT_THOUSANDS = re.compile(r"[+-]?[0-9]{1,3}(?:\.[0-9]{3}){2,}(?:,[0-9]+)?")  # 1.234.567
_CURRENCY = "$€£¥"
_DECORATION = "%,()'" + _CURRENCY
_ERROR_VALUES = frozenset(
    {
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#GETTING_DATA",
        "#SPILL!",
        "#CALC!",
        "#FIELD!",
        "#BLOCKED!",
        "#CONNECT!",
        "#BUSY!",
        "#UNKNOWN!",
    }
)
_COLUMN = re.compile(r"^[A-Za-z]{1,3}$")
_EXCEL_EPOCH = date(1899, 12, 30)
_TRUE, _FALSE = "TRUE", "FALSE"
_NAME_ERROR = "#NAME?"
_MAX_ROW = 1_048_576
_MAX_COL = 16_384


@dataclass(slots=True)
class CalcRequest:
    """Inputs for :func:`calc`."""

    workbook: Path
    policies: Path
    input_map: Path
    out_dir: Path
    outputs: list[str] = field(default_factory=list)
    variables: Path | None = None
    check_against: Path | None = None
    ir: Path | None = None
    as_of: date | None = None
    tolerance: float = 1e-9
    probe_cap: int = 50_000


@dataclass(slots=True, frozen=True)
class CalcResult:
    """The files a run wrote; ``rows_csv`` only with ``variables``."""

    results_csv: Path
    rows_csv: Path | None
    trust_report: Path
    paths_csv: Path
    run_json: Path


# -- input values --------------------------------------------------------------


def coerce_input(raw: str, stored: StoredValue) -> float | str | bool:
    """Convert a policy value as Excel converts what is typed into the target cell.

    The target's stored type decides:

    - a number cell takes a plain ASCII decimal (optional sign, digits, optional
      fraction and exponent) or an ISO ``YYYY-MM-DD`` date as its Excel 1900 serial;
    - a boolean cell takes ``TRUE``/``FALSE`` in any case;
    - a text cell takes the text as written, less one leading apostrophe (Excel's
      entry prefix);
    - a cell the workbook saved empty takes an ISO date as its serial, ``TRUE``/``FALSE``
      as a boolean, a plain decimal as a number, and any other text as text.

    Raises:
        CalcInputError: for a value that does not fit a number or boolean cell; on a
            number or empty cell, for a spelling whose meaning depends on the locale
            (``5%``, ``1,000``, ``$100``, ``(100)``, digits other than 0-9, a date not
            written ``YYYY-MM-DD``, a time), a leading apostrophe, an error value, a
            number outside Excel's range, and an ISO date before 1900-03-01.
    """
    text = raw.strip()
    if isinstance(stored, bool):
        upper = text.upper()
        if upper not in (_TRUE, _FALSE):
            raise CalcInputError(f"{raw!r} is not TRUE or FALSE")
        return upper == _TRUE
    if isinstance(stored, str):
        # Excel's entry prefix: '123 typed into a cell holds the text 123.
        return raw[1:] if raw.startswith("'") else raw
    if _ISO_DATE.fullmatch(text):
        return _iso_serial(raw, text)
    if stored is None and text.upper() in (_TRUE, _FALSE):
        return text.upper() == _TRUE
    if _PLAIN_NUMBER.fullmatch(text):
        return _excel_number(raw, text)
    spelling = _locale_spelling(text)
    if spelling is not None:
        raise CalcInputError(f"{raw!r} is {spelling}")
    if stored is None:
        return raw
    raise CalcInputError(f"{raw!r} is not a number")


def _iso_serial(raw: str, text: str) -> float:
    """An ISO date's Excel 1900 serial; dates before 1900-03-01 are refused (D4)."""
    year, month, day = (int(part) for part in text.split("-"))
    if (year, month, day) < (1900, 3, 1):
        raise CalcInputError(
            f"{raw!r} is before 1900-03-01: Excel's 1900 date system counts a "
            "29 February 1900 (the 1900 leap-year error), so its serials before March "
            "1900 are a day off the calendar, and it has no dates before 1900; give the "
            "serial number instead"
        )
    try:
        return float((date(year, month, day) - _EXCEL_EPOCH).days)
    except ValueError as exc:
        raise CalcInputError(f"{raw!r} is not a valid date") from exc


def _excel_number(raw: str, text: str) -> float:
    """A plain decimal as Excel holds it; refused outside the range Excel accepts."""
    value = float(text) + 0.0  # Excel has no negative zero
    if not math.isfinite(value) or abs(value) > _EXCEL_MAX:
        raise CalcInputError(f"{raw!r} is beyond Excel's largest number (9.99999999999999E+307)")
    mantissa = re.split(r"[eE]", text)[0]
    if (value == 0.0 and any(ch in "123456789" for ch in mantissa)) or (
        0.0 < abs(value) < _EXCEL_MIN
    ):
        raise CalcInputError(f"{raw!r} is below Excel's smallest number (2.2251E-308)")
    return value


def _locale_spelling(text: str) -> str | None:
    """Why Excel would type ``text`` into a number, date, time or error by rules calc
    does not apply (most of them the locale's); None when Excel keeps it as text."""
    if text.upper() in _ERROR_VALUES:
        return "an error value; calc cannot set one as an input"
    if text.startswith("'"):
        return (
            "text with a leading apostrophe, which Excel drops (it keeps the rest as "
            "text); give the value without it"
        )
    kind = _locale_kind(text)
    if kind is None:
        return None
    return (
        f"{kind}, which Excel reads by the locale's rules; give a plain decimal "
        "(0.05, 1000, -100) or an ISO date (YYYY-MM-DD)"
    )


def _locale_kind(text: str) -> str | None:
    if _TIME.fullmatch(text) or _DATE_TIME.fullmatch(text):
        return "a time or date-time"
    if _FRACTION.fullmatch(text) and " " in text.strip("+-"):
        return "a fraction"
    if _is_locale_date(text):
        return "a date not written YYYY-MM-DD"
    if _DOT_THOUSANDS.fullmatch(text):
        return "a number with dots between thousands"
    if any(ch.isdecimal() and not ("0" <= ch <= "9") for ch in text):
        if _PLAIN_NUMBER.fullmatch(_ascii_digits(text)):
            return "a number written with digits other than 0-9"
        return None
    bare = "".join(ch for ch in text if ch not in _DECORATION and not ch.isspace())
    if bare == text or not _PLAIN_NUMBER.fullmatch(bare):
        return None
    kinds = [
        ("%", "a percent"),
        (",", "a number with a comma (thousands or decimal)"),
        ("(", "a number in parentheses"),
    ]
    for mark, kind in kinds:
        if mark in text:
            return kind
    if any(ch in _CURRENCY for ch in text):
        return "a currency amount"
    return "a number with spaces or separators"


def _is_locale_date(text: str) -> bool:
    """True when some locale reads ``text`` as a date: digits joined by / - . whose
    parts can be a day (1-31), a month (1-12) and a year (1-2 digits, or 4 digits from
    1900) in day-month-year, month-day-year or year-month-day order, or two of them;
    or a month name with day and/or year numbers. ``123-45-6789`` is text everywhere."""
    if _MONTH_DATE.fullmatch(text):
        return any(ch.isdigit() for ch in text)
    m = _OTHER_DATE.fullmatch(text)
    if m is None:
        return False
    parts = [p for p in m.groups() if p is not None]

    def day(p: str) -> bool:
        return len(p) <= 2 and 1 <= int(p) <= 31

    def month(p: str) -> bool:
        return len(p) <= 2 and 1 <= int(p) <= 12

    def year(p: str) -> bool:
        return len(p) <= 2 or (len(p) == 4 and 1900 <= int(p) <= 9999)

    def long_year(p: str) -> bool:
        return len(p) == 4 and 1900 <= int(p) <= 9999

    if len(parts) == 3:
        a, b, c = parts
        return (
            (day(a) and month(b) and year(c))
            or (month(a) and day(b) and year(c))
            or (long_year(a) and month(b) and day(c))
        )
    a, b = parts
    return (
        (month(a) and day(b))
        or (day(a) and month(b))
        or (month(a) and long_year(b))
        or (long_year(a) and month(b))
    )


def _ascii_digits(text: str) -> str:
    return "".join(str(int(ch)) if ch.isdecimal() else ch for ch in text)


def _as_text(value: StoredValue) -> str:
    """Stored value as the text a policy file would hold (lossless for numbers)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return _TRUE if value else _FALSE
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else repr(value)
    return str(value)


_NUMBER_TEXT = re.compile(r"^[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$")


def _allowed(raw: str, value: float | str | bool | None, allowed: ListItems) -> bool:
    """Excel's list check on the value entered (``value`` is ``raw`` converted for its
    cell, None when it does not convert).

    A literal item is the list's text: it matches the entry as text ignoring case, or,
    when the entry is a number (an ISO date typed into a date cell is its serial), an
    item that reads as that number. An item read from a cell keeps the cell's type: a
    number matches number items only, text or a boolean non-number items ignoring case.
    """
    if allowed.literal:
        folded = raw.lower()
        for item in allowed.items:
            text = _as_text(item)
            if text.lower() == folded:
                return True
            if isinstance(value, float) and _NUMBER_TEXT.match(text) and float(text) == value:
                return True
        return False
    if isinstance(value, float):
        return any(isinstance(item, float) and item == value for item in allowed.items)
    text = (_as_text(value) if value is not None else raw).lower()
    return any(
        not isinstance(item, float) and _as_text(item).lower() == text for item in allowed.items
    )


# -- resolving targets ---------------------------------------------------------


def _single(graph: WorkbookGraph, target: str, what: str) -> CellRef:
    cells = graph.resolve_target(target)
    if len(cells) != 1:
        raise CalcInputError(f"{what} {target!r} must be a single cell")
    return cells[0]


def _input_cell(graph: WorkbookGraph, target: str, what: str) -> CellRef:
    cell = _single(graph, target, what)
    if graph.formula(cell) is not None:
        raise CalcInputError(f"{what} {target!r} maps onto formula cell {cell.a1()}")
    anchor = graph.array_anchor(cell)
    if anchor is not None:
        # Excel computes every cell of an array formula's range; none takes a constant.
        raise CalcInputError(
            f"{what} {target!r} maps onto formula cell {cell.a1()} "
            f"(in the array formula at {anchor.a1()})"
        )
    return cell


def _sheet(graph: WorkbookGraph, name: str, what: str) -> str:
    """The workbook's spelling of a sheet named in any case (Excel ignores case)."""
    sheet = graph.sheet(name)
    if sheet is None:
        raise CalcInputError(f"unknown sheet {name!r} in {what}")
    return sheet


def _column(letters: str, what: str) -> int:
    if not _COLUMN.match(letters):
        raise CalcInputError(f"{what} {letters!r} is not a column letter")
    return _col_to_number(letters)


def _output_cells(
    graph: WorkbookGraph, outputs: list[str], variables: list[VariableCells]
) -> list[CellRef]:
    """The cells to report.

    Raises:
        CalcInputError: naming every output target that does not resolve, or when
            nothing is asked for.
    """
    cells: dict[CellRef, None] = {}
    problems: list[str] = []
    for target in outputs:
        try:
            cells.update(dict.fromkeys(graph.resolve_target(target)))
        except CalcInputError as exc:
            problems.append(str(exc))
    if problems:
        raise CalcInputError("; ".join(problems))
    for var in variables:
        cells.update(dict.fromkeys(var.cells))
    if not cells:
        raise CalcInputError("nothing to calculate: give --outputs or --variables")
    return list(cells)


def _resolve_targets(
    graph: WorkbookGraph,
    specs: list[VariableSpec],
    outputs: list[str],
    fields: dict[str, str],
) -> tuple[list[VariableCells], list[CellRef], dict[str, CellRef]]:
    """Variables, output cells and mapped inputs, resolved together.

    Raises:
        CalcInputError: naming every bad target at once.
    """
    problems: list[str] = []
    variables: list[VariableCells] = []
    cells: list[CellRef] = []
    inputs: dict[str, CellRef] = {}
    try:
        variables = expand_variables(specs, graph)
    except CalcInputError as exc:
        problems.append(str(exc))
    try:
        cells = _output_cells(graph, outputs, variables)
    except CalcInputError as exc:
        # With no --outputs, "nothing to calculate" only echoes the variables' errors.
        if outputs or not problems:
            problems.append(str(exc))
    for name, target in fields.items():
        try:
            inputs[name] = _input_cell(graph, target, f"map field {name!r}")
        except CalcInputError as exc:
            problems.append(str(exc))
    problems.extend(_shared_cells(inputs, "map fields"))
    if problems:
        raise CalcInputError("; ".join(problems))
    return variables, cells, inputs


def _shared_cells(entries: dict[str, CellRef], what: str) -> list[str]:
    """One problem per cell that two entries resolve to: a cell holds one value, so
    one entry would be dropped without a word."""
    by_cell: dict[CellRef, list[str]] = {}
    for name, cell in entries.items():
        by_cell.setdefault(cell, []).append(name)
    return [
        f"{what} {' and '.join(repr(n) for n in sorted(names))} both resolve to {cell.a1()}"
        for cell, names in sorted(by_cell.items())
        if len(names) > 1
    ]


def _input_use(
    graph: WorkbookGraph, inputs: dict[str, CellRef], reach: frozenset[CellRef], id_field: str
) -> dict[str, dict[str, object]]:
    """Per map field, the cell it sets and whether a requested output reads it, for
    ``run.json``: True, False, or ``"unknown"`` when only a read the graph cannot place
    (through a dynamic reference, a name holding a formula, a 3-D reference) may reach
    it.

    Raises:
        CalcInputError: naming every field no formula can read: its value would
            change nothing, so each policy would get the stored inputs' answer as
            ``ok``. The id field is exempt (writing the id into a display cell is a
            documented use).
    """
    hidden: HiddenReads | None = None
    use: dict[str, dict[str, object]] = {}
    problems: list[str] = []
    for name, cell in inputs.items():
        if hidden is None:
            hidden = hidden_reads(graph)  # one workbook scan per run
        read = field_read(graph, cell, reach, hidden)
        if read is None and name != id_field:
            problems.append(
                f"map field {name!r} sets {cell.a1()}, which no formula reads, so its "
                "value would change nothing; map the field onto the input cell the "
                "formulas read"
            )
        use[name] = {
            "target": cell.a1(),
            "read_by_requested_outputs": False if read is None else read,
        }
    if problems:
        raise CalcInputError("; ".join(problems))
    return use


# -- reports -------------------------------------------------------------------


def _settings_record(settings: WorkbookSettings, refused: list[str]) -> dict[str, Any]:
    record: dict[str, Any] = asdict(settings)
    record["refused"] = refused
    return record


def _trust_report(
    *,
    graph: WorkbookGraph | None,
    trust: TrustResult | None,
    patterns: dict[CellRef, list[Reason]],
    settings: WorkbookSettings,
    refused: list[str],
    held: dict[CellRef, list[str]],
    check: dict[str, Any] | None,
    plan: ProbePlan | None = None,
) -> dict[str, Any]:
    failures = (
        sorted(
            trust.failures,
            key=lambda f: (f.reason.kind is not ReasonKind.silently_wrong, f.cell),
        )
        if trust is not None
        else []
    )
    return {
        "formula_cells": len(graph.formula_cells()) if graph is not None else None,
        "checked": trust.checked if trust is not None else None,
        "reproduced": trust.reproduced if trust is not None else None,
        "share": trust.share if trust is not None else None,
        "failures": [
            {"cell": f.cell.a1(), "kind": f.reason.kind.value, "detail": f.reason.detail}
            for f in failures
        ],
        "refused_patterns": [
            {"cell": cell.a1(), "detail": r.detail} for cell, rs in patterns.items() for r in rs
        ],
        "settings": _settings_record(settings, refused),
        "held_external_links": [{"cell": cell.a1(), "refs": refs} for cell, refs in held.items()],
        "check_against": check,
        "probes_unplaced": _unplaced_count(plan),
        "probes_over_cap": plan.over_cap if plan is not None else 0,
    }


def _unplaced_count(plan: ProbePlan | None) -> int:
    return sum(len(k) for k in plan.unplaced.values()) if plan is not None else 0


def _run_record(
    req: CalcRequest,
    sha: str,
    *,
    ironcalc: str | None,
    plan: ProbePlan | None,
    variables: dict[str, dict[str, object]],
    inputs: dict[str, dict[str, object]] | None = None,
) -> dict[str, Any]:
    from xl_marinade import __version__

    return {
        "ironcalc_version": ironcalc,
        "marinade_version": __version__,
        "workbook_sha256": sha,
        "tolerance": req.tolerance,
        "as_of": req.as_of.isoformat() if req.as_of else None,
        "probe_cap": req.probe_cap,
        "probes_unplaced": _unplaced_count(plan),
        "variables": variables,
        "inputs": inputs or {},
    }


# -- the run -------------------------------------------------------------------


@dataclass(slots=True)
class _Session:
    """Everything one run holds once the engine is loaded."""

    req: CalcRequest
    graph: WorkbookGraph
    engine: Engine
    plan: ProbePlan
    inputs: dict[str, CellRef]
    resets: list[CellRef]
    stored_name_errors: frozenset[CellRef] = frozenset()
    """Outputs that are ``#NAME?`` in the trust run (on the stored inputs)."""
    spills: list[ArrayRange] = field(default_factory=list)
    """Dynamic arrays, whose spill can grow past the range the workbook stored."""
    landing: LandingCheck | None = None
    """Roots for landed cells the plan did not probe."""
    static_up: frozenset[CellRef] = frozenset()
    """The outputs' upstream in the static graph."""

    def reset(self) -> None:
        """Put every mapped input (and the selector) back to its stored value."""
        for cell in self.resets:
            if not self.graph.has_stored(cell):
                self.engine.clear(cell)
                continue
            value = self.graph.stored(cell)
            if value is not None:
                self.engine.set_value(cell, value)

    def converted(self, cell: CellRef, raw: str) -> float | str | bool | None:
        """``raw`` as it would be set into ``cell``; None when it does not convert."""
        try:
            return coerce_input(raw, self.graph.stored(cell))
        except CalcInputError:
            return None

    def set_input(self, cell: CellRef, raw: str, what: str) -> None:
        try:
            value = coerce_input(raw, self.graph.stored(cell))
        except CalcInputError as exc:
            raise CalcInputError(f"{what} ({cell.a1()}): {exc}") from exc
        self.engine.set_value(cell, value)

    def recalc(self) -> ProbeReadings:
        self.engine.evaluate()
        return read_probes(self.engine, self.plan)


def _id_rows(graph: WorkbookGraph, sheet: str, column: int, first_row: int) -> dict[str, int]:
    """Id text -> row, from ``first_row`` down to the first id cell with no stored value.

    A formula id cell whose result is ``""`` has a stored value (it is not blank in
    Excel), so the walk goes on past it; ``""`` itself is not an id.
    """
    rows: dict[str, int] = {}
    row = first_row
    while graph.has_stored(CellRef(sheet, row, column)):
        pid = _as_text(graph.stored(CellRef(sheet, row, column))).strip()
        if pid:
            rows.setdefault(pid, row)
        row += 1
    return rows


@dataclass(slots=True)
class _Mismatches:
    """The input sets that did not reproduce one ``--check-against`` output.

    ``keys`` are the branch and lookup probes upstream of the output (``unplaced`` when
    one of them could not be placed). ``sets`` groups the mismatching sets by their path
    through those probes; per path it keeps the smallest reasons only, as many as a
    label can carry, so a policy compares once per distinct path and memory stays
    bounded however many sets mismatch.
    """

    keys: frozenset[PathKey]
    unplaced: bool
    sets: dict[PathSignature, list[Reason]] = field(default_factory=dict)

    def add(self, signature: PathSignature, reason: Reason) -> None:
        kept = self.sets.setdefault(signature, [])
        if reason in kept:
            return
        kept.append(reason)
        kept.sort(key=reason_key)
        del kept[MAX_REASONS_PER_KIND:]


def _json_value(value: EngineValue) -> float | str | bool | None:
    """A value as JSON can hold it: non-finite numbers become their ``repr``."""
    if isinstance(value, float) and not math.isfinite(value):
        return repr(value)
    return value


def _difference(stored: StoredValue, got: EngineValue) -> float | None:
    """``|stored - got|`` scaled as the tolerance is (relative, absolute below 1)."""
    numbers = [v for v in (stored, got) if isinstance(v, int | float) and not isinstance(v, bool)]
    if len(numbers) != 2:
        return None
    e, g = (float(v) for v in numbers)
    if not (math.isfinite(e) and math.isfinite(g)):
        return None
    return abs(e - g) / max(1.0, abs(e), abs(g))


def _check_against(
    s: _Session, check: CheckAgainst, evidence: ExercisedBuilder
) -> tuple[dict[CellRef, _Mismatches], dict[str, Any]]:
    """Recalculate every stored input set. A set that matched adds the paths it took to
    ``evidence``; per output, return the sets that did not with the path each took.

    Each set's readings are dropped once folded in, so memory does not grow with the
    number of sets (a mismatching set keeps only its path signature)."""
    graph = s.graph
    results_sheet = _sheet(graph, check.results_sheet, "results.sheet")
    results = _id_rows(
        graph, results_sheet, _column(check.id_column, "results.id_column"), check.first_row
    )
    resolved = {target: _single(graph, target, "check output") for target in check.outputs}
    shared = _shared_cells(resolved, "check outputs")
    if shared:
        raise CalcInputError("; ".join(shared))
    outputs = {
        resolved[target]: _column(col, f"results column for {target}")
        for target, col in check.outputs.items()
    }
    table_rows: dict[str, int] = {}
    table_cols: dict[CellRef, int] = {}
    table_sheet = ""
    if check.table is not None:
        t = check.table
        table_sheet = _sheet(graph, t.sheet, "inputs.table.sheet")
        table_rows = _id_rows(
            graph, table_sheet, _column(t.id_column, "table.id_column"), t.first_row
        )
        for name, letters in t.fields.items():
            if name not in s.inputs:
                raise CalcInputError(f"check table field {name!r} is not in the input map")
            table_cols[s.inputs[name]] = _column(letters, f"table column for {name}")
    selector = _input_cell(graph, check.selector, "selector") if check.selector else None

    # A set vouches only for the paths that fed the outputs it was compared on: through
    # the cells its own dynamic references landed on, not where the stored inputs did.
    compared = graph.upstream(outputs)
    scoped: dict[CellRef, _Mismatches] = {}
    mismatches: list[dict[str, Any]] = []
    for pid, row in results.items():
        s.reset()
        if selector is not None:
            s.set_input(selector, pid, f"selector for input set {pid!r}")
        else:
            trow = table_rows.get(pid)
            if trow is None:
                raise CalcInputError(f"input set {pid!r} has no row in the check input table")
            for cell, col in table_cols.items():
                src = CellRef(table_sheet, trow, col)
                if not graph.has_stored(src):
                    raise CalcInputError(f"input set {pid!r}: no value in {src.a1()}")
                # The table's value as stored (a TRUE stays a boolean): it is what
                # was in the cell when Excel computed that input set's results.
                stored_input = graph.stored(src)
                if stored_input is not None:
                    s.engine.set_value(cell, stored_input)
        try:
            reading = s.recalc()
        except CalcError as exc:
            raise CalcError(f"input set {pid!r}: {exc}") from exc
        matched = True
        for cell, col in outputs.items():
            stored = graph.stored(CellRef(results_sheet, row, col))
            got = s.engine.value(cell)
            if same(stored, got, s.req.tolerance):
                continue
            detail = f"input set {pid!r}: stored {stored!r}, engine {got!r}"
            entry = scoped.get(cell)
            if entry is None:
                entry = _Mismatches(*path_keys(s.plan, graph.upstream([cell])))
                scoped[cell] = entry
            entry.add(
                path_signature(reading, entry.keys, entry.unplaced),
                Reason(ReasonKind.check_against_mismatch, cell, detail),
            )
            mismatches.append(
                {
                    "input_set": pid,
                    "cell": cell.a1(),
                    "detail": detail,
                    "stored": _json_value(stored),
                    "computed": _json_value(got),
                    "relative_difference": _difference(stored, got),
                }
            )
            matched = False
        if matched:
            # Only an input set the engine reproduced vouches for the paths it took,
            # through the cells its dynamic references landed on too.
            evidence.add(reading, only=_fed(graph, outputs, compared, reading.dynamic_refs))
    return scoped, {"input_sets": len(results), "mismatches": mismatches}


def _mismatch_roots(
    roots: dict[CellRef, list[Reason]],
    readings: ProbeReadings,
    mismatches: dict[CellRef, _Mismatches],
) -> None:
    """Add a ``check_against_mismatch`` on each output a mismatching input set reached by
    the path this policy took (or may have taken) too."""
    for cell, entry in mismatches.items():
        mine = path_signature(readings, entry.keys, entry.unplaced)
        for theirs, reasons in entry.sets.items():
            if not same_path(theirs, mine):
                continue
            bucket = roots.setdefault(cell, [])
            seen = set(bucket)
            bucket.extend(r for r in reasons if r not in seen)


def _fed(
    graph: WorkbookGraph,
    roots: Iterable[CellRef],
    static: frozenset[CellRef] | set[CellRef],
    dynamic: dict[CellRef, list[CellRef]],
) -> frozenset[CellRef] | set[CellRef]:
    """The cells that fed ``roots`` in one recalculation: their upstream, where a cell
    with a dynamic reference contributes what it read apart from where its references
    land, plus the cells they landed on this time (``dynamic``). Not the reference an
    ``OFFSET`` starts from, nor where it landed with the stored inputs, unless this
    recalculation landed there. ``static`` is the roots' static upstream, returned as
    it is when nothing in it is dynamic."""
    if not any(graph.is_dynamic(c) for c in static) and not any(c in static for c in dynamic):
        return static
    seen: set[CellRef] = set()
    stack = list(roots)
    while stack:
        cur = stack.pop()
        if cur in seen:
            continue
        seen.add(cur)
        stack.extend(graph.read_precedents(cur))
        stack.extend(dynamic.get(cur, ()))
    return seen


def _paths(
    pid: str,
    plan: ProbePlan,
    readings: ProbeReadings,
    exercised: Exercised,
    reached: Container[CellRef],
) -> list[PathRecord]:
    """The branch and lookup edges this policy took in cells that fed its outputs."""
    records: list[PathRecord] = []
    for probe in plan.probes:
        if probe.cell not in reached:
            continue
        key = (probe.cell, probe.ordinal)
        if probe.kind is ProbeKind.branch:
            outcome = readings.branch.get(key)
            seen = exercised.branches
        elif probe.kind is ProbeKind.lookup:
            outcome = readings.lookup_edge.get(key)
            seen = exercised.lookup_edges
        else:
            continue
        if outcome is None:
            continue
        hit = (probe.family, probe.ordinal, outcome) in seen
        records.append(
            PathRecord(pid, probe.cell, probe.family, probe.kind.value, outcome, hit, probe.ordinal)
        )
    return records


def _spill_growth(s: _Session) -> dict[CellRef, list[CellRef]]:
    """Cells a dynamic array spilled into, for this policy, beyond the range the workbook
    stored, each mapped to its anchor: they read the anchor as its stored cells do.

    A spill grows right and down from its anchor, so a larger one fills the cell below
    the stored range in the anchor's column, or right of it in the anchor's row. A cell
    the workbook holds something in, or one calc writes an input into, holds a value of
    its own: it cannot be part of the spill (Excel would show ``#SPILL!``).
    """
    grown: dict[CellRef, list[CellRef]] = {}
    written = set(s.resets)
    for array in s.spills:
        anchor = array.anchor
        sheet, top, left = anchor

        def filled(cell: CellRef) -> bool:
            if cell in written or s.graph.holds(cell):
                return False
            return s.engine.value(cell) is not None

        rows = array.last_row
        while rows < _MAX_ROW and filled(CellRef(sheet, rows + 1, left)):
            rows += 1
        cols = array.last_col
        while cols < _MAX_COL and filled(CellRef(sheet, top, cols + 1)):
            cols += 1
        for r in range(top, rows + 1):
            for c in range(left, cols + 1):
                if r > array.last_row or c > array.last_col:
                    grown.setdefault(CellRef(sheet, r, c), []).append(anchor)
    return grown


def _policy_loop(
    s: _Session,
    policies: list[Policy],
    outputs: list[CellRef],
    variables: list[VariableCells],
    exercised: Exercised,
    static: StaticLabels,
    mismatches: dict[CellRef, _Mismatches],
    writer: PolicyWriter,
) -> None:
    """Recalculate each policy and hand its records to ``writer`` before the next one."""
    allowed = {name: s.graph.allowed_values(cell) for name, cell in s.inputs.items()}
    for policy in policies:
        s.reset()
        outside: list[Reason] = []
        for name, cell in s.inputs.items():
            raw = policy.values[name]
            options = allowed[name]
            what = f"policy {policy.id!r} field {name!r}"
            if options is None or _allowed(raw, s.converted(cell, raw), options):
                s.set_input(cell, raw, what)
                continue
            detail = f"{name}={raw!r} not in {[_as_text(o) for o in options.items]!r}"
            outside.append(Reason(ReasonKind.input_outside_allowed_values, cell, detail))
            try:
                s.set_input(cell, raw, what)
            except CalcInputError:
                # Refused anyway; the cell keeps its stored value for this policy.
                continue
        try:
            readings = s.recalc()
        except CalcError as exc:
            raise CalcError(f"policy {policy.id!r}: {exc}") from exc
        roots = policy_roots(s.plan, readings, exercised, include_unplaced=False)
        _mismatch_roots(roots, readings, mismatches)
        values = {cell: s.engine.value(cell) for cell in outputs}
        for cell, value in values.items():
            code = engine_error_code(s.engine, cell, value)
            if code is not None:
                reason = Reason(ReasonKind.engine_error, cell, engine_error_detail(code))
            elif _new_name_error(s, cell, value):
                reason = Reason(ReasonKind.engine_error, cell, NAME_OFF_PATH)
            else:
                continue
            bucket = roots.setdefault(cell, [])
            if reason not in bucket:
                bucket.append(reason)
        dynamic = dict(readings.dynamic_refs)
        for cell, anchors in _spill_growth(s).items():
            dynamic[cell] = sorted({*dynamic.get(cell, ()), *anchors})
        if s.landing is not None:
            for cell, reasons in s.landing.roots(dynamic).items():
                bucket = roots.setdefault(cell, [])
                bucket.extend(r for r in reasons if r not in bucket)
        assigned = assign(s.graph, outputs, roots, dynamic, static=static)
        computed: dict[CellRef, tuple[EngineValue, Status, list[Reason]]] = {}
        results: list[OutputValue] = []
        for cell in outputs:
            status, reasons = assigned[cell]
            if outside:
                status, reasons = Status.REFUSED, sorted([*reasons, *outside], key=reason_key)
            value = values[cell]
            computed[cell] = (value, status, reasons)
            results.append(OutputValue(policy.id, cell, value, status, reasons))
        writer.write_policy(
            results,
            _paths(
                policy.id, s.plan, readings, exercised, _fed(s.graph, outputs, s.static_up, dynamic)
            ),
            rows_for(policy.id, variables, computed.__getitem__),
        )


def _new_name_error(s: _Session, cell: CellRef, value: EngineValue) -> bool:
    """True when ``cell`` is ``#NAME?`` for this policy but was not in the trust run.

    Excel's ``#NAME?`` comes from the formula's text, not its inputs, so one that only
    a new policy reaches is a function the engine lacks on an arm the trust run never
    took: the engine has nothing to offer for it.
    """
    if value != _NAME_ERROR or cell in s.stored_name_errors:
        return False
    return s.engine.is_error(cell)


def _refuse(req: CalcRequest, settings: WorkbookSettings, reasons: list[str], sha: str) -> NoReturn:
    write_trust_report(
        req.out_dir / "trust_report.json",
        _trust_report(
            graph=None,
            trust=None,
            patterns={},
            settings=settings,
            refused=reasons,
            held={},
            check=None,
        ),
    )
    write_run(
        req.out_dir / "run.json",
        _run_record(req, sha, ironcalc=engine_version(), plan=None, variables={}),
    )
    raise WorkbookRefused(f"workbook refused: {', '.join(reasons)}", reasons=reasons)


def _sha256(path: Path) -> str:
    try:
        return hashlib.sha256(path.read_bytes()).hexdigest()
    except OSError as exc:
        raise CalcInputError(f"cannot read {path}: {exc}") from exc


def calc(req: CalcRequest) -> CalcResult:
    """Recalculate every policy and write ``results.csv``, ``paths.csv``,
    ``trust_report.json``, ``run.json`` and (with ``variables``) ``rows.csv``.

    Raises:
        WorkbookRefused: when a workbook setting refuses the whole workbook; only
            ``trust_report.json`` and ``run.json`` are written.
        CalcInputError: on bad input files, unknown targets, or an input mapped onto
            a formula cell.
        CalcUnavailable: if ``ironcalc`` is not installed.
    """
    if not (math.isfinite(req.tolerance) and req.tolerance >= 0):
        raise CalcInputError(f"tolerance must be a finite number >= 0, got {req.tolerance}")
    if req.probe_cap < 0:
        raise CalcInputError(f"probe_cap must be >= 0, got {req.probe_cap}")
    guard_out_dir(
        req.out_dir,
        CALC_OUTPUTS,
        [req.workbook, req.policies, req.input_map, req.variables, req.check_against, req.ir],
    )
    clear_outputs(req.out_dir, CALC_OUTPUTS)
    sha = _sha256(req.workbook)
    settings = read_settings(req.workbook)
    refused = refusal_reasons(settings)
    if refused:
        _refuse(req, settings, refused, sha)

    input_map = load_input_map(req.input_map)
    policies = load_policies(req.policies, input_map)
    specs: list[VariableSpec] = load_variables(req.variables) if req.variables else []
    check = load_check_against(req.check_against) if req.check_against else None
    targets = parse_outputs(req.outputs)

    out = req.out_dir
    rows_csv = out / "rows.csv" if req.variables is not None else None
    with tempfile.TemporaryDirectory(prefix="marinade-calc-") as tmp:
        work = Path(tmp)
        graph = open_graph(req.workbook, req.ir, work / "ir")
        prepared = prepare_engine_copy(req.workbook, work / "engine.xlsx", as_of=req.as_of)
        graph.add_arrays(prepared.arrays)
        variables, outputs, inputs = _resolve_targets(graph, specs, targets, input_map.fields)
        static_up = frozenset(graph.upstream(outputs))
        input_use = _input_use(graph, inputs, static_up, input_map.id_field or "id")
        patterns = _merge_patterns(
            _with_unheld_externals(
                graph,
                prepared,
                scan_patterns(graph, as_of=req.as_of, data_table_cells=prepared.data_table_cells),
            ),
            input_dependent_lookups(graph, set(inputs.values())),
            _relative_lists(graph, inputs),
        )
        with quiet_native_output():
            engine = Engine.load(prepared.path, req.workbook)
        # Refused cells refuse their dependants anyway: probing them would only use up
        # the cap (and their readings vouch for nothing).
        upstream = [c for c in static_up if graph.formula(c) is not None and c not in patterns]
        moving = set(inputs.values())
        if check is not None and check.selector is not None:
            moving.add(_input_cell(graph, check.selector, "selector"))
        region = _landing_region(graph, set(upstream), patterns)
        plan = plan_probes(
            graph, upstream, cap=req.probe_cap, as_of=req.as_of, inputs=moving, extra=region
        )
        covered = frozenset({*upstream, *region, *patterns})
        place_probes(engine, plan)
        with quiet_native_output():
            engine.evaluate()
        # A refused array formula refuses every cell of its range: none is checked.
        refused_cells = set(patterns) | {m for a in patterns for m in graph.array_members(a)}
        trust = run_trust(
            engine,
            graph,
            tolerance=req.tolerance,
            skip=refused_cells | set(prepared.data_table_cells),
        )
        failing = {f.cell for f in trust.failures}
        evidence = ExercisedBuilder(plan, exclude=failing | set(patterns))
        evidence.add(read_probes(engine, plan))

        resets = list(dict.fromkeys(inputs.values()))
        if check is not None and check.selector is not None:
            resets.append(_input_cell(graph, check.selector, "selector"))
        stored_names = frozenset(c for c in outputs if engine.value(c) == _NAME_ERROR)
        session = _Session(
            req,
            graph,
            engine,
            plan,
            inputs,
            list(dict.fromkeys(resets)),
            stored_names,
            [a for a in graph.arrays() if a.dynamic],
            LandingCheck(graph, covered, lambda cell: probe_kinds(graph, cell)),
            static_up,
        )
        mismatches: dict[CellRef, _Mismatches] = {}
        check_record: dict[str, Any] | None = None
        if check is not None:
            mismatches, check_record = _check_against(session, check, evidence)
        exercised = evidence.exercised
        static = propagate_static(graph, _static_roots(trust, patterns, plan, prepared))
        with PolicyWriter(out / "results.csv", out / "paths.csv", rows_csv) as writer:
            _policy_loop(
                session, policies, outputs, variables, exercised, static, mismatches, writer
            )

    write_trust_report(
        out / "trust_report.json",
        _trust_report(
            graph=graph,
            trust=trust,
            patterns=patterns,
            settings=settings,
            refused=[],
            held=prepared.held_externals,
            check=check_record,
            plan=plan,
        ),
    )
    write_run(
        out / "run.json",
        _run_record(
            req,
            sha,
            ironcalc=engine.version,
            plan=plan,
            variables=variables_record(specs, variables),
            inputs=input_use,
        ),
    )
    return CalcResult(
        out / "results.csv",
        rows_csv,
        out / "trust_report.json",
        out / "paths.csv",
        out / "run.json",
    )


def _relative_lists(
    graph: WorkbookGraph, inputs: dict[str, CellRef]
) -> dict[CellRef, list[Reason]]:
    """A ``refused_pattern`` on each mapped input whose dropdown list is a name with a
    relative reference: Excel's list depends on the cell validated, so it is unknown."""
    out: dict[CellRef, list[Reason]] = {}
    for cell in sorted(set(inputs.values())):
        found = graph.list_source_name(cell)
        if found is not None and found.relative:
            detail = f"{RELATIVE_NAME}: {found.name}"
            out[cell] = [Reason(ReasonKind.refused_pattern, cell, detail)]
    return out


def _landing_region(
    graph: WorkbookGraph, planned: set[CellRef], patterns: dict[CellRef, list[Reason]]
) -> list[CellRef]:
    """Formula cells to probe because a dynamic reference may land on them:
    every formula cell on a sheet an ``OFFSET``, a built range or an ``INDIRECT`` with a
    literal sheet reaches from the planned cells, and everything upstream of those, to
    a fixed point. Probing them lets the trust run and ``--check-against`` exercise
    their branches and lookups, as for any cell upstream of an output."""
    covered = set(planned)
    region: list[CellRef] = []
    done: set[str] = set()
    queue = sorted(planned)
    while queue:
        cell = queue.pop()
        for sheet in sorted(reach_sheets(graph, cell) - done):
            done.add(sheet)
            for found in sorted(graph.upstream(graph.formula_cells_on(sheet))):
                if found in covered or found in patterns or graph.formula(found) is None:
                    continue
                covered.add(found)
                region.append(found)
                queue.append(found)
    return region


def _merge_patterns(
    *sources: dict[CellRef, list[Reason]],
) -> dict[CellRef, list[Reason]]:
    out: dict[CellRef, list[Reason]] = {}
    for source in sources:
        for cell, reasons in source.items():
            bucket = out.setdefault(cell, [])
            bucket.extend(r for r in reasons if r not in bucket)
    return {cell: sorted(rs, key=reason_key) for cell, rs in sorted(out.items())}


def _with_unheld_externals(
    graph: WorkbookGraph, prepared: PreparedCopy, patterns: dict[CellRef, list[Reason]]
) -> dict[CellRef, list[Reason]]:
    """Add a ``refused_pattern`` for every formula cell whose engine copy still holds an
    external reference (a range, or one with no cached value): IronCalc errors on it, and
    off the stored path an ``IFERROR`` could turn that into a plausible value."""
    out = {cell: list(reasons) for cell, reasons in patterns.items()}
    for cell in graph.formula_cells():
        text = graph.formula(cell) or ""
        if count_external_refs(text) <= len(prepared.held_externals.get(cell, [])):
            continue
        reason = Reason(ReasonKind.refused_pattern, cell, NOT_HELD)
        bucket = out.setdefault(cell, [])
        if reason not in bucket:
            bucket.append(reason)
    return {cell: sorted(rs, key=reason_key) for cell, rs in sorted(out.items())}


def _static_roots(
    trust: TrustResult,
    patterns: dict[CellRef, list[Reason]],
    plan: ProbePlan,
    prepared: PreparedCopy,
) -> dict[CellRef, list[Reason]]:
    """Roots that hold for every policy: trust failures, refused patterns, unplaced
    probes and held external links. A ``--check-against`` mismatch is not one: it holds
    only for policies on the mismatching set's path (:func:`_mismatch_roots`)."""
    roots: dict[CellRef, list[Reason]] = {}

    def add(cell: CellRef, reason: Reason) -> None:
        bucket = roots.setdefault(cell, [])
        if reason not in bucket:
            bucket.append(reason)

    for failure in trust.failures:
        add(failure.cell, failure.reason)
    for source in (patterns, unplaced_roots(plan)):
        for cell, reasons in source.items():
            for reason in reasons:
                add(cell, reason)
    for cell, refs in prepared.held_externals.items():
        add(cell, Reason(ReasonKind.external_link_held, cell, ", ".join(refs)))
    return {cell: sorted(rs, key=reason_key) for cell, rs in sorted(roots.items())}
