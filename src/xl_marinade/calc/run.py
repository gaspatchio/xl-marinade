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
from xl_marinade.calc.graph import StoredValue, WorkbookGraph, open_graph
from xl_marinade.calc.models import (
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
    write_run,
    write_trust_report,
)
from xl_marinade.calc.patterns import NOT_HELD, input_dependent_lookups, scan_patterns
from xl_marinade.calc.prepare import PreparedCopy, count_external_refs, prepare_engine_copy
from xl_marinade.calc.probes import (
    ProbeKind,
    ProbePlan,
    ProbeReadings,
    engine_error_code,
    place_probes,
    plan_probes,
    read_probes,
)
from xl_marinade.calc.rows import VariableCells, expand_variables, rows_for, variables_record
from xl_marinade.calc.settings import read_settings, refusal_reasons
from xl_marinade.calc.status import (
    MAX_REASONS_PER_KIND,
    NAME_OFF_PATH,
    Exercised,
    ExercisedBuilder,
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
from xl_marinade.errors import CalcInputError, WorkbookRefused

_ISO_DATE = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}$")
_COLUMN = re.compile(r"^[A-Za-z]{1,3}$")
_EXCEL_EPOCH = date(1899, 12, 30)
_TRUE, _FALSE = "TRUE", "FALSE"
_NAME_ERROR = "#NAME?"


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
    """Convert a policy value to the type of the target cell's stored value.

    A number cell takes a number, or an ISO ``YYYY-MM-DD`` date as its Excel 1900
    serial; a boolean cell takes ``TRUE``/``FALSE`` in any case; any other cell takes
    text. An empty target cell takes a number when the text is one, else text.

    Raises:
        CalcInputError: if the value does not fit a number or boolean cell.
    """
    text = raw.strip()
    if isinstance(stored, bool):
        upper = text.upper()
        if upper not in (_TRUE, _FALSE):
            raise CalcInputError(f"{raw!r} is not TRUE or FALSE")
        return upper == _TRUE
    if isinstance(stored, float):
        if _ISO_DATE.match(text):
            try:
                return float((date.fromisoformat(text) - _EXCEL_EPOCH).days)
            except ValueError as exc:
                raise CalcInputError(f"{raw!r} is not a valid date") from exc
        number = _number(text)
        if number is None:
            raise CalcInputError(f"{raw!r} is not a number")
        return number
    if stored is None:
        number = _number(text)
        return number if number is not None else raw
    return raw


def _number(text: str) -> float | None:
    try:
        value = float(text)
    except ValueError:
        return None
    return value if math.isfinite(value) else None


def _as_text(value: StoredValue) -> str:
    """Stored value as the text a policy file would hold (lossless for numbers)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return _TRUE if value else _FALSE
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else repr(value)
    return str(value)


def _allowed(raw: str, value: float | str | bool | None, allowed: list[str]) -> bool:
    """Excel's list check: the entry matches an item ignoring case, or by number.

    ``value`` is the entry converted for its cell (None if it does not convert), so an
    ISO date typed into a date cell matches a list of date serials.
    """
    folded = raw.casefold()
    if any(folded == item.casefold() for item in allowed):
        return True
    numbers = {n for item in allowed if (n := _number(item)) is not None}
    converted = value if isinstance(value, float) else None
    return any(n is not None and n in numbers for n in (_number(raw), converted))


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
    cells: dict[CellRef, None] = {}
    for target in outputs:
        cells.update(dict.fromkeys(graph.resolve_target(target)))
    for var in variables:
        cells.update(dict.fromkeys(var.cells))
    if not cells:
        raise CalcInputError("nothing to calculate: give --outputs or --variables")
    return list(cells)


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
    }


def _run_record(
    req: CalcRequest,
    sha: str,
    *,
    ironcalc: str | None,
    plan: ProbePlan | None,
    variables: dict[str, dict[str, object]],
) -> dict[str, Any]:
    from xl_marinade import __version__

    return {
        "ironcalc_version": ironcalc,
        "marinade_version": __version__,
        "workbook_sha256": sha,
        "tolerance": req.tolerance,
        "as_of": req.as_of.isoformat() if req.as_of else None,
        "probe_cap": req.probe_cap,
        "probes_unplaced": sum(len(k) for k in plan.unplaced.values()) if plan else 0,
        "variables": variables,
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
    """Id text -> row, from ``first_row`` down to the first blank id cell."""
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
    outputs = {
        _single(graph, target, "check output"): _column(col, f"results column for {target}")
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

    # A set vouches only for the paths that fed the outputs it was compared on.
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
                s.set_input(cell, _as_text(graph.stored(src)), f"input set {pid!r}")
        reading = s.recalc()
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
            # Only an input set the engine reproduced vouches for the paths it took.
            evidence.add(reading, only=compared)
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


def _paths(
    pid: str, plan: ProbePlan, readings: ProbeReadings, exercised: Exercised
) -> list[PathRecord]:
    records: list[PathRecord] = []
    for probe in plan.probes:
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
            detail = f"{name}={raw!r} not in {options!r}"
            outside.append(Reason(ReasonKind.input_outside_allowed_values, cell, detail))
            try:
                s.set_input(cell, raw, what)
            except CalcInputError:
                # Refused anyway; the cell keeps its stored value for this policy.
                continue
        readings = s.recalc()
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
        assigned = assign(s.graph, outputs, roots, readings.dynamic_refs, static=static)
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
            _paths(policy.id, s.plan, readings, exercised),
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
        variables = expand_variables(specs, graph)
        outputs = _output_cells(graph, targets, variables)
        inputs = {
            name: _input_cell(graph, target, f"map field {name!r}")
            for name, target in input_map.fields.items()
        }
        prepared = prepare_engine_copy(req.workbook, work / "engine.xlsx", as_of=req.as_of)
        patterns = _merge_patterns(
            _with_unheld_externals(
                graph,
                prepared,
                scan_patterns(graph, as_of=req.as_of, data_table_cells=prepared.data_table_cells),
            ),
            input_dependent_lookups(graph, set(inputs.values())),
        )
        with quiet_native_output():
            engine = Engine.load(prepared.path)
        # Refused cells refuse their dependants anyway: probing them would only use up
        # the cap (and their readings vouch for nothing).
        upstream = [
            c for c in graph.upstream(outputs) if graph.formula(c) is not None and c not in patterns
        ]
        plan = plan_probes(graph, upstream, cap=req.probe_cap)
        place_probes(engine, plan)
        with quiet_native_output():
            engine.evaluate()
        trust = run_trust(
            engine,
            graph,
            tolerance=req.tolerance,
            skip=set(patterns) | set(prepared.data_table_cells),
        )
        failing = {f.cell for f in trust.failures}
        evidence = ExercisedBuilder(plan, exclude=failing | set(patterns))
        evidence.add(read_probes(engine, plan))

        resets = list(dict.fromkeys(inputs.values()))
        if check is not None and check.selector is not None:
            resets.append(_input_cell(graph, check.selector, "selector"))
        stored_names = frozenset(c for c in outputs if engine.value(c) == _NAME_ERROR)
        session = _Session(
            req, graph, engine, plan, inputs, list(dict.fromkeys(resets)), stored_names
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
        ),
    )
    return CalcResult(
        out / "results.csv",
        rows_csv,
        out / "trust_report.json",
        out / "paths.csv",
        out / "run.json",
    )


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
