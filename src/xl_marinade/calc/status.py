"""Status assignment: turn root causes into a status and reasons for every output.

A *root* is a cell carrying reasons of its own: a trust failure, a refused pattern, an
external link held at its cached value, a probe that could not be placed, or a probe
reading for the current policy (a branch or lookup edge the trust run never exercised,
an input close to a rounding flip or a comparison). An output inherits the reasons of
every root that reaches it forward through the graph, extended per policy by the
dynamic references (``OFFSET`` / ``INDIRECT``) actually used: a cell depends on the
cells its dynamic reference landed on.

Reasons propagate as labels: per cell, per reason kind, the (at most) five reasons with
the smallest roots. Taking the five smallest is exact under union, so propagation to a
fixed point gives the same result as walking each root's closure, in any visiting
order. Roots that do not change between policies can be propagated once per run
(:func:`propagate_static`) and merged into every policy's :func:`assign`.
"""

from __future__ import annotations

from collections import deque
from collections.abc import Container, Iterable
from dataclasses import dataclass

from xl_marinade.calc.graph import WorkbookGraph
from xl_marinade.calc.models import SEVERITY, CellRef, Reason, ReasonKind, Status
from xl_marinade.calc.probes import ProbeKind, ProbePlan, ProbeReadings

MAX_REASONS_PER_KIND = 5
DYNAMIC_NOT_CHECKED = "dynamic reference not checked"
LOOKUP_TYPE = "lookup value type differs from keys"

_NO_ROOT = CellRef("", 0, 0)
_RANK: dict[Status, int] = {
    Status.REFUSED: 3,
    Status.UNCERTAIN: 2,
    Status.UNVERIFIED: 1,
    Status.OK: 0,
}
_UNPLACED: dict[ProbeKind, tuple[ReasonKind, str]] = {
    ProbeKind.branch: (ReasonKind.not_checked_path, "branch not checked"),
    ProbeKind.lookup: (ReasonKind.not_checked_path, "lookup position not checked"),
    ProbeKind.rounding: (ReasonKind.not_checked_value, "rounding not checked"),
    ProbeKind.comparison: (ReasonKind.not_checked_value, "comparison not checked"),
    ProbeKind.dynamic_ref: (ReasonKind.refused_pattern, DYNAMIC_NOT_CHECKED),
}

Label = dict[ReasonKind, tuple[Reason, ...]]
"""Per reason kind, the smallest reasons (by root, then detail) reaching a cell."""


def reason_key(r: Reason) -> tuple[ReasonKind, CellRef, str]:
    """Sort key for reasons; ``root`` may be None."""
    return (r.kind, r.root or _NO_ROOT, r.detail)


# -- exercised outcomes ------------------------------------------------------


@dataclass(slots=True)
class Exercised:
    """Outcomes the trust run saw, keyed ``(family, ordinal, outcome)``."""

    branches: set[tuple[str, int, str]]
    lookup_edges: set[tuple[str, int, str]]


def _families(plan: ProbePlan, kind: ProbeKind) -> dict[tuple[CellRef, int], tuple[str, str]]:
    """``(cell, ordinal)`` -> ``(family, function)`` for the plan's probes of one kind."""
    return {
        (p.cell, p.ordinal): (p.family, p.meta.get("function", ""))
        for p in plan.probes
        if p.kind is kind
    }


def _family_of(table: dict[tuple[CellRef, int], tuple[str, str]], key: tuple[CellRef, int]) -> str:
    # A reading outside the plan counts on its own, never borrowing a family.
    entry = table.get(key)
    return entry[0] if entry is not None else f"cell:{key[0].a1()}"


class ExercisedBuilder:
    """Folds trust readings into an :class:`Exercised` one at a time, so a run keeps
    the outcomes seen, not the readings (see :func:`exercised_from` for the rules)."""

    def __init__(self, plan: ProbePlan, *, exclude: Iterable[CellRef] = ()) -> None:
        self._skip = set(exclude)
        self._branch_fam = _families(plan, ProbeKind.branch)
        self._lookup_fam = _families(plan, ProbeKind.lookup)
        self.exercised = Exercised(set(), set())

    def add(self, reading: ProbeReadings, *, only: Container[CellRef] | None = None) -> None:
        """Count ``reading``'s outcomes; with ``only``, just those of probes in those cells."""
        out = self.exercised
        for key, outcome in reading.branch.items():
            if not self._counts(reading, ProbeKind.branch, key, only):
                continue
            out.branches.add((_family_of(self._branch_fam, key), key[1], outcome))
        for key, edge in reading.lookup_edge.items():
            if not self._counts(reading, ProbeKind.lookup, key, only):
                continue
            out.lookup_edges.add((_family_of(self._lookup_fam, key), key[1], edge))

    def _counts(
        self,
        reading: ProbeReadings,
        kind: ProbeKind,
        key: tuple[CellRef, int],
        only: Container[CellRef] | None,
    ) -> bool:
        if key[0] in self._skip or (kind, *key) in reading.unsure:
            return False
        return only is None or key[0] in only


def exercised_from(
    plan: ProbePlan,
    readings: list[ProbeReadings],
    *,
    exclude: Iterable[CellRef] = (),
) -> Exercised:
    """Every branch outcome and lookup edge seen in the trust readings, per formula family.

    ``interior`` counts for the whole family whichever row it landed on; ``first``,
    ``last`` and ``not_found`` count only when the family hit that edge itself. Only a
    cell that reproduced vouches for its family: readings of cells in ``exclude``
    (trust failures, refused patterns) and unsure readings (an enclosing branch had no
    reading) are not evidence.
    """
    builder = ExercisedBuilder(plan, exclude=exclude)
    for reading in readings:
        builder.add(reading)
    return builder.exercised


# -- path signatures ---------------------------------------------------------

PathKey = tuple[CellRef, int, ProbeKind]
"""A branch or lookup-edge probe: ``(cell, ordinal, kind)``."""

_PATH_KINDS = (ProbeKind.branch, ProbeKind.lookup)


@dataclass(slots=True, frozen=True)
class PathSignature:
    """The path one recalculation took through a set of path probes.

    ``outcomes`` holds every live reading of those probes (a probe in an arm the cell
    did not evaluate has none). ``unsure`` is set when the path cannot be told: one of
    the probes was unplaced, or its reading is unsure.
    """

    outcomes: frozenset[tuple[PathKey, str]]
    unsure: bool


def path_keys(plan: ProbePlan, cells: Iterable[CellRef]) -> tuple[frozenset[PathKey], bool]:
    """The branch and lookup probes of ``cells``, and whether any of them is unplaced."""
    scope = set(cells)
    keys = frozenset(
        (p.cell, p.ordinal, p.kind)
        for p in plan.probes
        if p.kind in _PATH_KINDS and p.cell in scope
    )
    unplaced = any(
        kind in _PATH_KINDS
        for cell, kinds in plan.unplaced.items()
        if cell in scope
        for kind in kinds
    )
    return keys, unplaced


def path_signature(
    readings: ProbeReadings, keys: frozenset[PathKey], unplaced: bool
) -> PathSignature:
    """The path ``readings`` took through ``keys`` (from :func:`path_keys`).

    Uses the same rules as :func:`exercised_from`: a skipped probe has no outcome, and
    an unsure reading leaves the path unknown.
    """
    outcomes: set[tuple[PathKey, str]] = set()
    unsure = unplaced
    for kind, table in (
        (ProbeKind.branch, readings.branch),
        (ProbeKind.lookup, readings.lookup_edge),
    ):
        for (cell, ordinal), outcome in table.items():
            key = (cell, ordinal, kind)
            if key not in keys:
                continue
            outcomes.add((key, outcome))
            if (kind, cell, ordinal) in readings.unsure:
                unsure = True
    return PathSignature(frozenset(outcomes), unsure)


def same_path(a: PathSignature, b: PathSignature) -> bool:
    """Whether two recalculations may have taken the same path.

    An unknown path on either side counts as the same: it errs toward refusing.
    """
    return a.unsure or b.unsure or a.outcomes == b.outcomes


# -- roots -------------------------------------------------------------------


def engine_error_detail(code: str) -> str:
    """Reason detail for an error code the engine returned that Excel does not have."""
    return f"engine error {code} is not an Excel error"


def _add(roots: dict[CellRef, list[Reason]], cell: CellRef, reason: Reason) -> None:
    bucket = roots.setdefault(cell, [])
    if reason not in bucket:
        bucket.append(reason)


def _sorted_roots(roots: dict[CellRef, list[Reason]]) -> dict[CellRef, list[Reason]]:
    return {cell: sorted(roots[cell], key=reason_key) for cell in sorted(roots)}


def unplaced_roots(plan: ProbePlan) -> dict[CellRef, list[Reason]]:
    """Roots for probes that could not be placed; the same for every policy.

    Branch and lookup probes give ``not_checked_path``, rounding and comparison probes
    ``not_checked_value``, and a dynamic-reference probe a ``refused_pattern``.
    """
    roots: dict[CellRef, list[Reason]] = {}
    for cell, kinds in plan.unplaced.items():
        for kind in kinds:
            reason_kind, detail = _UNPLACED[kind]
            _add(roots, cell, Reason(reason_kind, cell, detail))
    return _sorted_roots(roots)


def policy_roots(
    plan: ProbePlan,
    readings: ProbeReadings,
    exercised: Exercised,
    *,
    include_unplaced: bool = True,
) -> dict[CellRef, list[Reason]]:
    """Roots one policy's probe readings give, sorted by cell then reason.

    Pass ``include_unplaced=False`` when :func:`unplaced_roots` already went into the
    static roots, so the per-policy work stays small.
    """
    roots: dict[CellRef, list[Reason]] = {}
    branch_fam = _families(plan, ProbeKind.branch)
    for key, outcome in readings.branch.items():
        if (_family_of(branch_fam, key), key[1], outcome) in exercised.branches:
            continue
        function = branch_fam.get(key, ("", "branch"))[1]
        detail = f"{function} #{key[1]} took {outcome!r}, not exercised in the trust run"
        _add(roots, key[0], Reason(ReasonKind.unexercised_branch, key[0], detail))
    lookup_fam = _families(plan, ProbeKind.lookup)
    for key, edge in readings.lookup_edge.items():
        if (_family_of(lookup_fam, key), key[1], edge) in exercised.lookup_edges:
            continue
        function = lookup_fam.get(key, ("", "lookup"))[1]
        detail = f"{function} #{key[1]} landed {edge}, not exercised in the trust run"
        _add(roots, key[0], Reason(ReasonKind.unexercised_lookup_edge, key[0], detail))
    for cell in readings.near_rounding:
        detail = "rounding input within 1e-12 of a flip point"
        _add(roots, cell, Reason(ReasonKind.near_rounding_flip, cell, detail))
    for cell in readings.near_comparison:
        detail = "comparison sides within 1e-12 of each other"
        _add(roots, cell, Reason(ReasonKind.near_comparison, cell, detail))
    for cell in readings.dynamic_unread:
        _add(roots, cell, Reason(ReasonKind.refused_pattern, cell, DYNAMIC_NOT_CHECKED))
    for cell, code in readings.engine_errors.items():
        _add(roots, cell, Reason(ReasonKind.engine_error, cell, engine_error_detail(code)))
    for cell in readings.type_mismatch:
        _add(roots, cell, Reason(ReasonKind.refused_pattern, cell, LOOKUP_TYPE))
    if include_unplaced:
        for cell, reasons in unplaced_roots(plan).items():
            for reason in reasons:
                _add(roots, cell, reason)
    return _sorted_roots(roots)


# -- propagation -------------------------------------------------------------


def _label(reasons: Iterable[Reason]) -> Label:
    by_kind: dict[ReasonKind, set[Reason]] = {}
    for reason in reasons:
        by_kind.setdefault(reason.kind, set()).add(reason)
    return {
        kind: tuple(sorted(rs, key=reason_key)[:MAX_REASONS_PER_KIND])
        for kind, rs in by_kind.items()
    }


def _merge(into: Label | None, add: Label) -> Label | None:
    """``into`` widened by ``add``; None when nothing changed."""
    out = dict(into) if into else {}
    changed = False
    for kind, extra in add.items():
        cur = out.get(kind, ())
        merged = tuple(sorted(set(cur) | set(extra), key=reason_key)[:MAX_REASONS_PER_KIND])
        if merged != cur:
            out[kind] = merged
            changed = True
    return out if changed else None


def _components(succ: dict[CellRef, list[CellRef]]) -> list[list[CellRef]]:
    """Strongly connected components of ``succ``, upstream components first.

    Iterative Tarjan: it emits a component only after every component it reaches, so
    the reversed emission order is a topological order of the condensation.
    """
    index: dict[CellRef, int] = {}
    low: dict[CellRef, int] = {}
    on_stack: set[CellRef] = set()
    stack: list[CellRef] = []
    emitted: list[list[CellRef]] = []
    for start in succ:
        if start in index:
            continue
        index[start] = low[start] = len(index)
        stack.append(start)
        on_stack.add(start)
        work: list[tuple[CellRef, int]] = [(start, 0)]
        while work:
            node, i = work[-1]
            nexts = succ[node]
            if i < len(nexts):
                work[-1] = (node, i + 1)
                nxt = nexts[i]
                if nxt not in index:
                    index[nxt] = low[nxt] = len(index)
                    stack.append(nxt)
                    on_stack.add(nxt)
                    work.append((nxt, 0))
                elif nxt in on_stack:
                    low[node] = min(low[node], index[nxt])
                continue
            work.pop()
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[node])
            if low[node] == index[node]:
                comp: list[CellRef] = []
                while True:
                    member = stack.pop()
                    on_stack.discard(member)
                    comp.append(member)
                    if member == node:
                        break
                emitted.append(comp)
    emitted.reverse()
    return emitted


def _propagate(
    graph: WorkbookGraph,
    seeds: dict[CellRef, Label],
    extra: dict[CellRef, list[CellRef]],
) -> dict[CellRef, Label]:
    """Labels of every cell reachable forward from ``seeds``; ``extra`` adds edges.

    One pass over the reached subgraph's components in topological order: a component
    (a circular block, or a single cell) gets its members' seeds plus everything
    flowing in, shared by all its members, and passes that on.
    """
    succ: dict[CellRef, list[CellRef]] = {}
    queue = deque(sorted(seeds))
    succ.update((cell, []) for cell in queue)
    while queue:
        cur = queue.popleft()
        nexts = sorted(graph.dependents(cur).union(extra.get(cur, ())))
        succ[cur] = nexts
        for nxt in nexts:
            if nxt not in succ:
                succ[nxt] = []
                queue.append(nxt)
    labels: dict[CellRef, Label] = {}
    incoming: dict[CellRef, Label] = {}
    for comp in _components(succ):
        label: Label = {}
        for member in comp:
            for part in (seeds.get(member), incoming.pop(member, None)):
                if part:
                    label = _merge(label, part) or label
        members = set(comp)
        for member in comp:
            labels[member] = label
            for nxt in succ[member]:
                if nxt not in members:
                    incoming[nxt] = _merge(incoming.get(nxt), label) or incoming.get(nxt, {})
    return labels


@dataclass(slots=True)
class StaticLabels:
    """Labels of the run's static roots, propagated once through the static graph."""

    labels: dict[CellRef, Label]


def propagate_static(graph: WorkbookGraph, roots: dict[CellRef, list[Reason]]) -> StaticLabels:
    """Propagate roots that hold for every policy (trust failures, refused patterns,
    external links, :func:`unplaced_roots`) once per run, for :func:`assign`."""
    seeds = {cell: _label(reasons) for cell, reasons in roots.items() if reasons}
    return StaticLabels(_propagate(graph, seeds, {}))


def _status(reasons: list[Reason]) -> Status:
    status = Status.OK
    for reason in reasons:
        candidate = SEVERITY[reason.kind]
        if _RANK[candidate] > _RANK[status]:
            status = candidate
    return status


def assign(
    graph: WorkbookGraph,
    outputs: list[CellRef],
    roots: dict[CellRef, list[Reason]],
    dynamic_refs: dict[CellRef, list[CellRef]],
    *,
    static: StaticLabels | None = None,
) -> dict[CellRef, tuple[Status, list[Reason]]]:
    """Status and reasons for each output, in ``outputs`` order.

    An output carries the reasons of every root whose closure reaches it, through the
    static graph plus this policy's ``dynamic_refs`` (cell -> cells its dynamic
    reference landed on), at most five per kind with the smallest roots first, sorted.
    Its status is the most severe of those reasons; ``ok`` with no reasons otherwise.
    ``static`` (from :func:`propagate_static`) adds the run-wide roots.
    """
    extra: dict[CellRef, list[CellRef]] = {}
    for cell, targets in dynamic_refs.items():
        for target in targets:
            extra.setdefault(target, []).append(cell)
    seeds = {cell: _label(reasons) for cell, reasons in roots.items() if reasons}
    base = static.labels if static is not None else {}
    if base:
        # Static reasons on a dynamic target flow on to the cell that pointed at it.
        for target, cells in extra.items():
            if target not in base:
                continue
            for cell in cells:
                merged = _merge(seeds.get(cell), base[target])
                if merged is not None:
                    seeds[cell] = merged
    labels = _propagate(graph, seeds, extra)
    out: dict[CellRef, tuple[Status, list[Reason]]] = {}
    for cell in outputs:
        if cell in out:
            continue
        label = _merge(base.get(cell), labels.get(cell, {})) or base.get(cell, {})
        reasons = sorted((r for rs in label.values() for r in rs), key=reason_key)
        out[cell] = (_status(reasons), reasons)
    return out
