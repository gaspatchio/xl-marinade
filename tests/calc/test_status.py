"""Tests for status assignment: root causes propagated to every requested output."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import pytest

from tests.calc.conftest import F
from xl_marinade.calc.graph import WorkbookGraph
from xl_marinade.calc.models import CellRef, Reason, ReasonKind, Status
from xl_marinade.calc.probes import Probe, ProbeKind, ProbePlan, ProbeReadings
from xl_marinade.calc.status import (
    Exercised,
    assign,
    exercised_from,
    path_keys,
    path_signature,
    policy_roots,
    propagate_static,
    same_path,
    unplaced_roots,
)

Book = Callable[..., Path]
IrFor = Callable[[Path], Path]


def _c(a1: str) -> CellRef:
    return CellRef.parse(f"S!{a1}")


A1, B1, C1, D9, E1, G1 = (_c(x) for x in ("A1", "B1", "C1", "D9", "E1", "G1"))
B40, B50 = _c("B40"), _c("B50")

WRONG = Reason(ReasonKind.silently_wrong, B1, "engine 3, stored 4")
NEAR = Reason(ReasonKind.near_rounding_flip, B1, "near")


@pytest.fixture
def graph(make_book: Book, ir_for: IrFor) -> WorkbookGraph:
    path = make_book(
        "st",
        {
            "S": {
                "A1": 1,
                "B1": F("=A1*2", 2),
                "C1": F("=B1+1", 3),
                "G1": 4,
                "D9": F("=G1*3", 12),
                "E1": F("=7*2", 14),
                "A40": 1,
                "A50": -1,
                "B40": F("=IF(A40>0,1,2)", 1),
                "B50": F("=IF(A50>0,1,2)", 2),
                "C40": F("=VLOOKUP(A40,$A$40:$A$50,1,FALSE)", 1),
                "C50": F("=VLOOKUP(A50,$A$40:$A$50,1,FALSE)", -1),
            }
        },
    )
    return WorkbookGraph.from_ir(ir_for(path))


def _readings(**kw: object) -> ProbeReadings:
    base: dict[str, object] = {
        "branch": {},
        "lookup_edge": {},
        "near_rounding": set(),
        "near_comparison": set(),
        "dynamic_refs": {},
    }
    base.update(kw)
    return ProbeReadings(**base)  # type: ignore[arg-type]


def _plan(
    graph: WorkbookGraph, unplaced: dict[CellRef, list[ProbeKind]] | None = None
) -> ProbePlan:
    probes = [
        Probe(ProbeKind.branch, cell, graph.family(cell), 0, ("x",), {"function": "IF"})
        for cell in (B40, B50)
    ] + [
        Probe(ProbeKind.lookup, cell, graph.family(cell), 0, ("x", "y"), {"function": "VLOOKUP"})
        for cell in (_c("C40"), _c("C50"))
    ]
    return ProbePlan(probes, unplaced or {})


def test_families_are_shared(graph: WorkbookGraph) -> None:
    assert graph.family(B40) == graph.family(B50)
    assert graph.family(_c("C40")) == graph.family(_c("C50"))


def test_failing_root_refuses_downstream_output(graph: WorkbookGraph) -> None:
    out = assign(graph, [C1], {B1: [WRONG]}, {})
    assert out == {C1: (Status.REFUSED, [WRONG])}


def test_near_rounding_root_makes_output_uncertain(graph: WorkbookGraph) -> None:
    assert assign(graph, [C1], {B1: [NEAR]}, {}) == {C1: (Status.UNCERTAIN, [NEAR])}


def test_both_reasons_most_severe_wins(graph: WorkbookGraph) -> None:
    status, reasons = assign(graph, [C1], {B1: [NEAR, WRONG]}, {})[C1]
    assert status is Status.REFUSED
    assert reasons == [NEAR, WRONG]  # sorted by kind value


def test_output_not_downstream_stays_ok(graph: WorkbookGraph) -> None:
    out = assign(graph, [E1, A1], {B1: [WRONG]}, {})
    assert out == {E1: (Status.OK, []), A1: (Status.OK, [])}


def test_output_that_is_itself_a_root(graph: WorkbookGraph) -> None:
    assert assign(graph, [B1], {B1: [WRONG]}, {})[B1] == (Status.REFUSED, [WRONG])


def test_dynamic_ref_to_failing_cell_refuses(graph: WorkbookGraph) -> None:
    bad = Reason(ReasonKind.silently_wrong, D9, "x")
    assert assign(graph, [C1], {D9: [bad]}, {})[C1] == (Status.OK, [])
    assert assign(graph, [C1], {D9: [bad]}, {C1: [D9]})[C1] == (Status.REFUSED, [bad])


def test_dynamic_ref_carries_roots_upstream_of_target(graph: WorkbookGraph) -> None:
    # G1 -> D9 statically; C1's dynamic reference lands on D9 this policy.
    bad = Reason(ReasonKind.engine_error, G1, "x")
    assert assign(graph, [C1], {G1: [bad]}, {B1: [D9]})[C1] == (Status.REFUSED, [bad])


def test_chained_dynamic_refs(graph: WorkbookGraph) -> None:
    # E1 points at G1 (a root); C1 points at E1.
    bad = Reason(ReasonKind.engine_error, G1, "x")
    assert assign(graph, [C1], {G1: [bad]}, {C1: [E1], E1: [G1]})[C1][0] is Status.REFUSED


def test_external_link_reason_keeps_ok(graph: WorkbookGraph) -> None:
    held = Reason(ReasonKind.external_link_held, B1, "external link held at cached value")
    assert assign(graph, [C1], {B1: [held]}, {})[C1] == (Status.OK, [held])


def test_reasons_capped_at_five_per_kind_smallest_roots_first(graph: WorkbookGraph) -> None:
    roots = {
        cell: [Reason(ReasonKind.silently_wrong, cell, "w")]
        for cell in (_c("A1"), _c("B1"), _c("C1"), _c("G1"), _c("D9"), _c("A40"))
    }
    roots[_c("A40")].append(Reason(ReasonKind.near_comparison, _c("A40"), "n"))
    dyn = {C1: [D9, _c("B40")]}
    status, reasons = assign(graph, [C1], roots, dyn)[C1]
    assert status is Status.REFUSED
    wrong = [r.root for r in reasons if r.kind is ReasonKind.silently_wrong]
    assert wrong == sorted(roots)[:5]
    assert [r.kind for r in reasons].count(ReasonKind.near_comparison) == 1
    key = lambda r: (r.kind, r.root or CellRef("", 0, 0), r.detail)  # noqa: E731
    assert reasons == sorted(reasons, key=key)


def test_static_precompute_matches_one_shot(graph: WorkbookGraph) -> None:
    static_roots = {G1: [Reason(ReasonKind.engine_error, G1, "x")], B1: [WRONG]}
    per_policy = {A1: [Reason(ReasonKind.near_comparison, A1, "n")]}
    outputs = [C1, E1, D9, B1]
    dyn = {E1: [D9]}
    pre = propagate_static(graph, static_roots)
    together = {**static_roots, **per_policy}
    assert assign(graph, outputs, per_policy, dyn, static=pre) == assign(
        graph, outputs, together, dyn
    )


def test_branch_seen_in_family_member_counts_as_exercised(graph: WorkbookGraph) -> None:
    plan = _plan(graph)
    trust = _readings(branch={(B40, 0): "then"})
    exercised = exercised_from(plan, [trust])
    assert exercised == Exercised({(graph.family(B40), 0, "then")}, set())
    # Row 50 takes the branch row 40 took: exercised.
    seen = policy_roots(plan, _readings(branch={(B50, 0): "then"}), exercised)
    assert seen == {}
    # Row 50 takes the other branch: unexercised.
    roots = policy_roots(plan, _readings(branch={(B50, 0): "else"}), exercised)
    assert list(roots) == [B50]
    assert [r.kind for r in roots[B50]] == [ReasonKind.unexercised_branch]
    status, reasons = assign(graph, [B50], roots, {})[B50]
    assert status is Status.UNVERIFIED and reasons == roots[B50]


def test_lookup_edges(graph: WorkbookGraph) -> None:
    c40, c50 = _c("C40"), _c("C50")
    plan = _plan(graph)
    exercised = exercised_from(plan, [_readings(lookup_edge={(c40, 0): "interior"})])
    assert policy_roots(plan, _readings(lookup_edge={(c50, 0): "interior"}), exercised) == {}
    for edge in ("first", "last", "not_found"):
        roots = policy_roots(plan, _readings(lookup_edge={(c50, 0): edge}), exercised)
        assert [r.kind for r in roots[c50]] == [ReasonKind.unexercised_lookup_edge]
        assert assign(graph, [c50], roots, {})[c50][0] is Status.UNVERIFIED
    hit = exercised_from(plan, [_readings(lookup_edge={(c40, 0): "last"})])
    assert policy_roots(plan, _readings(lookup_edge={(c50, 0): "last"}), hit) == {}
    assert c50 in policy_roots(plan, _readings(lookup_edge={(c50, 0): "first"}), hit)


def test_exercised_across_several_trust_readings(graph: WorkbookGraph) -> None:
    plan = _plan(graph)
    exercised = exercised_from(
        plan, [_readings(branch={(B40, 0): "then"}), _readings(branch={(B40, 0): "else"})]
    )
    assert policy_roots(plan, _readings(branch={(B50, 0): "else"}), exercised) == {}


def test_near_flips_become_roots(graph: WorkbookGraph) -> None:
    plan = _plan(graph)
    roots = policy_roots(
        plan, _readings(near_rounding={B1}, near_comparison={B1, D9}), Exercised(set(), set())
    )
    assert [r.kind for r in roots[B1]] == [
        ReasonKind.near_comparison,
        ReasonKind.near_rounding_flip,
    ]
    assert [r.kind for r in roots[D9]] == [ReasonKind.near_comparison]
    assert all(r.root == cell for cell, rs in roots.items() for r in rs)


def test_unplaced_probes_and_unread_dynamic_refs(graph: WorkbookGraph) -> None:
    plan = _plan(
        graph,
        unplaced={
            B1: [ProbeKind.branch, ProbeKind.branch, ProbeKind.lookup],
            C1: [ProbeKind.rounding, ProbeKind.comparison],
            D9: [ProbeKind.dynamic_ref],
        },
    )
    static = unplaced_roots(plan)
    assert {r.kind for r in static[B1]} == {ReasonKind.not_checked_path}
    assert {r.kind for r in static[C1]} == {ReasonKind.not_checked_value}
    assert static[D9] == [Reason(ReasonKind.refused_pattern, D9, "dynamic reference not checked")]
    readings = _readings(dynamic_unread={E1})
    roots = policy_roots(plan, readings, Exercised(set(), set()))
    assert roots[E1] == [Reason(ReasonKind.refused_pattern, E1, "dynamic reference not checked")]
    for cell in (B1, C1, D9):
        assert roots[cell] == static[cell]
    lean = policy_roots(plan, readings, Exercised(set(), set()), include_unplaced=False)
    assert list(lean) == [E1]
    out = assign(graph, [B1, C1, D9, E1], roots, {})
    assert out[B1][0] is Status.UNVERIFIED
    assert out[C1] == (Status.UNCERTAIN, sorted(static[B1] + static[C1]))
    assert out[D9][0] is Status.REFUSED
    assert out[E1][0] is Status.REFUSED


def test_assign_is_deterministic(graph: WorkbookGraph) -> None:
    roots = {
        G1: [Reason(ReasonKind.engine_error, G1, "b"), Reason(ReasonKind.engine_error, G1, "a")],
        B1: [NEAR, WRONG],
    }
    first = assign(graph, [C1, D9, E1], roots, {E1: [D9]})
    again = assign(graph, [E1, D9, C1], dict(reversed(roots.items())), {E1: [D9]})
    assert first == again
    assert list(first) == [C1, D9, E1]
    assert first[D9][1] == [
        Reason(ReasonKind.engine_error, G1, "a"),
        Reason(ReasonKind.engine_error, G1, "b"),
    ]


def test_circular_block_shares_its_roots(make_book: Book, ir_for: IrFor) -> None:
    path = make_book(
        "circ",
        {
            "S": {
                "A1": 1,
                "X1": F("=Y1+A1", 0),
                "Y1": F("=X1+1", 0),
                "Z1": F("=Y1*2", 0),
                "W1": F("=A1+1", 2),
            }
        },
    )
    g = WorkbookGraph.from_ir(ir_for(path))
    x1, y1, z1, w1 = _c("X1"), _c("Y1"), _c("Z1"), _c("W1")
    bad = Reason(ReasonKind.engine_error, y1, "circular")
    out = assign(g, [x1, y1, z1, w1], {y1: [bad]}, {})
    assert out[x1] == out[y1] == out[z1] == (Status.REFUSED, [bad])
    assert out[w1] == (Status.OK, [])
    near = Reason(ReasonKind.near_comparison, A1, "n")
    both = assign(g, [z1], {y1: [bad], A1: [near]}, {})[z1]
    assert both == (Status.REFUSED, [bad, near])


def test_failing_or_refused_cells_vouch_for_nothing(graph: WorkbookGraph) -> None:
    plan = _plan(graph)
    trust = _readings(branch={(B40, 0): "then"}, lookup_edge={(_c("C40"), 0): "last"})
    exercised = exercised_from(plan, [trust], exclude={B40, _c("C40")})
    assert exercised == Exercised(set(), set())
    roots = policy_roots(plan, _readings(branch={(B50, 0): "then"}), exercised)
    assert [r.kind for r in roots[B50]] == [ReasonKind.unexercised_branch]


def test_unsure_readings_are_not_evidence(graph: WorkbookGraph) -> None:
    plan = _plan(graph)
    trust = _readings(
        branch={(B40, 0): "then"},
        unsure={(ProbeKind.branch, B40, 0)},
    )
    assert exercised_from(plan, [trust]) == Exercised(set(), set())
    # For a policy an unsure reading still counts as a path taken.
    roots = policy_roots(plan, trust, Exercised(set(), set()))
    assert [r.kind for r in roots[B40]] == [ReasonKind.unexercised_branch]


def test_engine_errors_and_lookup_type_mismatch_become_roots(graph: WorkbookGraph) -> None:
    plan = _plan(graph)
    roots = policy_roots(
        plan,
        _readings(engine_errors={B1: "#N/IMPL!"}, type_mismatch={D9}),
        Exercised(set(), set()),
    )
    assert roots[B1] == [
        Reason(ReasonKind.engine_error, B1, "engine error #N/IMPL! is not an Excel error")
    ]
    assert roots[D9] == [
        Reason(ReasonKind.refused_pattern, D9, "lookup value type differs from keys")
    ]


# -- path signatures (scoping a --check-against mismatch) --------------------


def test_path_keys_cover_branch_and_lookup_probes_of_the_cells(graph: WorkbookGraph) -> None:
    keys, unplaced = path_keys(_plan(graph), {B40, _c("C40")})
    assert keys == frozenset({(B40, 0, ProbeKind.branch), (_c("C40"), 0, ProbeKind.lookup)})
    assert not unplaced


def test_same_outcomes_are_the_same_path(graph: WorkbookGraph) -> None:
    keys, unplaced = path_keys(_plan(graph), {B40, B50})
    a = path_signature(_readings(branch={(B40, 0): "then", (B50, 0): "else"}), keys, unplaced)
    b = path_signature(_readings(branch={(B40, 0): "then", (B50, 0): "else"}), keys, unplaced)
    c = path_signature(_readings(branch={(B40, 0): "else", (B50, 0): "else"}), keys, unplaced)
    assert same_path(a, b)
    assert not same_path(a, c)


def test_readings_outside_the_keys_do_not_tell_paths_apart(graph: WorkbookGraph) -> None:
    keys, unplaced = path_keys(_plan(graph), {B40})
    a = path_signature(_readings(branch={(B40, 0): "then", (B50, 0): "then"}), keys, unplaced)
    b = path_signature(_readings(branch={(B40, 0): "then", (B50, 0): "else"}), keys, unplaced)
    assert same_path(a, b)


def test_a_probe_skipped_in_one_reading_is_a_different_path(graph: WorkbookGraph) -> None:
    keys, unplaced = path_keys(_plan(graph), {B40, B50})
    a = path_signature(_readings(branch={(B40, 0): "then"}), keys, unplaced)
    b = path_signature(_readings(branch={(B40, 0): "then", (B50, 0): "else"}), keys, unplaced)
    assert not same_path(a, b)


@pytest.mark.parametrize("side", ["mismatch", "policy"])
def test_an_unsure_probe_makes_the_paths_the_same(graph: WorkbookGraph, side: str) -> None:
    keys, unplaced = path_keys(_plan(graph), {B40, B50})
    unsure = {(ProbeKind.branch, B50, 0)}
    a = _readings(branch={(B40, 0): "then", (B50, 0): "then"})
    b = _readings(branch={(B40, 0): "else", (B50, 0): "else"})
    (a if side == "mismatch" else b).unsure = unsure
    assert same_path(path_signature(a, keys, unplaced), path_signature(b, keys, unplaced))


def test_an_unsure_probe_outside_the_keys_changes_nothing(graph: WorkbookGraph) -> None:
    keys, unplaced = path_keys(_plan(graph), {B40})
    a = _readings(branch={(B40, 0): "then"}, unsure={(ProbeKind.branch, B50, 0)})
    b = _readings(branch={(B40, 0): "else"})
    assert not same_path(path_signature(a, keys, unplaced), path_signature(b, keys, unplaced))


@pytest.mark.parametrize("kind", [ProbeKind.branch, ProbeKind.lookup])
def test_an_unplaced_path_probe_makes_every_path_the_same(
    graph: WorkbookGraph, kind: ProbeKind
) -> None:
    keys, unplaced = path_keys(_plan(graph, {B1: [kind]}), {B1, B40})
    assert unplaced
    a = path_signature(_readings(branch={(B40, 0): "then"}), keys, unplaced)
    b = path_signature(_readings(branch={(B40, 0): "else"}), keys, unplaced)
    assert same_path(a, b)


def test_an_unplaced_value_probe_does_not_hide_the_path(graph: WorkbookGraph) -> None:
    keys, unplaced = path_keys(_plan(graph, {B1: [ProbeKind.rounding]}), {B1, B40})
    assert not unplaced


def test_no_path_probes_means_one_shared_path(graph: WorkbookGraph) -> None:
    keys, unplaced = path_keys(_plan(graph), {B1, C1})
    assert keys == frozenset()
    a = path_signature(_readings(branch={(B40, 0): "then"}), keys, unplaced)
    b = path_signature(_readings(), keys, unplaced)
    assert same_path(a, b)
