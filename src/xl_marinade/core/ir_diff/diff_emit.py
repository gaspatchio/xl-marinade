# ABOUTME: Stage 5a — Emit changes for all entity classes.
# ABOUTME: Includes uniform-change suppression for binding-level vs cell-level dedup.

from __future__ import annotations

import re

from xl_marinade.core.ir_diff import change_types as CT
from xl_marinade.core.ir_diff.model import (
    AxisMap,
    BindingKey,
    BindingMatch,
    CellKey,
    CellMatch,
    Change,
    IRModel,
    SheetMatch,
)

# A single-quoted sheet prefix whose name needs no quoting (plain identifier).
# The lookbehind keeps us out of escaped-quote names like 'Bob''s'!.
_QUOTED_SIMPLE_SHEET = re.compile(r"(?<!')'([A-Za-z_][A-Za-z0-9_]*)'!")


def normalize_sheet_quoting(f: str | None) -> str | None:
    """Drop single-quotes around sheet names that don't need them
    ('Calculations'!R[1]C -> Calculations!R[1]C).

    Two saves of one workbook can store the same reference with and
    without quotes, which made quoting-only differences surface as
    formula changes. Double-quoted string literals are left untouched;
    names needing quotes (spaces, punctuation, embedded quotes) don't
    match the identifier pattern and keep theirs."""
    if not f or "'" not in f:
        return f
    parts = f.split('"')
    for i in range(0, len(parts), 2):  # even segments sit outside string literals
        parts[i] = _QUOTED_SIMPLE_SHEET.sub(r"\1!", parts[i])
    return '"'.join(parts)


def same_formula(fa: str | None, fb: str | None) -> bool:
    """Formula equality modulo redundant sheet-name quoting."""
    if fa == fb:
        return True
    if fa is None or fb is None:
        return False
    return normalize_sheet_quoting(fa) == normalize_sheet_quoting(fb)


def diff_metadata(a: IRModel, b: IRModel) -> list[Change]:
    """Diff ir_metadata key-value pairs."""
    changes = []
    all_keys = sorted(set(a.metadata.keys()) | set(b.metadata.keys()))
    diffs = {}
    for k in all_keys:
        va = a.metadata.get(k)
        vb = b.metadata.get(k)
        if va != vb:
            diffs[k] = {"old": va, "new": vb}
    if diffs:
        changes.append(
            Change(
                type=CT.IR_METADATA_CHANGED,
                sort_key=(),
                details={"changed_keys": diffs},
            )
        )
    return changes


def diff_roots(a: IRModel, b: IRModel) -> list[Change]:
    """Diff user_roots."""
    a_roots = [(r.sheet, r.range_a1, r.label_hint) for r in a.roots]
    b_roots = [(r.sheet, r.range_a1, r.label_hint) for r in b.roots]
    if a_roots != b_roots:
        return [
            Change(
                type=CT.ROOT_CHANGED,
                sort_key=(),
                details={
                    "root_a": [{"sheet": r.sheet, "range": r.range_a1} for r in a.roots],
                    "root_b": [{"sheet": r.sheet, "range": r.range_a1} for r in b.roots],
                },
            )
        ]
    return []


def diff_sheets(sheet_match: SheetMatch) -> list[Change]:
    """Emit sheet add/remove/rename changes."""
    changes = []
    for sa, sb, s_bind, s_cell, s_coord in sorted(sheet_match.renamed):
        changes.append(
            Change(
                type=CT.SHEET_RENAMED,
                sort_key=(sa,),
                details={
                    "sheet_a": sa,
                    "sheet_b": sb,
                    "score_bind": round(s_bind, 4),
                    "score_cell": round(s_cell, 4),
                },
            )
        )
    for s in sorted(sheet_match.removed):
        changes.append(Change(type=CT.SHEET_REMOVED, sort_key=(s,), details={"sheet": s}))
    for s in sorted(sheet_match.added):
        changes.append(Change(type=CT.SHEET_ADDED, sort_key=(s,), details={"sheet": s}))
    return changes


def diff_axes(axis_maps: dict[str, AxisMap], sheet_match: SheetMatch) -> list[Change]:
    """Emit row/column insertion/deletion changes."""
    changes = []
    for sa in sorted(axis_maps.keys()):
        sb = sheet_match.matched.get(sa, sa)
        axis = axis_maps[sa]
        for at_row, count in axis.rows_inserted:
            changes.append(
                Change(
                    type=CT.ROWS_INSERTED,
                    sort_key=(sb, at_row),
                    details={"sheet": sb, "at_row": at_row, "count": count},
                )
            )
        for at_row, count in axis.rows_deleted:
            changes.append(
                Change(
                    type=CT.ROWS_DELETED,
                    sort_key=(sb, at_row),
                    details={"sheet": sb, "at_row": at_row, "count": count},
                )
            )
        for at_col, count in axis.cols_inserted:
            changes.append(
                Change(
                    type=CT.COLS_INSERTED,
                    sort_key=(sb, at_col),
                    details={"sheet": sb, "at_col": at_col, "count": count},
                )
            )
        for at_col, count in axis.cols_deleted:
            changes.append(
                Change(
                    type=CT.COLS_DELETED,
                    sort_key=(sb, at_col),
                    details={"sheet": sb, "at_col": at_col, "count": count},
                )
            )
    return changes


def diff_names(a: IRModel, b: IRModel) -> list[Change]:
    """Diff defined_names."""
    changes = []
    all_keys = sorted(set(a.names.keys()) | set(b.names.keys()))
    for nk in all_keys:
        nd_a = a.names.get(nk)
        nd_b = b.names.get(nk)
        if nd_a and not nd_b:
            changes.append(
                Change(type=CT.NAME_REMOVED, sort_key=nk, details={"name": nk[0], "scope": nk[1]})
            )
        elif nd_b and not nd_a:
            changes.append(
                Change(
                    type=CT.NAME_ADDED,
                    sort_key=nk,
                    details={"name": nk[0], "scope": nk[1], "destinations": nd_b.destinations},
                )
            )
        elif nd_a and nd_b:
            if nd_a.destinations != nd_b.destinations:
                changes.append(
                    Change(
                        type=CT.NAME_DESTINATIONS_CHANGED,
                        sort_key=nk,
                        details={
                            "name": nk[0],
                            "scope": nk[1],
                            "old": nd_a.destinations,
                            "new": nd_b.destinations,
                        },
                    )
                )
            if nd_a.is_external != nd_b.is_external:
                changes.append(
                    Change(
                        type=CT.NAME_METADATA_CHANGED,
                        sort_key=nk,
                        details={
                            "name": nk[0],
                            "scope": nk[1],
                            "field": "is_external",
                            "old": nd_a.is_external,
                            "new": nd_b.is_external,
                        },
                    )
                )
    return changes


def diff_tables(a: IRModel, b: IRModel) -> list[Change]:
    """Diff table_candidates."""
    changes = []
    all_ids = sorted(set(a.tables.keys()) | set(b.tables.keys()))
    for tid in all_ids:
        ta = a.tables.get(tid)
        tb = b.tables.get(tid)
        if ta and not tb:
            changes.append(
                Change(
                    type=CT.TABLE_CANDIDATE_REMOVED,
                    sort_key=(ta.sheet, ta.r1, ta.c1),
                    details={"candidate_id": tid, "sheet": ta.sheet},
                )
            )
        elif tb and not ta:
            changes.append(
                Change(
                    type=CT.TABLE_CANDIDATE_ADDED,
                    sort_key=(tb.sheet, tb.r1, tb.c1),
                    details={"candidate_id": tid, "sheet": tb.sheet},
                )
            )
        elif ta and tb:
            diffs = {}
            for attr in (
                "sheet",
                "kind",
                "r1",
                "c1",
                "r2",
                "c2",
                "range_a1",
                "confidence",
                "reasons_json",
                "members",
            ):
                va = getattr(ta, attr)
                vb = getattr(tb, attr)
                if va != vb:
                    diffs[attr] = {"old": va, "new": vb}
            if diffs:
                changes.append(
                    Change(
                        type=CT.TABLE_CANDIDATE_CHANGED,
                        sort_key=(tb.sheet, tb.r1, tb.c1),
                        details={"candidate_id": tid, "changed": diffs},
                    )
                )
    return changes


def _formula_change_is_partition_artifact(
    a: IRModel,
    b: IRModel,
    ak: BindingKey,
    bk: BindingKey,
) -> bool:
    """True when a representative-formula mismatch has no cell-level support.

    The binding builder may partition the same cells differently across two
    extractions (e.g. v1 keeps W9:W993 as one binding whose representative is
    the row-9 seed formula; v2 splits W9 off so the representative becomes
    the row-10 steady-state formula). The representatives then differ while
    every actual cell formula is byte-identical — reporting that as
    BINDING_FORMULA_CHANGED fabricates a methodology shift downstream
    (Cycle 17 #365: 8/9 phantom logic changes on one model's v1↔v2).

    Cross-check the union of both bindings' absolute ranges: if every cell's
    formula_r1c1 matches across models, the representative mismatch is a
    partition artifact, not a logic change. Sound for moved bindings too —
    a real move changes cells at absolute coordinates, so the check fails
    and the change is emitted as before.
    """
    r1 = min(ak.top_left_row, bk.top_left_row)
    c1 = min(ak.top_left_col, bk.top_left_col)
    r2 = max(ak.top_left_row + ak.shape_rows, bk.top_left_row + bk.shape_rows) - 1
    c2 = max(ak.top_left_col + ak.shape_cols, bk.top_left_col + bk.shape_cols) - 1
    for row in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            ca = a.cells.get(CellKey(ak.sheet, row, col))
            cb = b.cells.get(CellKey(bk.sheet, row, col))
            if not same_formula(ca.formula_r1c1 if ca else None, cb.formula_r1c1 if cb else None):
                return False
    return True


def diff_bindings(
    a: IRModel,
    b: IRModel,
    binding_match: BindingMatch,
    scope_mode: bool,
) -> list[Change]:
    """Diff bindings and classify unmatched ones."""
    changes = []

    # Matched pairs
    for ak in sorted(binding_match.matched.keys()):
        bk = binding_match.matched[ak]
        ad = a.bindings.get(ak)
        bd = b.bindings.get(bk)
        if not ad or not bd:
            continue

        mt = binding_match.match_type.get(ak, "exact")

        # Common binding ID fields for all matched-pair changes
        _ids = {
            "binding_id_a": ad.original_binding_id,
            "binding_id_b": bd.original_binding_id,
        }

        if mt == "moved":
            changes.append(
                Change(
                    type=CT.BINDING_MOVED,
                    sort_key=(bk.sheet, bk.top_left_row, bk.top_left_col),
                    details={
                        "old_address": ad.address_a1,
                        "new_address": bd.address_a1,
                        "sheet": bk.sheet,
                        **_ids,
                    },
                )
            )

        if (ak.shape_rows, ak.shape_cols) != (bk.shape_rows, bk.shape_cols):
            changes.append(
                Change(
                    type=CT.BINDING_RESIZED,
                    sort_key=(bk.sheet, bk.top_left_row, bk.top_left_col),
                    details={
                        "sheet": bk.sheet,
                        "address": bd.address_a1,
                        "old_shape": (ak.shape_rows, ak.shape_cols),
                        "new_shape": (bk.shape_rows, bk.shape_cols),
                        **_ids,
                    },
                )
            )

        if ad.binding_type != bd.binding_type:
            changes.append(
                Change(
                    type=CT.BINDING_TYPE_CHANGED,
                    sort_key=(bk.sheet, bk.top_left_row, bk.top_left_col),
                    details={
                        "sheet": bk.sheet,
                        "address": bd.address_a1,
                        "old": ad.binding_type,
                        "new": bd.binding_type,
                        **_ids,
                    },
                )
            )

        if not same_formula(
            ad.formula_r1c1, bd.formula_r1c1
        ) and not _formula_change_is_partition_artifact(a, b, ak, bk):
            changes.append(
                Change(
                    type=CT.BINDING_FORMULA_CHANGED,
                    sort_key=(bk.sheet, bk.top_left_row, bk.top_left_col),
                    details={
                        "sheet": bk.sheet,
                        "address": bd.address_a1,
                        "old_formula": ad.formula_r1c1,
                        "new_formula": bd.formula_r1c1,
                        **_ids,
                    },
                )
            )

        if ad.label != bd.label or ad.classification != bd.classification:
            changes.append(
                Change(
                    type=CT.BINDING_LABEL_CHANGED,
                    sort_key=(bk.sheet, bk.top_left_row, bk.top_left_col),
                    details={
                        "sheet": bk.sheet,
                        "address": bd.address_a1,
                        "old_label": ad.label,
                        "new_label": bd.label,
                        "old_class": ad.classification,
                        "new_class": bd.classification,
                        **_ids,
                    },
                )
            )

        # Metadata changes
        meta_diffs = {}
        for attr in (
            "confidence",
            "is_orphan",
            "extraction_source",
            "evidence_sha256",
            "spatial_sha256",
        ):
            va = getattr(ad, attr)
            vb = getattr(bd, attr)
            if va != vb:
                meta_diffs[attr] = {"old": va, "new": vb}
        if meta_diffs:
            changes.append(
                Change(
                    type=CT.BINDING_METADATA_CHANGED,
                    sort_key=(bk.sheet, bk.top_left_row, bk.top_left_col),
                    details={
                        "sheet": bk.sheet,
                        "address": bd.address_a1,
                        "changed": meta_diffs,
                        **_ids,
                    },
                )
            )

    # Unmatched A bindings
    for ak in sorted(binding_match.removed):
        ad = a.bindings.get(ak)
        sk = (ak.sheet, ak.top_left_row, ak.top_left_col)
        _id_a = {"binding_id_a": ad.original_binding_id} if ad else {}
        if scope_mode:
            changes.append(
                Change(
                    type=CT.BINDING_OUT_OF_SCOPE,
                    sort_key=sk,
                    details={
                        "sheet": ak.sheet,
                        "address": ad.address_a1 if ad else "",
                        "label": ad.label if ad else None,
                        **_id_a,
                    },
                )
            )
        else:
            changes.append(
                Change(
                    type=CT.BINDING_REMOVED,
                    sort_key=sk,
                    details={
                        "sheet": ak.sheet,
                        "address": ad.address_a1 if ad else "",
                        "label": ad.label if ad else None,
                        **_id_a,
                    },
                )
            )

    # Unmatched B bindings
    for bk in sorted(binding_match.added):
        bd = b.bindings.get(bk)
        sk = (bk.sheet, bk.top_left_row, bk.top_left_col)
        _id_b = {"binding_id_b": bd.original_binding_id} if bd else {}
        if scope_mode:
            changes.append(
                Change(
                    type=CT.BINDING_NOW_IN_SCOPE,
                    sort_key=sk,
                    details={
                        "sheet": bk.sheet,
                        "address": bd.address_a1 if bd else "",
                        "label": bd.label if bd else None,
                        **_id_b,
                    },
                )
            )
        else:
            changes.append(
                Change(
                    type=CT.BINDING_ADDED,
                    sort_key=sk,
                    details={
                        "sheet": bk.sheet,
                        "address": bd.address_a1 if bd else "",
                        "label": bd.label if bd else None,
                        **_id_b,
                    },
                )
            )

    return changes


# Cycle-17 #430: two extraction runs can serialize the SAME float64 differently
# (Excel E-notation "6.09E-2" vs plain decimal "0.0609...", or 16 vs 17 sig
# figs), so a byte/sha comparison of value_json flags VALUE_CHANGED on cells
# whose numeric value is identical. Measured across two real-model regression
# comparisons: one model's (v1)↔(v2) had 84,536/84,537 value diffs that were
# pure serialization noise (all within 1e-9 relative), vs another model's
# v1↔v4 where 94% were genuine. A numeric-aware comparison kills the noise
# channel without touching real value/string changes. 1e-9 relative is the
# empirically-validated separator: the noisy model's noise is all <1e-9, every
# real change on both workbooks is >1e-6.
_VALUE_REL_TOL = 1e-9


def _parse_cell_number(value_json: str | None) -> float | None:
    """Parse a cell value_json as a float, including JSON-string-wrapped
    numbers (e.g. '"6.09E-2"'). Returns None for non-numeric / unparseable /
    non-finite values, which then fall through to a normal (string) compare."""
    if value_json is None:
        return None
    import json as _json
    import math as _math

    try:
        v = _json.loads(value_json)
    except Exception:
        v = value_json
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
    elif isinstance(v, str):
        try:
            f = float(v)
        except (ValueError, TypeError):
            return None
    else:
        return None
    return f if _math.isfinite(f) else None


def _values_numerically_equivalent(va_json: str | None, vb_json: str | None) -> bool:
    """True when both values parse to finite floats equal within _VALUE_REL_TOL
    relative tolerance — i.e. the same number serialized differently. Anything
    non-numeric (text, blanks, type changes) returns False so it is still
    emitted as a real change."""
    a = _parse_cell_number(va_json)
    b = _parse_cell_number(vb_json)
    if a is None or b is None:
        return False
    return abs(a - b) <= _VALUE_REL_TOL * max(abs(a), abs(b))


def _value_is_absent(value_json: str | None) -> bool:
    """A cell has no cached value (no value blob / empty serialization)."""
    return value_json is None or value_json in ("", '""', "null")


def _is_recalc_value_gap(sa, sb) -> bool:
    """Cycle-17 #430/#366: a formula-bearing cell whose formula is UNCHANGED but
    whose cached value is present on one side and absent on the other. This is a
    save/recalc-state artifact, not a model change — the unchanged formula means
    the cell computes the same thing; the missing value just means that version
    was saved without recalculation (one model's v2 had 201,633 such cells). Mirrors
    #365's principle of not reporting a binding-level artifact as a real change.
    Constant cells (no formula) are excluded: there a value appearing/vanishing
    IS a real input edit."""
    fa, fb = sa.formula_r1c1, sb.formula_r1c1
    if not fa or not fb or not same_formula(fa, fb):
        return False
    return _value_is_absent(sa.value_json) != _value_is_absent(sb.value_json)


def diff_cells(
    a: IRModel,
    b: IRModel,
    cell_match: CellMatch,
    binding_match: BindingMatch,
) -> list[Change]:
    """Diff matched cells. Applies uniform-change suppression."""
    changes = []

    # Collect per-binding uniform checks

    for ck_a in sorted(cell_match.matched.keys()):
        ck_b = cell_match.matched[ck_a]
        sa = a.cells.get(ck_a)
        sb = b.cells.get(ck_b)
        if not sa or not sb:
            continue

        sk = (ck_b.sheet, ck_b.row, ck_b.col)

        if not same_formula(sa.formula_r1c1, sb.formula_r1c1):
            changes.append(
                Change(
                    type=CT.FORMULA_CHANGED,
                    sort_key=sk,
                    details={
                        "cell": f"{ck_b.sheet}!{ck_b.row},{ck_b.col}",
                        "old_formula": sa.formula_r1c1,
                        "new_formula": sb.formula_r1c1,
                    },
                )
            )

        if (
            sa.value_sha256 != sb.value_sha256
            and not _values_numerically_equivalent(sa.value_json, sb.value_json)
            and not _is_recalc_value_gap(sa, sb)
        ):
            changes.append(
                Change(
                    type=CT.VALUE_CHANGED,
                    sort_key=sk,
                    details={
                        "cell": f"{ck_b.sheet}!{ck_b.row},{ck_b.col}",
                        "old_value": sa.value_json,
                        "new_value": sb.value_json,
                    },
                )
            )

        if sa.format_sha256 != sb.format_sha256:
            changes.append(
                Change(
                    type=CT.FORMAT_CHANGED,
                    sort_key=sk,
                    details={"cell": f"{ck_b.sheet}!{ck_b.row},{ck_b.col}"},
                )
            )

        if sa.data_type != sb.data_type:
            changes.append(
                Change(
                    type=CT.DTYPE_CHANGED,
                    sort_key=sk,
                    details={
                        "cell": f"{ck_b.sheet}!{ck_b.row},{ck_b.col}",
                        "old": sa.data_type,
                        "new": sb.data_type,
                    },
                )
            )

        if sa.is_array_formula != sb.is_array_formula:
            changes.append(
                Change(
                    type=CT.ARRAY_FORMULA_FLAG_CHANGED,
                    sort_key=sk,
                    details={"cell": f"{ck_b.sheet}!{ck_b.row},{ck_b.col}"},
                )
            )

        if sa.is_spilled != sb.is_spilled or sa.spill_origin != sb.spill_origin:
            changes.append(
                Change(
                    type=CT.SPILL_CHANGED,
                    sort_key=sk,
                    details={"cell": f"{ck_b.sheet}!{ck_b.row},{ck_b.col}"},
                )
            )

    return changes


def diff_edges(
    a: IRModel,
    b: IRModel,
    cell_match: CellMatch,
    binding_match: BindingMatch,
) -> list[Change]:
    """Diff dependency graph edges."""
    changes = []

    # Cell internal edges
    for edge in sorted(a.cell_edges - b.cell_edges):
        changes.append(
            Change(
                type=CT.CELL_EDGE_REMOVED,
                sort_key=(edge.from_key.sheet, edge.from_key.row, edge.from_key.col),
                details={"from": _ck_str(edge.from_key), "to": _ck_str(edge.to_key)},
            )
        )
    for edge in sorted(b.cell_edges - a.cell_edges):
        changes.append(
            Change(
                type=CT.CELL_EDGE_ADDED,
                sort_key=(edge.from_key.sheet, edge.from_key.row, edge.from_key.col),
                details={"from": _ck_str(edge.from_key), "to": _ck_str(edge.to_key)},
            )
        )

    # External edges
    for edge in sorted(a.external_edges - b.external_edges):
        changes.append(
            Change(
                type=CT.EXTERNAL_EDGE_REMOVED,
                sort_key=(edge.from_key.sheet, edge.from_key.row, edge.from_key.col),
                details={"from": _ck_str(edge.from_key), "external_ref": edge.external_ref},
            )
        )
    for edge in sorted(b.external_edges - a.external_edges):
        changes.append(
            Change(
                type=CT.EXTERNAL_EDGE_ADDED,
                sort_key=(edge.from_key.sheet, edge.from_key.row, edge.from_key.col),
                details={"from": _ck_str(edge.from_key), "external_ref": edge.external_ref},
            )
        )

    # Range edges
    for edge in sorted(a.range_edges - b.range_edges):
        changes.append(
            Change(
                type=CT.RANGE_EDGE_REMOVED,
                sort_key=(edge.from_key.sheet, edge.from_key.row, edge.from_key.col),
                details={"from": _ck_str(edge.from_key), "to_range": edge.to_range_a1},
            )
        )
    for edge in sorted(b.range_edges - a.range_edges):
        changes.append(
            Change(
                type=CT.RANGE_EDGE_ADDED,
                sort_key=(edge.from_key.sheet, edge.from_key.row, edge.from_key.col),
                details={"from": _ck_str(edge.from_key), "to_range": edge.to_range_a1},
            )
        )

    # Binding edges
    for edge in sorted(a.binding_edges - b.binding_edges):
        changes.append(
            Change(
                type=CT.BINDING_EDGE_REMOVED,
                sort_key=(edge.from_key.sheet, edge.from_key.top_left_row),
                details={"from_sheet": edge.from_key.sheet, "to_sheet": edge.to_key.sheet},
            )
        )
    for edge in sorted(b.binding_edges - a.binding_edges):
        changes.append(
            Change(
                type=CT.BINDING_EDGE_ADDED,
                sort_key=(edge.from_key.sheet, edge.from_key.top_left_row),
                details={"from_sheet": edge.from_key.sheet, "to_sheet": edge.to_key.sheet},
            )
        )

    return changes


def diff_label_evidence(a: IRModel, b: IRModel) -> list[Change]:
    """Diff binding label candidate cells."""
    if a.label_evidence != b.label_evidence:
        # Group changes by binding
        a_by_bk = _group_label_evidence(a)
        b_by_bk = _group_label_evidence(b)
        changes = []
        for bk in sorted(set(a_by_bk.keys()) | set(b_by_bk.keys())):
            if a_by_bk.get(bk) != b_by_bk.get(bk):
                changes.append(
                    Change(
                        type=CT.BINDING_LABEL_EVIDENCE_CHANGED,
                        sort_key=(bk.sheet, bk.top_left_row, bk.top_left_col),
                        details={"binding_sheet": bk.sheet},
                    )
                )
        return changes
    return []


def _group_label_evidence(model: IRModel):
    from collections import defaultdict

    groups = defaultdict(set)
    for le in model.label_evidence:
        groups[le.binding_key].add(le)
    return groups


def diff_time_annotations(a: IRModel, b: IRModel) -> list[Change]:
    """Diff time index candidates and binding time annotations."""
    changes = []

    # Time index candidates (compare as sorted tuples)
    a_tics = {(t.sheet, t.rank): t for t in a.time_index_candidates}
    b_tics = {(t.sheet, t.rank): t for t in b.time_index_candidates}

    for k in sorted(set(a_tics.keys()) | set(b_tics.keys())):
        ta = a_tics.get(k)
        tb = b_tics.get(k)
        if ta and not tb:
            changes.append(
                Change(
                    type=CT.TIME_INDEX_CANDIDATE_REMOVED,
                    sort_key=k,
                    details={"sheet": k[0], "rank": k[1]},
                )
            )
        elif tb and not ta:
            changes.append(
                Change(
                    type=CT.TIME_INDEX_CANDIDATE_ADDED,
                    sort_key=k,
                    details={"sheet": k[0], "rank": k[1]},
                )
            )
        elif ta and tb:
            if (ta.binding_key, ta.confidence, ta.reasons_json) != (
                tb.binding_key,
                tb.confidence,
                tb.reasons_json,
            ):
                changes.append(
                    Change(
                        type=CT.TIME_INDEX_CANDIDATE_CHANGED,
                        sort_key=k,
                        details={"sheet": k[0], "rank": k[1]},
                    )
                )

    # Binding time annotations
    all_bks = sorted(
        set(a.binding_time_annotations.keys()) | set(b.binding_time_annotations.keys())
    )
    for bk in all_bks:
        ba = a.binding_time_annotations.get(bk)
        bb = b.binding_time_annotations.get(bk)
        sk = (bk.sheet, bk.top_left_row, bk.top_left_col)
        if ba and not bb:
            changes.append(
                Change(
                    type=CT.BINDING_TIME_ANNOTATION_REMOVED,
                    sort_key=sk,
                    details={"binding_sheet": bk.sheet},
                )
            )
        elif bb and not ba:
            changes.append(
                Change(
                    type=CT.BINDING_TIME_ANNOTATION_ADDED,
                    sort_key=sk,
                    details={"binding_sheet": bk.sheet},
                )
            )
        elif ba and bb and ba != bb:
            changes.append(
                Change(
                    type=CT.BINDING_TIME_ANNOTATION_CHANGED,
                    sort_key=sk,
                    details={"binding_sheet": bk.sheet},
                )
            )

    return changes


def diff_families(a: IRModel, b: IRModel, binding_match: BindingMatch) -> list[Change]:
    """Diff formula families (summary-only, not verification-active)."""
    changes = []
    all_fks = sorted(set(a.families.keys()) | set(b.families.keys()))

    for fk in all_fks:
        fa = a.families.get(fk)
        fb = b.families.get(fk)
        sk = fk  # (sheet, formula_r1c1)

        if fa and not fb:
            changes.append(
                Change(
                    type=CT.FAMILY_REMOVED, sort_key=sk, details={"sheet": fk[0], "formula": fk[1]}
                )
            )
        elif fb and not fa:
            changes.append(
                Change(
                    type=CT.FAMILY_ADDED,
                    sort_key=sk,
                    details={"sheet": fk[0], "formula": fk[1], "member_count": fb.member_count},
                )
            )
        elif fa and fb:
            if fa.member_count != fb.member_count:
                changes.append(
                    Change(
                        type=CT.FAMILY_RESIZED,
                        sort_key=sk,
                        details={
                            "sheet": fk[0],
                            "formula": fk[1],
                            "old_count": fa.member_count,
                            "new_count": fb.member_count,
                        },
                    )
                )
            if fa.representative_binding_key != fb.representative_binding_key:
                changes.append(
                    Change(
                        type=CT.FAMILY_REPRESENTATIVE_CHANGED,
                        sort_key=sk,
                        details={"sheet": fk[0], "formula": fk[1]},
                    )
                )

    return changes


def diff_resolution_metrics(a: IRModel, b: IRModel) -> list[Change]:
    """Diff resolution_metrics."""
    changes = []
    all_keys = sorted(set(a.resolution_metrics.keys()) | set(b.resolution_metrics.keys()))
    for k in all_keys:
        va = a.resolution_metrics.get(k)
        vb = b.resolution_metrics.get(k)
        if va is not None and vb is None:
            changes.append(
                Change(
                    type=CT.RESOLUTION_METRIC_REMOVED,
                    sort_key=k,
                    details={"function": k[0], "status": k[1], "count": va},
                )
            )
        elif vb is not None and va is None:
            changes.append(
                Change(
                    type=CT.RESOLUTION_METRIC_ADDED,
                    sort_key=k,
                    details={"function": k[0], "status": k[1], "count": vb},
                )
            )
        elif va != vb:
            changes.append(
                Change(
                    type=CT.RESOLUTION_METRIC_CHANGED,
                    sort_key=k,
                    details={"function": k[0], "status": k[1], "old_count": va, "new_count": vb},
                )
            )
    return changes


def _ck_str(ck):
    """Format CellKey as string."""
    return f"{ck.sheet}!R{ck.row}C{ck.col}"
