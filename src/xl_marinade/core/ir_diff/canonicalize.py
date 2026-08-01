# ABOUTME: Stage 3 — Rewrite Version A's IRModel into Version B's namespace.
# ABOUTME: Transforms sheet names, coordinates, formulas, and JSON blobs.

from __future__ import annotations

from xl_marinade.core.ir_diff.formula_rewrite import rewrite_formula_r1c1
from xl_marinade.core.ir_diff.json_rewrite import (
    canonicalize_destinations_json,
    rewrite_json_blob,
)
from xl_marinade.core.ir_diff.model import (
    AxisMap,
    BindingDesc,
    BindingEdgeTuple,
    BindingKey,
    BindingTimeAnnotation,
    CellEdge,
    CellKey,
    CellSig,
    CellSigLite,
    ExternalEdge,
    FamilyDesc,
    IRModel,
    LabelEvidence,
    NameDesc,
    RangeEdge,
    SheetMatch,
    TableDesc,
    TimeIndexCandidate,
    UserRoot,
)


def canonicalize_model(
    model: IRModel,
    sheet_match: SheetMatch,
    axis_maps: dict[str, AxisMap],
    *,
    is_identity: bool = False,
) -> IRModel:
    """Rewrite a model into the target (B) namespace.

    For model A: rewrites sheet names via sheet_match, coordinates via axis_maps,
    formulas via rho/kappa, and JSON blobs for address-bearing content.

    For model B: call with is_identity=True to skip coordinate transformation
    but still normalize JSON for consistent comparison.

    Args:
        model: The IR model to rewrite.
        sheet_match: Sheet matching result (mu map).
        axis_maps: Per-sheet axis maps (rho/kappa).
        is_identity: If True, skip coordinate transformation (for B-side normalization).

    Returns:
        New IRModel in the target namespace (or the input model unchanged
        when is_identity=True).
    """
    if is_identity:
        # The identity path rebuilds every cell/binding/edge with unchanged
        # content (all rewrites are skipped and JSON is kept verbatim), so
        # the input model IS the canonical form. Nothing downstream mutates
        # the models; return it directly instead of deep-copying ~1M objects.
        return model

    out = IRModel()
    out.metadata = dict(model.metadata)

    # Build aggregate maps for formula/JSON rewriting
    global_sheet_map = dict(sheet_match.matched)
    global_row_map: dict[int, int | None] = {}
    global_col_map: dict[int, int | None] = {}

    if not is_identity:
        for sa, axis in axis_maps.items():
            global_row_map.update(axis.row_map)
            global_col_map.update(axis.col_map)

    # rewrite_formula_r1c1 is a pure function of (formula, sheet_map,
    # row_map, col_map), and formulas repeat massively across cells (one
    # R1C1 pattern per column/block). Memoize per (sheet, formula) — the
    # axis maps are fixed per sheet for the duration of this call.
    _rewrite_cache: dict[tuple[str, str], str | None] = {}

    def _rewrite_cached(formula: str, sheet: str, sheet_axis) -> str | None:
        key = (sheet, formula)
        if key in _rewrite_cache:
            return _rewrite_cache[key]
        result = rewrite_formula_r1c1(
            formula,
            global_sheet_map,
            sheet_axis.row_map,
            sheet_axis.col_map,
        )
        _rewrite_cache[key] = result
        return result

    # --- Sheets ---
    for s in model.sheet_names:
        new_name = global_sheet_map.get(s, s) if not is_identity else s
        if new_name and (is_identity or s in sheet_match.matched):
            out.sheet_names.append(new_name)
    out.sheet_names.sort()

    # --- Roots ---
    for root in model.roots:
        new_sheet = global_sheet_map.get(root.sheet, root.sheet) if not is_identity else root.sheet
        out.roots.append(
            UserRoot(
                sheet=new_sheet,
                range_a1=root.range_a1,
                label_hint=root.label_hint,
            )
        )

    # --- Cells ---
    cell_key_map: dict[CellKey, CellKey] = {}  # old key -> new key

    for ck, csig in model.cells.items():
        new_key = _transform_cell_key(ck, global_sheet_map, axis_maps, is_identity)
        if new_key is None:
            continue  # Cell on unmatched sheet or deleted row/col

        cell_key_map[ck] = new_key

        # Rewrite formula
        new_formula = csig.formula_r1c1
        if not is_identity and csig.formula_r1c1:
            sheet_axis = axis_maps.get(ck.sheet)
            if sheet_axis:
                new_formula = _rewrite_cached(csig.formula_r1c1, ck.sheet, sheet_axis)

        # Rewrite spill origin
        new_spill = csig.spill_origin
        if not is_identity and csig.spill_origin:
            new_spill = _transform_cell_key(csig.spill_origin, global_sheet_map, axis_maps, False)

        out.cells[new_key] = CellSig(
            formula_r1c1=new_formula,
            value_sha256=csig.value_sha256,
            value_json=csig.value_json,
            format_sha256=csig.format_sha256,
            format_json=csig.format_json,
            data_type=csig.data_type,
            is_array_formula=csig.is_array_formula,
            is_spilled=csig.is_spilled,
            spill_origin=new_spill,
        )

    # --- Bindings ---
    binding_key_map: dict[BindingKey, BindingKey] = {}

    for bk, bdesc in model.bindings.items():
        new_bk = _transform_binding_key(bk, global_sheet_map, axis_maps, is_identity)
        if new_bk is None:
            continue
        binding_key_map[bk] = new_bk

        new_formula = bdesc.formula_r1c1
        if not is_identity and bdesc.formula_r1c1:
            sheet_axis = axis_maps.get(bk.sheet)
            if sheet_axis:
                new_formula = _rewrite_cached(bdesc.formula_r1c1, bk.sheet, sheet_axis)

        # Rebuild members_by_offset with lite sigs (offsets stay the same)
        new_members = {}
        for (dr, dc), lite in bdesc.members_by_offset.items():
            new_lite_formula = lite.formula_r1c1
            if not is_identity and lite.formula_r1c1:
                sheet_axis = axis_maps.get(bk.sheet)
                if sheet_axis:
                    new_lite_formula = _rewrite_cached(lite.formula_r1c1, bk.sheet, sheet_axis)
            new_members[(dr, dc)] = CellSigLite(
                formula_r1c1=new_lite_formula,
                data_type=lite.data_type,
                is_array_formula=lite.is_array_formula,
                is_spilled=lite.is_spilled,
            )

        new_evidence = bdesc.evidence_json
        new_spatial = bdesc.spatial_json
        if not is_identity:
            new_evidence = rewrite_json_blob(bdesc.evidence_json, global_sheet_map, {}, {})
            new_spatial = rewrite_json_blob(bdesc.spatial_json, global_sheet_map, {}, {})

        out.bindings[new_bk] = BindingDesc(
            key=new_bk,
            binding_type=bdesc.binding_type,
            formula_r1c1=new_formula,
            label=bdesc.label,
            classification=bdesc.classification,
            confidence=bdesc.confidence,
            is_orphan=bdesc.is_orphan,
            extraction_source=bdesc.extraction_source,
            evidence_sha256=bdesc.evidence_sha256,
            evidence_json=new_evidence if not is_identity else bdesc.evidence_json,
            spatial_sha256=bdesc.spatial_sha256,
            spatial_json=new_spatial if not is_identity else bdesc.spatial_json,
            address_a1=bdesc.address_a1,
            members_by_offset=new_members,
            original_binding_id=bdesc.original_binding_id,
        )

    # --- Cell-to-binding ---
    for ck, bkeys in model.cell_to_binding.items():
        new_ck = cell_key_map.get(ck)
        if new_ck is None:
            continue
        new_bkeys = [binding_key_map[bk] for bk in bkeys if bk in binding_key_map]
        if new_bkeys:
            out.cell_to_binding[new_ck] = new_bkeys

    # --- Edges ---
    for edge in model.cell_edges:
        new_from = cell_key_map.get(edge.from_key)
        new_to = cell_key_map.get(edge.to_key)
        if new_from and new_to:
            out.cell_edges.add(CellEdge(new_from, new_to))

    for edge in model.external_edges:
        new_from = cell_key_map.get(edge.from_key)
        if new_from:
            out.external_edges.add(ExternalEdge(new_from, edge.external_ref))

    for edge in model.range_edges:
        new_from = cell_key_map.get(edge.from_key)
        new_sheet = (
            global_sheet_map.get(edge.to_sheet, edge.to_sheet) if not is_identity else edge.to_sheet
        )
        if new_from:
            out.range_edges.add(
                RangeEdge(
                    from_key=new_from,
                    to_sheet=new_sheet,
                    to_r1=edge.to_r1,
                    to_c1=edge.to_c1,
                    to_r2=edge.to_r2,
                    to_c2=edge.to_c2,
                    to_range_a1=edge.to_range_a1,
                    cell_count=edge.cell_count,
                )
            )

    for edge in model.binding_edges:
        new_from = binding_key_map.get(edge.from_key)
        new_to = binding_key_map.get(edge.to_key)
        if new_from and new_to:
            out.binding_edges.add(BindingEdgeTuple(new_from, new_to, edge.edge_count))

    # --- Names ---
    for nk, nd in model.names.items():
        new_scope = nd.scope
        if not is_identity and nd.scope != "workbook":
            new_scope = global_sheet_map.get(nd.scope, nd.scope)
        new_dests = (
            canonicalize_destinations_json(nd.destinations, global_sheet_map)
            if not is_identity
            else nd.destinations
        )
        out.names[(nd.name, new_scope)] = NameDesc(
            name=nd.name,
            scope=new_scope,
            destinations=new_dests,
            is_external=nd.is_external,
        )

    # --- Tables ---
    for tid, td in model.tables.items():
        new_sheet = global_sheet_map.get(td.sheet, td.sheet) if not is_identity else td.sheet
        out.tables[tid] = TableDesc(
            candidate_id=td.candidate_id,
            sheet=new_sheet,
            kind=td.kind,
            r1=td.r1,
            c1=td.c1,
            r2=td.r2,
            c2=td.c2,
            range_a1=td.range_a1,
            confidence=td.confidence,
            reasons_json=td.reasons_json,
            members=td.members,
        )

    # --- Families ---
    for fk, fd in model.families.items():
        new_sheet = global_sheet_map.get(fd.sheet, fd.sheet) if not is_identity else fd.sheet
        new_rep = binding_key_map.get(fd.representative_binding_key, fd.representative_binding_key)
        new_members = tuple(binding_key_map.get(mk, mk) for mk in fd.member_binding_keys)
        new_fk = (new_sheet, fd.formula_r1c1)
        out.families[new_fk] = FamilyDesc(
            sheet=new_sheet,
            formula_r1c1=fd.formula_r1c1,
            member_count=fd.member_count,
            representative_binding_key=new_rep,
            member_binding_keys=new_members,
            original_family_id=fd.original_family_id,
        )

    # --- Label evidence ---
    for le in model.label_evidence:
        new_bk = binding_key_map.get(le.binding_key)
        if new_bk:
            new_sheet = global_sheet_map.get(le.sheet, le.sheet) if not is_identity else le.sheet
            out.label_evidence.add(
                LabelEvidence(
                    binding_key=new_bk,
                    candidate_type=le.candidate_type,
                    candidate_address=le.candidate_address,
                    cell_address=le.cell_address,
                    sheet=new_sheet,
                    row=le.row,
                    col=le.col,
                    value_text=le.value_text,
                )
            )

    # --- Time annotations ---
    for tic in model.time_index_candidates:
        new_bk = binding_key_map.get(tic.binding_key)
        new_sheet = global_sheet_map.get(tic.sheet, tic.sheet) if not is_identity else tic.sheet
        if new_bk:
            out.time_index_candidates.append(
                TimeIndexCandidate(
                    sheet=new_sheet,
                    binding_key=new_bk,
                    rank=tic.rank,
                    confidence=tic.confidence,
                    reasons_json=tic.reasons_json,
                )
            )

    for bk, bta in model.binding_time_annotations.items():
        new_bk = binding_key_map.get(bk)
        new_ti_bk = binding_key_map.get(bta.time_index_binding_key)
        if new_bk and new_ti_bk:
            out.binding_time_annotations[new_bk] = BindingTimeAnnotation(
                binding_key=new_bk,
                time_index_binding_key=new_ti_bk,
                is_time_dependent=bta.is_time_dependent,
                confidence=bta.confidence,
                reasons_json=bta.reasons_json,
                evidence_flags_json=bta.evidence_flags_json,
            )

    # --- Resolution metrics (no coordinate rewriting needed) ---
    out.resolution_metrics = dict(model.resolution_metrics)

    return out


def _transform_cell_key(
    ck: CellKey,
    sheet_map: dict[str, str],
    axis_maps: dict[str, AxisMap],
    is_identity: bool,
) -> CellKey | None:
    """Transform a cell key through sheet/axis maps."""
    if is_identity:
        return ck

    new_sheet = sheet_map.get(ck.sheet)
    if new_sheet is None:
        return None  # Sheet not matched

    axis = axis_maps.get(ck.sheet)
    if axis is None:
        return CellKey(new_sheet, ck.row, ck.col)

    new_row = axis.row_map.get(ck.row)
    new_col = axis.col_map.get(ck.col)

    if new_row is None or new_col is None:
        return None  # Row or col deleted

    return CellKey(new_sheet, new_row, new_col)


def _transform_binding_key(
    bk: BindingKey,
    sheet_map: dict[str, str],
    axis_maps: dict[str, AxisMap],
    is_identity: bool,
) -> BindingKey | None:
    """Transform a binding key through sheet/axis maps."""
    if is_identity:
        return bk

    new_sheet = sheet_map.get(bk.sheet)
    if new_sheet is None:
        return None

    axis = axis_maps.get(bk.sheet)
    if axis is None:
        return BindingKey(new_sheet, bk.top_left_row, bk.top_left_col, bk.shape_rows, bk.shape_cols)

    new_row = axis.row_map.get(bk.top_left_row)
    new_col = axis.col_map.get(bk.top_left_col)

    if new_row is None or new_col is None:
        return None

    return BindingKey(new_sheet, new_row, new_col, bk.shape_rows, bk.shape_cols)
