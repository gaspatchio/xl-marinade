# ABOUTME: Stage 1 — Load an IR database into an immutable in-memory IRModel.
# ABOUTME: Resolves all database-local IDs to textual canonical structures.

from __future__ import annotations

import sqlite3

from xl_marinade.core.db_uri import connect_read_only
from xl_marinade.core.ir_diff.model import (
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
    TableDesc,
    TimeIndexCandidate,
    UserRoot,
)


def _table_exists(conn: sqlite3.Connection, name: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone()
    return row is not None


def load_model(db_path: str) -> IRModel:
    """Load an IR database into an in-memory IRModel.

    All database-local IDs (sheet_id, formula_id, blob_id, cell_id) are resolved
    to textual canonical structures. The returned model uses only CellKey, BindingKey,
    and string-based identifiers.

    Args:
        db_path: Path to the IR SQLite database.

    Returns:
        Populated IRModel instance.
    """
    conn = connect_read_only(db_path)
    conn.row_factory = sqlite3.Row
    try:
        return _load(conn)
    finally:
        conn.close()


def _load(conn: sqlite3.Connection) -> IRModel:
    model = IRModel()

    # --- Metadata ---
    if _table_exists(conn, "ir_metadata"):
        for row in conn.execute("SELECT key, value FROM ir_metadata ORDER BY key"):
            model.metadata[row["key"]] = row["value"]

    # --- Roots ---
    for row in conn.execute("SELECT sheet, range_a1, label_hint FROM user_roots ORDER BY root_id"):
        model.roots.append(
            UserRoot(
                sheet=row["sheet"],
                range_a1=row["range_a1"],
                label_hint=row["label_hint"],
            )
        )

    # --- Sheets ---
    sheet_map: dict[int, str] = {}  # sheet_id -> sheet_name
    for row in conn.execute("SELECT sheet_id, sheet_name FROM sheets ORDER BY sheet_name"):
        sheet_map[row["sheet_id"]] = row["sheet_name"]
        model.sheet_names.append(row["sheet_name"])

    # --- Formulas ---
    formula_map: dict[int, tuple[str, str | None]] = {}  # formula_id -> (r1c1, a1_example)
    for row in conn.execute("SELECT formula_id, formula_r1c1, formula_a1_example FROM formulas"):
        formula_map[row["formula_id"]] = (row["formula_r1c1"], row["formula_a1_example"])

    # --- JSON blobs ---
    blob_map: dict[int, tuple[str, str]] = {}  # blob_id -> (sha256, json)
    for row in conn.execute("SELECT blob_id, sha256, json FROM json_blobs"):
        blob_map[row["blob_id"]] = (row["sha256"], row["json"])

    # --- Cells ---
    # We need cell_id -> CellKey for edge resolution
    cell_id_to_key: dict[int, CellKey] = {}

    for row in conn.execute("""
        SELECT
            c.cell_id, c.sheet_id, c.row, c.col,
            c.formula_id, c.value_blob_id, c.format_blob_id,
            c.data_type, c.is_array_formula, c.is_spilled, c.spilled_from_cell_id
        FROM cells c
        ORDER BY c.sheet_id, c.row, c.col
    """):
        sheet_name = sheet_map.get(row["sheet_id"], f"__unknown_sheet_{row['sheet_id']}")
        key = CellKey(sheet=sheet_name, row=row["row"], col=row["col"])
        cell_id_to_key[row["cell_id"]] = key

        # Resolve formula
        formula_r1c1 = None
        if row["formula_id"] is not None and row["formula_id"] in formula_map:
            formula_r1c1 = formula_map[row["formula_id"]][0]

        # Resolve blobs
        val_sha, val_json = blob_map.get(row["value_blob_id"] or -1, (None, None))  # type: ignore[arg-type]
        fmt_sha, fmt_json = blob_map.get(row["format_blob_id"] or -1, (None, None))  # type: ignore[arg-type]

        sig = CellSig(
            formula_r1c1=formula_r1c1,
            value_sha256=val_sha,
            value_json=val_json,
            format_sha256=fmt_sha,
            format_json=fmt_json,
            data_type=row["data_type"],
            is_array_formula=bool(row["is_array_formula"]),
            is_spilled=bool(row["is_spilled"]),
            spill_origin=None,  # resolved in second pass
        )
        model.cells[key] = sig

    # Second pass: resolve spill origins
    for row in conn.execute("""
        SELECT cell_id, spilled_from_cell_id
        FROM cells
        WHERE spilled_from_cell_id IS NOT NULL
    """):
        key = cell_id_to_key.get(row["cell_id"])
        origin_key = cell_id_to_key.get(row["spilled_from_cell_id"])
        if key and key in model.cells:
            old = model.cells[key]
            model.cells[key] = CellSig(
                formula_r1c1=old.formula_r1c1,
                value_sha256=old.value_sha256,
                value_json=old.value_json,
                format_sha256=old.format_sha256,
                format_json=old.format_json,
                data_type=old.data_type,
                is_array_formula=old.is_array_formula,
                is_spilled=old.is_spilled,
                spill_origin=origin_key,
            )

    # --- Bindings ---
    binding_id_to_key: dict[str, BindingKey] = {}

    for row in conn.execute("""
        SELECT
            b.binding_id, b.sheet_id, b.address_a1,
            tl.row AS top_left_row, tl.col AS top_left_col,
            b.shape_rows, b.shape_cols, b.binding_type,
            b.formula_id, b.label, b.classification, b.confidence,
            b.is_orphan, b.extraction_source,
            b.evidence_blob_id, b.spatial_candidates_blob_id
        FROM bindings b
        JOIN cells tl ON tl.cell_id = b.top_left_cell_id
        ORDER BY tl.row, tl.col, b.binding_id
    """):
        sheet_name = sheet_map.get(row["sheet_id"], f"__unknown_sheet_{row['sheet_id']}")
        bkey = BindingKey(
            sheet=sheet_name,
            top_left_row=row["top_left_row"],
            top_left_col=row["top_left_col"],
            shape_rows=row["shape_rows"],
            shape_cols=row["shape_cols"],
        )
        binding_id_to_key[row["binding_id"]] = bkey

        formula_r1c1 = None
        if row["formula_id"] is not None and row["formula_id"] in formula_map:
            formula_r1c1 = formula_map[row["formula_id"]][0]

        ev_sha, ev_json = blob_map.get(row["evidence_blob_id"] or -1, (None, None))  # type: ignore[arg-type]
        sp_sha, sp_json = blob_map.get(row["spatial_candidates_blob_id"] or -1, (None, None))  # type: ignore[arg-type]

        desc = BindingDesc(
            key=bkey,
            binding_type=row["binding_type"],
            formula_r1c1=formula_r1c1,
            label=row["label"],
            classification=row["classification"],
            confidence=row["confidence"],
            is_orphan=bool(row["is_orphan"]),
            extraction_source=row["extraction_source"],
            evidence_sha256=ev_sha,
            evidence_json=ev_json,
            spatial_sha256=sp_sha,
            spatial_json=sp_json,
            address_a1=row["address_a1"],
            original_binding_id=row["binding_id"],
        )
        model.bindings[bkey] = desc

    # --- Cell-to-binding membership ---
    for row in conn.execute("""
        SELECT ctb.cell_id, ctb.binding_id
        FROM cell_to_binding ctb
        ORDER BY ctb.binding_id, ctb.cell_id
    """):
        cell_key = cell_id_to_key.get(row["cell_id"])
        binding_key = binding_id_to_key.get(row["binding_id"])
        if cell_key and binding_key:
            model.cell_to_binding.setdefault(cell_key, []).append(binding_key)

    # Build members_by_offset for each binding
    for cell_key, bkeys in model.cell_to_binding.items():
        cell_sig = model.cells.get(cell_key)
        if not cell_sig:
            continue
        lite = CellSigLite(
            formula_r1c1=cell_sig.formula_r1c1,
            data_type=cell_sig.data_type,
            is_array_formula=cell_sig.is_array_formula,
            is_spilled=cell_sig.is_spilled,
        )
        for bkey in bkeys:
            bdesc = model.bindings.get(bkey)
            if bdesc:
                dr = cell_key.row - bkey.top_left_row
                dc = cell_key.col - bkey.top_left_col
                # BindingDesc.members_by_offset is a mutable dict (default_factory)
                bdesc.members_by_offset[(dr, dc)] = lite

    # --- Internal cell edges ---
    for row in conn.execute("""
        SELECT from_cell_id, to_cell_id FROM cell_edges_internal
    """):
        fk = cell_id_to_key.get(row["from_cell_id"])
        tk = cell_id_to_key.get(row["to_cell_id"])
        if fk and tk:
            model.cell_edges.add(CellEdge(from_key=fk, to_key=tk))

    # --- External edges ---
    for row in conn.execute("""
        SELECT from_cell_id, external_ref FROM cell_edges_external
    """):
        fk = cell_id_to_key.get(row["from_cell_id"])
        if fk:
            model.external_edges.add(ExternalEdge(from_key=fk, external_ref=row["external_ref"]))

    # --- Range edges ---
    for row in conn.execute("""
        SELECT
            from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2,
            to_range_a1, cell_count
        FROM range_edges
    """):
        fk = cell_id_to_key.get(row["from_cell_id"])
        to_sheet = sheet_map.get(row["to_sheet_id"], f"__unknown_sheet_{row['to_sheet_id']}")
        if fk:
            model.range_edges.add(
                RangeEdge(
                    from_key=fk,
                    to_sheet=to_sheet,
                    to_r1=row["to_r1"],
                    to_c1=row["to_c1"],
                    to_r2=row["to_r2"],
                    to_c2=row["to_c2"],
                    to_range_a1=row["to_range_a1"],
                    cell_count=row["cell_count"],
                )
            )

    # --- Binding edges ---
    for row in conn.execute("""
        SELECT from_binding_id, to_binding_id, edge_count FROM binding_edges
    """):
        fk = binding_id_to_key.get(row["from_binding_id"])
        tk = binding_id_to_key.get(row["to_binding_id"])
        if fk and tk:
            model.binding_edges.add(
                BindingEdgeTuple(
                    from_key=fk,
                    to_key=tk,
                    edge_count=row["edge_count"],
                )
            )

    # --- Defined names ---
    for row in conn.execute("""
        SELECT name, scope, destinations, is_external
        FROM defined_names
        ORDER BY name, scope
    """):
        nk = (row["name"], row["scope"])
        model.names[nk] = NameDesc(
            name=row["name"],
            scope=row["scope"],
            destinations=row["destinations"],
            is_external=bool(row["is_external"]),
        )

    # --- Table candidates ---
    for row in conn.execute("""
        SELECT
            tc.candidate_id, s.sheet_name, tc.kind,
            tc.r1, tc.c1, tc.r2, tc.c2, tc.range_a1,
            tc.confidence, tc.reasons_top3_json
        FROM table_candidates tc
        JOIN sheets s ON s.sheet_id = tc.sheet_id
        ORDER BY s.sheet_name, tc.r1, tc.c1, tc.candidate_id
    """):
        # Load members for this candidate
        members = []
        for mrow in conn.execute(
            """
            SELECT ordinal, binding_id, role_hint
            FROM table_candidate_members
            WHERE candidate_id = ?
            ORDER BY ordinal
        """,
            (row["candidate_id"],),
        ):
            bk_repr = mrow["binding_id"]  # original ID, will be resolved during matching
            members.append((mrow["ordinal"], bk_repr, mrow["role_hint"]))

        model.tables[row["candidate_id"]] = TableDesc(
            candidate_id=row["candidate_id"],
            sheet=row["sheet_name"],
            kind=row["kind"],
            r1=row["r1"],
            c1=row["c1"],
            r2=row["r2"],
            c2=row["c2"],
            range_a1=row["range_a1"],
            confidence=row["confidence"],
            reasons_json=row["reasons_top3_json"],
            members=tuple(members),
        )

    # --- Formula families ---
    for row in conn.execute("""
        SELECT
            ff.family_id, s.sheet_name, f.formula_r1c1,
            ff.member_count, ff.representative_binding_id
        FROM formula_families ff
        JOIN sheets s ON s.sheet_id = ff.sheet_id
        JOIN formulas f ON f.formula_id = ff.formula_id
        ORDER BY s.sheet_name, f.formula_r1c1
    """):
        rep_key = binding_id_to_key.get(row["representative_binding_id"])
        fam_key = (row["sheet_name"], row["formula_r1c1"])

        # Load member binding keys
        member_keys = []
        for mrow in conn.execute(
            """
            SELECT binding_id FROM formula_family_members
            WHERE family_id = ?
            ORDER BY ordinal
        """,
            (row["family_id"],),
        ):
            mk = binding_id_to_key.get(mrow["binding_id"])
            if mk:
                member_keys.append(mk)

        if rep_key:
            model.families[fam_key] = FamilyDesc(
                sheet=row["sheet_name"],
                formula_r1c1=row["formula_r1c1"],
                member_count=row["member_count"],
                representative_binding_key=rep_key,
                member_binding_keys=tuple(member_keys),
                original_family_id=row["family_id"],
            )

    # --- Label evidence ---
    if _table_exists(conn, "binding_label_candidate_cells"):
        for row in conn.execute("""
            SELECT
                bl.binding_id, s.sheet_name,
                bl.candidate_type, bl.candidate_address,
                bl.cell_address, bl.row, bl.col, bl.value_text
            FROM binding_label_candidate_cells bl
            JOIN sheets s ON s.sheet_id = bl.sheet_id
            ORDER BY bl.binding_id, bl.candidate_type, bl.candidate_address, bl.cell_address
        """):
            bkey = binding_id_to_key.get(row["binding_id"])
            if bkey:
                model.label_evidence.add(
                    LabelEvidence(
                        binding_key=bkey,
                        candidate_type=row["candidate_type"],
                        candidate_address=row["candidate_address"],
                        cell_address=row["cell_address"],
                        sheet=row["sheet_name"],
                        row=row["row"],
                        col=row["col"],
                        value_text=row["value_text"],
                    )
                )

    # --- Time index candidates ---
    if _table_exists(conn, "time_index_candidates"):
        for row in conn.execute("""
            SELECT s.sheet_name, tic.binding_id, tic.rank, tic.confidence, tic.reasons_top3_json
            FROM time_index_candidates tic
            JOIN sheets s ON s.sheet_id = tic.sheet_id
            ORDER BY s.sheet_name, tic.rank
        """):
            bkey = binding_id_to_key.get(row["binding_id"])
            if bkey:
                model.time_index_candidates.append(
                    TimeIndexCandidate(
                        sheet=row["sheet_name"],
                        binding_key=bkey,
                        rank=row["rank"],
                        confidence=row["confidence"],
                        reasons_json=row["reasons_top3_json"],
                    )
                )

    # --- Binding time annotations ---
    if _table_exists(conn, "binding_time_annotations"):
        for row in conn.execute("""
            SELECT
                binding_id, time_index_binding_id, is_time_dependent,
                confidence, reasons_top3_json, evidence_flags_json
            FROM binding_time_annotations
            ORDER BY binding_id
        """):
            bkey = binding_id_to_key.get(row["binding_id"])
            ti_key = binding_id_to_key.get(row["time_index_binding_id"])
            if bkey and ti_key:
                model.binding_time_annotations[bkey] = BindingTimeAnnotation(
                    binding_key=bkey,
                    time_index_binding_key=ti_key,
                    is_time_dependent=bool(row["is_time_dependent"]),
                    confidence=row["confidence"],
                    reasons_json=row["reasons_top3_json"],
                    evidence_flags_json=row["evidence_flags_json"],
                )

    # --- Resolution metrics ---
    if _table_exists(conn, "resolution_metrics"):
        for row in conn.execute("""
            SELECT function_name, status, count
            FROM resolution_metrics
            ORDER BY function_name, status
        """):
            model.resolution_metrics[(row["function_name"], row["status"])] = row["count"]

    # --- Resolve missing labels from label evidence (scan_above) ---
    # BindingDesc is frozen, so we replace entries whose label is None with a
    # copy that carries the scan_above candidate, so the diff binding_map
    # carries human-readable names for both versions.
    _resolve_missing_labels(model)

    return model


def _resolve_missing_labels(model: IRModel) -> None:
    """Fill in BindingDesc.label for bindings that lack an explicit label.

    Uses scan_above label evidence (the same fallback tier the extractor's
    label cascade uses) so the diff binding_map carries human-readable names.
    BindingDesc is frozen, so we replace the dict entry with a new instance.
    """
    if not model.label_evidence:
        return

    # Build a lookup: binding_key → best scan_above text
    scan_above: dict[BindingKey, str] = {}
    for ev in model.label_evidence:
        if ev.candidate_type != "scan_above":
            continue
        if not ev.value_text or ev.value_text.startswith("="):
            continue
        # Keep the first candidate per binding (they're ordered by address)
        if ev.binding_key not in scan_above:
            scan_above[ev.binding_key] = ev.value_text

    for bkey, desc in list(model.bindings.items()):
        if desc.label:
            continue
        candidate = scan_above.get(bkey)
        if candidate:
            # Replace frozen dataclass instance with a copy carrying the label
            model.bindings[bkey] = BindingDesc(
                key=desc.key,
                binding_type=desc.binding_type,
                formula_r1c1=desc.formula_r1c1,
                label=candidate,
                classification=desc.classification,
                confidence=desc.confidence,
                is_orphan=desc.is_orphan,
                extraction_source=desc.extraction_source,
                evidence_sha256=desc.evidence_sha256,
                evidence_json=desc.evidence_json,
                spatial_sha256=desc.spatial_sha256,
                spatial_json=desc.spatial_json,
                address_a1=desc.address_a1,
                original_binding_id=desc.original_binding_id,
                members_by_offset=desc.members_by_offset,
            )
