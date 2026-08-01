# ABOUTME: Top-level IR diff pipeline entry point.
# ABOUTME: Wires all stages together into diff_ir(db_a, db_b) -> dict.

from __future__ import annotations

import re

from xl_marinade.core.ir_diff import change_types as CT
from xl_marinade.core.ir_diff.axis_detect import detect_axis_edits
from xl_marinade.core.ir_diff.binding_match import match_bindings
from xl_marinade.core.ir_diff.canonicalize import canonicalize_model
from xl_marinade.core.ir_diff.cell_match import match_cells
from xl_marinade.core.ir_diff.diff_emit import (
    diff_axes,
    diff_bindings,
    diff_cells,
    diff_edges,
    diff_families,
    diff_label_evidence,
    diff_metadata,
    diff_names,
    diff_resolution_metrics,
    diff_roots,
    diff_sheets,
    diff_tables,
    diff_time_annotations,
    normalize_sheet_quoting,
)
from xl_marinade.core.ir_diff.loader import load_model
from xl_marinade.core.ir_diff.model import BindingMatch, Change, IRModel
from xl_marinade.core.ir_diff.ordering import order_changes
from xl_marinade.core.ir_diff.sheet_match import match_sheets
from xl_marinade.core.ir_diff.verify import verify_diff

# Cache-invalidation version for the diff algorithm AND its inputs.
#
# Every persistent diff cache (L2 Redis in app/services/diff_service.py,
# L3 disk in documentation_agent/reasoning/executor.py) folds this number
# into its key/payload, so bumping it here invalidates ALL cached diffs in
# one stroke. Bump whenever a change would alter diff_ir() output for an
# unchanged pair of databases — this includes:
#   - the diff algorithm itself (matching, canonicalization, emit, ordering);
#   - the executor's post-diff enrichment shape or logic;
#   - a change to how IR is produced that diff_ir reads (e.g. the Cycle-17
#     binding-label backfill — diff_ir pairs bindings BY label, so relabelling
#     a workbook changes every diff that touches it).
# The L1 in-memory dict is NOT versioned: it is only ever populated from a
# fresh compute or a version-matched persistent hit, and is cleared on process
# restart (which any code change requires), so it cannot serve a stale diff.
#
# History: v4 = Cycle-17 #350 disk payload widened with value_changes.
# v5 = Cycle-17 #429 — unify disk+Redis under this single knob and invalidate
# the stale-by-label caches left by the cb1bb1aa label backfill.
# v6 = Cycle-17 #430 — numeric-aware VALUE_CHANGED comparison (diff_emit.py):
# serialization-only float differences no longer emit as changes.
# v7 = Cycle-17 #347/VB-053 — _classify_cached_value_continuity now expands
# RANGE member_addresses (e.g. 'F110:EM110') before testing membership in the
# individual-cell value_changes set, so row-bindings whose cells changed value
# are correctly tagged 'changed' (methodology shift) instead of 'same'. Pure
# enrichment-logic change; stale enrich caches carry the wrong continuity tag.
IR_DIFF_CACHE_VERSION = 7


def diff_ir(db_a: str, db_b: str) -> dict:
    """Compare two IR databases and return a replay-complete changelist.

    Deterministic 5-stage pipeline:
    1. Load and normalize raw IR rows.
    2. Match sheets and derive axis maps.
    3. Canonicalize A into B's namespace.
    4. Match bindings and cells on canonicalized space.
    5. Emit, order, and verify changes.

    Args:
        db_a: Path to Version A (earlier) IR database.
        db_b: Path to Version B (later) IR database.

    Returns:
        Dict with version, summary, and ordered changes list.

    Raises:
        DiffVerificationError: If the verification pass fails.
    """
    # Stage 1: Load (parallel — independent reads from separate SQLite files)
    from concurrent.futures import ThreadPoolExecutor

    with ThreadPoolExecutor(max_workers=2) as pool:
        future_a = pool.submit(load_model, db_a)
        future_b = pool.submit(load_model, db_b)
        a = future_a.result()
        b = future_b.result()

    # Stage 2: Match sheets and detect axis edits
    sheet_match = match_sheets(a, b)
    axis_maps = detect_axis_edits(a, b, sheet_match)

    # Stage 3: Canonicalize A into B's namespace
    a_norm = canonicalize_model(a, sheet_match, axis_maps, is_identity=False)
    b_norm = canonicalize_model(b, sheet_match, {}, is_identity=True)

    # Stage 4: Match bindings and cells
    binding_match = match_bindings(a_norm, b_norm)
    cell_match = match_cells(a_norm, b_norm, binding_match)

    # Stage 5: Emit changes
    changes: list[Change] = []

    changes += diff_metadata(a_norm, b_norm)
    root_changes = diff_roots(a_norm, b_norm)
    changes += root_changes
    scope_mode = any(c.type == CT.ROOT_CHANGED for c in root_changes)

    changes += diff_sheets(sheet_match)
    changes += diff_axes(axis_maps, sheet_match)
    changes += diff_names(a_norm, b_norm)
    changes += diff_tables(a_norm, b_norm)
    changes += diff_bindings(a_norm, b_norm, binding_match, scope_mode)
    changes += diff_cells(a_norm, b_norm, cell_match, binding_match)
    changes += diff_label_evidence(a_norm, b_norm)
    changes += diff_time_annotations(a_norm, b_norm)
    changes += diff_edges(a_norm, b_norm, cell_match, binding_match)
    changes += diff_families(a_norm, b_norm, binding_match)
    changes += diff_resolution_metrics(a_norm, b_norm)

    # Order deterministically
    changes = order_changes(changes)

    # Verify
    verify_diff(a_norm, b_norm, changes, binding_match, cell_match)

    # Build output
    return _build_output(a, b, changes, binding_match, a_norm, b_norm)


def _classify_formula_change(old_f: str | None, new_f: str | None) -> str:
    """Classify a formula change as reference_shift or logic_change.

    Replaces absolute row/col numbers in R1C1 refs with wildcards.
    If the patterns match, only absolute targets changed → reference_shift.
    """
    if not old_f or not new_f:
        return "logic_change"
    # Quoting-only spelling differences ('Sheet'! vs Sheet!) are not logic.
    # normalize_sheet_quoting returns None only for a None input; both are
    # non-None here (guarded above), so `or` just re-narrows the type to str.
    old_f = normalize_sheet_quoting(old_f) or old_f
    new_f = normalize_sheet_quoting(new_f) or new_f
    # Replace absolute row/col numbers: R<digits> → R*, C<digits> → C*
    # but preserve relative refs like R[-1] and C[2]
    old_pat = re.sub(r"(?<=[RC])(\d+)", "*", old_f)
    new_pat = re.sub(r"(?<=[RC])(\d+)", "*", new_f)
    return "reference_shift" if old_pat == new_pat else "logic_change"


def _build_output(
    a,
    b,
    changes: list[Change],
    binding_match: BindingMatch,
    a_norm: IRModel,
    b_norm: IRModel,
) -> dict:
    """Build the output dictionary from changes."""
    summary = _build_summary(changes)
    binding_map = _build_binding_map(binding_match, a_norm, b_norm, changes)

    # Build lookup: binding_id → original formula from pre-canonicalization model A.
    # Canonicalization can null out formula_r1c1 when absolute refs point to
    # unmapped rows/cols, so the diff emits old_formula=None. We backfill from
    # the original model so the frontend can display the previous formula.
    a_original_formulas: dict[str, str] = {}
    a_binding_by_id: dict[str, object] = {}
    for bdesc in a.bindings.values():
        if bdesc.original_binding_id:
            a_binding_by_id[bdesc.original_binding_id] = bdesc
            if bdesc.formula_r1c1:
                a_original_formulas[bdesc.original_binding_id] = bdesc.formula_r1c1

    serialized_changes = []
    for i, c in enumerate(changes):
        entry = {"seq": i + 1, "type": c.type, **c.details}
        if c.type in CT.IR_INFERENCE_TYPES:
            entry["layer"] = "ir_inference"
        if c.type == CT.BINDING_FORMULA_CHANGED:
            # Backfill old_formula from original model when canonicalization lost it
            if not entry.get("old_formula") and entry.get("binding_id_a"):
                original = a_original_formulas.get(entry["binding_id_a"])
                if original:
                    entry["old_formula"] = original
                else:
                    # Last resort: extract formula from the binding's member cells.
                    # Some bindings have formula_id=None but their cells carry formulas
                    # (e.g. LET formulas with varying absolute refs).
                    bdesc = a_binding_by_id.get(entry["binding_id_a"])
                    if bdesc and hasattr(bdesc, "members_by_offset"):
                        for sig in bdesc.members_by_offset.values():  # type: ignore[union-attr]
                            if sig.formula_r1c1:
                                entry["old_formula"] = sig.formula_r1c1
                                break
            # Classify: reference_shift (only absolute refs differ) vs logic_change
            entry["modification_kind"] = _classify_formula_change(
                entry.get("old_formula"), entry.get("new_formula")
            )
        serialized_changes.append(entry)

    return {
        "version": "1.0",
        "schema_version_a": a.metadata.get("schema_version", "unknown"),
        "schema_version_b": b.metadata.get("schema_version", "unknown"),
        "workbook_sha256_a": a.metadata.get("workbook_sha256"),
        "workbook_sha256_b": b.metadata.get("workbook_sha256"),
        "root_a": {"sheet": a.roots[0].sheet, "range": a.roots[0].range_a1} if a.roots else None,
        "root_b": {"sheet": b.roots[0].sheet, "range": b.roots[0].range_a1} if b.roots else None,
        "summary": summary,
        "binding_map": binding_map,
        "changes": serialized_changes,
    }


_MODIFIED_CHANGE_TYPES = frozenset(
    {
        CT.BINDING_FORMULA_CHANGED,
        CT.BINDING_TYPE_CHANGED,
        CT.BINDING_RESIZED,
        CT.BINDING_LABEL_CHANGED,
    }
)


def _build_binding_map(
    binding_match: BindingMatch,
    a_norm: IRModel,
    b_norm: IRModel,
    changes: list[Change],
) -> list[dict]:
    """Build a flat binding-level diff map for frontend overlay.

    Each entry maps a binding from version A to version B with its diff state.
    """
    # Index which A binding_ids have binding-level changes and what types
    changed_a_ids: dict[str, set[str]] = {}  # binding_id_a -> set of change types
    for c in changes:
        bid_a = c.details.get("binding_id_a")
        if bid_a and c.type in (
            CT.BINDING_MOVED,
            CT.BINDING_RESIZED,
            CT.BINDING_TYPE_CHANGED,
            CT.BINDING_FORMULA_CHANGED,
            CT.BINDING_LABEL_CHANGED,
            CT.BINDING_METADATA_CHANGED,
        ):
            changed_a_ids.setdefault(bid_a, set()).add(c.type)

    entries: list[dict] = []

    # Matched pairs
    for ak in sorted(binding_match.matched.keys()):
        bk = binding_match.matched[ak]
        ad = a_norm.bindings.get(ak)
        bd = b_norm.bindings.get(bk)
        if not ad or not bd:
            continue

        bid_a = ad.original_binding_id
        change_types = changed_a_ids.get(bid_a, set())

        if CT.BINDING_MOVED in change_types:
            diff_state = "moved"
        elif change_types & _MODIFIED_CHANGE_TYPES:
            diff_state = "modified"
        else:
            diff_state = "unchanged"

        entries.append(
            {
                "binding_id_a": bid_a,
                "binding_id_b": bd.original_binding_id,
                "diff_state": diff_state,
                "address_a": ad.address_a1,
                "address_b": bd.address_a1,
                "label_a": ad.label,
                "label_b": bd.label,
            }
        )

    # Removed (in A only)
    for ak in sorted(binding_match.removed):
        ad = a_norm.bindings.get(ak)
        if not ad:
            continue
        entries.append(
            {
                "binding_id_a": ad.original_binding_id,
                "binding_id_b": None,
                "diff_state": "removed",
                "address_a": ad.address_a1,
                "address_b": None,
                "label_a": ad.label,
                "label_b": None,
            }
        )

    # Added (in B only)
    for bk in sorted(binding_match.added):
        bd = b_norm.bindings.get(bk)
        if not bd:
            continue
        entries.append(
            {
                "binding_id_a": None,
                "binding_id_b": bd.original_binding_id,
                "diff_state": "added",
                "address_a": None,
                "address_b": bd.address_a1,
                "label_a": None,
                "label_b": bd.label,
            }
        )

    return entries


def _build_summary(changes: list[Change]) -> dict:
    """Count changes by category."""
    s = {
        "total_changes": len(changes),
        "sheets_added": 0,
        "sheets_removed": 0,
        "sheets_renamed": 0,
        "rows_inserted": 0,
        "rows_deleted": 0,
        "cols_inserted": 0,
        "cols_deleted": 0,
        "bindings_added": 0,
        "bindings_removed": 0,
        "bindings_scope_in": 0,
        "bindings_scope_out": 0,
        "bindings_moved": 0,
        "bindings_resized": 0,
        "bindings_formula_changed": 0,
        "bindings_type_changed": 0,
        "bindings_label_changed": 0,
        "bindings_metadata_changed": 0,
        "cells_formula_changed": 0,
        "cells_value_changed": 0,
        "cells_format_changed": 0,
        "cells_extras_changed": 0,
        "cells_dtype_changed": 0,
        "cells_array_flag_changed": 0,
        "cells_spill_changed": 0,
        "names_added": 0,
        "names_removed": 0,
        "names_changed": 0,
        "tables_changed": 0,
        "edges_changed": 0,
        "families_changed": 0,
        "time_annotations_changed": 0,
        "label_evidence_changed": 0,
        "resolution_metrics_changed": 0,
        "metadata_changed": 0,
        # Roll-up of CT.IR_INFERENCE_TYPES (tables / label evidence / time
        # annotations): extractor-derived changes, not workbook edits.
        "ir_inference_changes": 0,
    }

    type_to_key = {
        CT.SHEET_ADDED: "sheets_added",
        CT.SHEET_REMOVED: "sheets_removed",
        CT.SHEET_RENAMED: "sheets_renamed",
        CT.ROWS_INSERTED: "rows_inserted",
        CT.ROWS_DELETED: "rows_deleted",
        CT.COLS_INSERTED: "cols_inserted",
        CT.COLS_DELETED: "cols_deleted",
        CT.BINDING_ADDED: "bindings_added",
        CT.BINDING_REMOVED: "bindings_removed",
        CT.BINDING_NOW_IN_SCOPE: "bindings_scope_in",
        CT.BINDING_OUT_OF_SCOPE: "bindings_scope_out",
        CT.BINDING_MOVED: "bindings_moved",
        CT.BINDING_RESIZED: "bindings_resized",
        CT.BINDING_FORMULA_CHANGED: "bindings_formula_changed",
        CT.BINDING_TYPE_CHANGED: "bindings_type_changed",
        CT.BINDING_LABEL_CHANGED: "bindings_label_changed",
        CT.BINDING_METADATA_CHANGED: "bindings_metadata_changed",
        CT.FORMULA_CHANGED: "cells_formula_changed",
        CT.VALUE_CHANGED: "cells_value_changed",
        CT.FORMAT_CHANGED: "cells_format_changed",
        CT.EXTRAS_CHANGED: "cells_extras_changed",
        CT.DTYPE_CHANGED: "cells_dtype_changed",
        CT.ARRAY_FORMULA_FLAG_CHANGED: "cells_array_flag_changed",
        CT.SPILL_CHANGED: "cells_spill_changed",
        CT.NAME_ADDED: "names_added",
        CT.NAME_REMOVED: "names_removed",
        CT.NAME_DESTINATIONS_CHANGED: "names_changed",
        CT.NAME_METADATA_CHANGED: "names_changed",
        CT.BINDING_LABEL_EVIDENCE_CHANGED: "label_evidence_changed",
        CT.IR_METADATA_CHANGED: "metadata_changed",
    }

    edge_types = {
        CT.CELL_EDGE_ADDED,
        CT.CELL_EDGE_REMOVED,
        CT.RANGE_EDGE_ADDED,
        CT.RANGE_EDGE_REMOVED,
        CT.RANGE_EDGE_COUNT_CHANGED,
        CT.EXTERNAL_EDGE_ADDED,
        CT.EXTERNAL_EDGE_REMOVED,
        CT.BINDING_EDGE_ADDED,
        CT.BINDING_EDGE_REMOVED,
        CT.BINDING_EDGE_WEIGHT_CHANGED,
    }

    table_types = {CT.TABLE_CANDIDATE_ADDED, CT.TABLE_CANDIDATE_REMOVED, CT.TABLE_CANDIDATE_CHANGED}

    family_types = {
        CT.FAMILY_ADDED,
        CT.FAMILY_REMOVED,
        CT.FAMILY_RESIZED,
        CT.FAMILY_REPRESENTATIVE_CHANGED,
        CT.FAMILY_FORMULA_CHANGED,
    }

    time_types = {
        CT.TIME_INDEX_CANDIDATE_ADDED,
        CT.TIME_INDEX_CANDIDATE_REMOVED,
        CT.TIME_INDEX_CANDIDATE_CHANGED,
        CT.BINDING_TIME_ANNOTATION_ADDED,
        CT.BINDING_TIME_ANNOTATION_REMOVED,
        CT.BINDING_TIME_ANNOTATION_CHANGED,
    }

    metric_types = {
        CT.RESOLUTION_METRIC_ADDED,
        CT.RESOLUTION_METRIC_REMOVED,
        CT.RESOLUTION_METRIC_CHANGED,
    }

    for c in changes:
        if c.type in type_to_key:
            s[type_to_key[c.type]] += 1
        elif c.type in edge_types:
            s["edges_changed"] += 1
        elif c.type in table_types:
            s["tables_changed"] += 1
        elif c.type in family_types:
            s["families_changed"] += 1
        elif c.type in time_types:
            s["time_annotations_changed"] += 1
        elif c.type in metric_types:
            s["resolution_metrics_changed"] += 1

    s["ir_inference_changes"] = sum(1 for c in changes if c.type in CT.IR_INFERENCE_TYPES)

    return s
