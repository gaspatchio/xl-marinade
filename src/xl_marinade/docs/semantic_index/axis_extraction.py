# ABOUTME: Axis extraction for semantic index - extracts row/col axes with samples and links to Sprint 6 time index.
# ABOUTME: Produces stable axis entities with header addresses, samples, and time index evidence pointers.

import json
import logging
import sqlite3
from typing import Any

from xl_marinade.core.ref_converter import col_num_to_letter, parse_cell_address

logger = logging.getLogger(__name__)

# B1 (Gap 1): single shared constant for the named time axis. Reconciles the old
# 'time_index' literal here with the producer and the json_spec consumer.
TIME_AXIS_NAME = "time"


def extract_axes(
    *,
    ir_db_path: str,
    table: dict[str, Any],
) -> list[dict[str, Any]]:
    """
    Extract axes (row/column) for a detected table.

    Args:
        ir_db_path: Path to IR database
        table: Table dict from detect_tables with keys:
            - sheet: str
            - range_a1: str
            - table_type: str
            - confidence: float
            - evidence: dict with row_axis/column_axis info

    Returns:
        List of axis dicts with keys:
            - orientation: str ("row" or "column")
            - header_range_a1: str (binding range for axis headers)
            - axis_name: str (semantic label)
            - samples_json: str (JSON array of sample values)
            - confidence: float
            - binding_id: str (axis binding ID)
    """
    logger.debug(f"Extracting axes for table: {table['sheet']}!{table['range_a1']}")

    axes = []
    evidence = table.get("evidence", {})

    # Extract row axis
    row_axis_info = evidence.get("row_axis")
    if row_axis_info:
        binding_id = row_axis_info["binding_id"]
        samples = _load_axis_samples(ir_db_path, binding_id, max_samples=5, orientation="row")

        axis = {
            "orientation": "row",
            "header_range_a1": _get_binding_range(ir_db_path, binding_id),
            "axis_name": row_axis_info.get("label", "row_index"),
            "samples_json": json.dumps(samples),
            "confidence": 0.7
            if row_axis_info.get("is_monotonic") or row_axis_info.get("is_unique")
            else 0.5,
            "binding_id": binding_id,
        }
        axes.append(axis)

    # Extract column axis
    col_axis_info = evidence.get("column_axis")
    if col_axis_info:
        binding_id = col_axis_info["binding_id"]
        samples = _load_axis_samples(ir_db_path, binding_id, max_samples=10, orientation="column")

        axis = {
            "orientation": "column",
            "header_range_a1": _get_binding_range(ir_db_path, binding_id),
            "axis_name": col_axis_info.get("label", "column_headers"),
            "samples_json": json.dumps(samples),
            "confidence": 0.6,
            "binding_id": binding_id,
        }
        axes.append(axis)

    # Fallback: infer axes from table range if none were detected
    if not axes:
        inferred = _infer_axes_from_range(ir_db_path, table)
        axes.extend(inferred)

    logger.debug(f"Extracted {len(axes)} axes")

    return axes


def link_time_axis(
    *,
    ir_db_path: str,
    axes: list[dict[str, Any]],
    counters: dict[str, int] | None = None,
) -> list[dict[str, Any]]:
    """
    Link axes to Sprint 6 time index evidence where available.

    Checks if any axis corresponds to a time_index_candidate from Sprint 6
    and adds time_index_binding_id and time_index_range_a1 fields.

    Args:
        ir_db_path: Path to IR database
        axes: List of axis dicts from extract_axes
        counters: Optional mutable dict for build-time observability. When passed,
            these keys are incremented so a silently-severed time link becomes a
            countable regression (Phase 1 de-brittling): 'time_linked' (axis got
            axis_kind='time'), 'drop_shape_le_1' (binding IS time-dependent but a
            1x1 scalar reading the index — deliberately excluded by the shape
            gate), 'drop_not_time_dependent' (binding present but carries no
            is_time_dependent annotation). The 'drop_none_binding_id' gate is
            counted by the caller at the binding_id back-fill site.

    Returns:
        Updated list of axis dicts with time index links added where applicable
    """
    logger.debug("Linking time axis evidence via binding_time_annotations")

    con = sqlite3.connect(ir_db_path)
    con.row_factory = sqlite3.Row

    try:
        # B1 fix for the broken join: the old code matched an axis's HEADER binding_id
        # against the cross-sheet TIME-INDEX binding_id (exact equality) which ~never
        # holds. The real series->axis link lives in binding_time_annotations
        # (binding_id -> time_index_binding_id). Read it, gated by the shape predicate
        # (is_time_dependent=1 AND MAX(shape_rows,shape_cols)>1) so 1x1 scalars that
        # merely read the time index never get tagged 'time'.
        have_bta = (
            con.execute("""
            SELECT name FROM sqlite_master
            WHERE type='table' AND name='binding_time_annotations'
        """).fetchone()
            is not None
        )
        have_tic = (
            con.execute("""
            SELECT name FROM sqlite_master
            WHERE type='table' AND name='time_index_candidates'
        """).fetchone()
            is not None
        )

        if not have_bta:
            logger.info("No binding_time_annotations table found (IR time pass not run)")
            return axes

        # rank=1 time-axis range, for provenance (the axis the series point AT).
        time_index_range_by_binding: dict[str, str] = {}
        if have_tic:
            for row in con.execute("""
                SELECT t.binding_id, b.address_a1 AS range_a1
                FROM time_index_candidates t
                JOIN bindings b ON t.binding_id = b.binding_id
                WHERE t.rank = 1
            """).fetchall():
                time_index_range_by_binding[row["binding_id"]] = row["range_a1"]

        # series binding_id -> {time_index_binding_id, confidence}, gated by shape.
        link_by_binding: dict[str, dict[str, Any]] = {}
        for row in con.execute("""
            SELECT a.binding_id, a.time_index_binding_id, a.confidence
            FROM binding_time_annotations a
            JOIN bindings b ON a.binding_id = b.binding_id
            WHERE a.is_time_dependent = 1
              AND MAX(b.shape_rows, b.shape_cols) > 1
        """).fetchall():
            link_by_binding[row["binding_id"]] = {
                "time_index_binding_id": row["time_index_binding_id"],
                "time_index_confidence": float(row["confidence"])
                if row["confidence"] is not None
                else None,
            }

        logger.info(f"Found {len(link_by_binding)} gated time-dependent series bindings")

        # Observability: the set of time-dependent bindings the shape gate EXCLUDES
        # (1x1 scalars that merely read the time index). Queried only when the caller
        # asks for drop counters, so the default path pays nothing.
        shape_gated_bindings: set = set()
        if counters is not None:
            shape_gated_bindings = {
                row["binding_id"]
                for row in con.execute("""
                    SELECT a.binding_id
                    FROM binding_time_annotations a
                    JOIN bindings b ON a.binding_id = b.binding_id
                    WHERE a.is_time_dependent = 1
                      AND MAX(b.shape_rows, b.shape_cols) <= 1
                """).fetchall()
            }

        linked_count = 0
        for axis in axes:
            binding_id = axis.get("binding_id")
            if binding_id and binding_id in link_by_binding:
                info = link_by_binding[binding_id]
                tib = info["time_index_binding_id"]
                axis["time_index_binding_id"] = tib
                axis["time_index_range_a1"] = time_index_range_by_binding.get(tib)
                axis["time_index_confidence"] = info["time_index_confidence"]
                axis["axis_kind"] = TIME_AXIS_NAME
                # Upgrade generic axis names to the single shared constant.
                if axis.get("axis_name") in {"row_index", "column_headers"}:
                    axis["axis_name"] = TIME_AXIS_NAME
                linked_count += 1
                if counters is not None:
                    counters["time_linked"] = counters.get("time_linked", 0) + 1
            elif counters is not None and binding_id:
                # Binding present but did not link: distinguish the shape-gated scalar
                # (a real but deliberately-excluded time reader) from a binding that
                # carries no time-dependence annotation at all.
                if binding_id in shape_gated_bindings:
                    counters["drop_shape_le_1"] = counters.get("drop_shape_le_1", 0) + 1
                else:
                    counters["drop_not_time_dependent"] = (
                        counters.get("drop_not_time_dependent", 0) + 1
                    )

        logger.info(f"Linked {linked_count} axes to time index evidence")

        return axes

    finally:
        con.close()


# B3 (Gap 5): generic axis-name literals eligible for a dimension rename. The B1
# 'time' axis (axis_kind='time' / axis_name='time') is deliberately NOT in this set
# so it is never overwritten by a dimension tag.
_GENERIC_AXIS_NAMES = frozenset({"row_index", "column_headers"})

# B3 (Gap 5): the dimension-id -> clean-axis-name vocabulary. This is a versioned
# deterministic rename (a semantic assertion that the tag's id maps to this label),
# NOT a pure SQL join. The names are the snake_case suffix of each ontology
# concept.dimension.* id (concept.dimension.smoker_status -> 'smoker_status'), kept
# stable so two builds produce byte-identical axis_name values. The map is BUILT
# from the loaded ontology (the single source of truth for which 35 dimension ids
# exist) rather than hand-duplicated.
_DIMENSION_NAME_MAP_VERSION = "b3-v1"


def _build_dimension_name_map() -> dict[str, str]:
    """Return {concept.dimension.* id -> clean snake_case axis name}.

    Reads the controlled vocabulary from the actuarial ontology so the set of
    dimension ids stays in sync with ontology_v1.json (35 ids). The axis name is
    the id suffix after 'concept.dimension.' (already clean snake_case), which is
    the report's fixed id->name table.
    """
    from xl_marinade.docs.ontology_loader import load_ontology

    ontology = load_ontology(domain="actuarial")
    prefix = "concept.dimension."
    mapping: dict[str, str] = {}
    for concept in ontology.get("concepts") or []:
        cid = str(concept.get("id") or "")
        if cid.startswith(prefix):
            mapping[cid] = cid[len(prefix) :]
    return mapping


def rename_axes_from_dimension_tags(semantic_db_path: str) -> int:
    """B3 (Gap 5) pass: overwrite generic semantic_axes.axis_name from the
    highest-confidence concept.dimension.* tag already computed on that axis.

    Precedence (documented):
      1. an axis-target dimension tag (target_type='axis', target_id=axis_id) —
         per-axis, so unambiguous; applied directly.
      2. else a table-target dimension tag on the axis's table_id — but ONLY when
         that axis is the SOLE remaining generic axis of its table. A single
         table-level dimension tag is orientation-ambiguous: Mortality's table
         tag is 'age', yet its column axis is the sex x smoker key
         ['MNS','FNS','MS','FS'] — naming that column 'age' is a silent false
         assertion. When two+ generic axes of one table would draw on the same
         single table tag, they stay generic (Q3 + no-fabrication).
    Level 1 covers one model (48/100 axes carry an axis-target dim tag) and
    another model's Policyholder 'sex' axis; level 2 names single-axis tables only.

    Determinism: only axes whose current name is generic ('row_index' /
    'column_headers') AND whose axis_kind is not 'time' are renamed (B1's time
    axis is preserved). When several tags share the max confidence at the chosen
    level, the tie-break is (confidence DESC, concept_id ASC). The renamed axis
    inherits the source tag's confidence (0.4-0.8); no downstream 0.35 filter
    drops it (audited: query.py / candidates_v1.py / entity_model/builder.py read
    semantic_axes unfiltered). No fabrication: an axis is renamed only when a real
    dimension tag exists; otherwise it stays generic.

    Returns the number of axes renamed.
    """
    name_map = _build_dimension_name_map()

    con = sqlite3.connect(semantic_db_path)
    con.row_factory = sqlite3.Row
    try:
        axes = con.execute(
            """
            SELECT axis_id, table_id, axis_name, axis_kind
            FROM semantic_axes
            ORDER BY axis_id
            """
        ).fetchall()

        # Eligible = generic-named, non-time axes. Pinned order for determinism.
        eligible = [
            axis
            for axis in axes
            if axis["axis_kind"] != TIME_AXIS_NAME
            and axis["axis_name"] != TIME_AXIS_NAME
            and axis["axis_name"] in _GENERIC_AXIS_NAMES
        ]

        # Pass 1: axis-target tags are per-axis and unambiguous — apply directly.
        updates: list[tuple[str, str, float]] = []  # (axis_id, name, conf)
        named_axis_ids = set()
        table_fallback_candidates: dict[str, list[Any]] = {}
        for axis in eligible:
            chosen = _choose_axis_target_tag(con, axis["axis_id"], name_map)
            if chosen is not None:
                new_name, new_conf = chosen
                updates.append((axis["axis_id"], new_name, new_conf))
                named_axis_ids.add(axis["axis_id"])
            else:
                table_fallback_candidates.setdefault(axis["table_id"], []).append(axis)

        # Pass 2: table-target fallback. A single table-level dimension tag is
        # orientation-AMBIGUOUS: it cannot be deterministically attributed to one
        # of several axes of the same table (e.g. Mortality's table tag is 'age',
        # but its column axis is the sex x smoker key ['MNS','FNS','MS','FS'] —
        # naming that column 'age' is a silent false assertion the no-fabrication
        # rule + Q3 forbid). So a table fallback only names an axis when it is the
        # SOLE remaining generic axis of that table; otherwise those axes stay
        # generic.
        for table_id, axis_list in sorted(table_fallback_candidates.items()):
            if len(axis_list) != 1:
                continue
            axis = axis_list[0]
            chosen = _choose_table_target_tag(con, table_id, name_map)
            if chosen is not None:
                new_name, new_conf = chosen
                updates.append((axis["axis_id"], new_name, new_conf))

        renamed = 0
        for axis_id, new_name, new_conf in sorted(updates):
            con.execute(
                "UPDATE semantic_axes SET axis_name = ?, confidence = ? WHERE axis_id = ?",
                (new_name, new_conf, axis_id),
            )
            renamed += 1

        con.commit()
        logger.info(
            "B3 axis rename: %d/%d axes named from dimension tags (map=%s)",
            renamed,
            len(axes),
            _DIMENSION_NAME_MAP_VERSION,
        )
        return renamed
    finally:
        con.close()


def _choose_dimension_tag_at(
    con: sqlite3.Connection,
    target_type: str,
    target_id: str,
    name_map: dict[str, str],
) -> tuple[str, float] | None:
    """Pick (axis_name, confidence) from the dimension tags at one target.

    Highest confidence wins; ties break on concept_id ASC. Only ids present in
    name_map (real concept.dimension.* ids) are eligible — no fabricated names.
    """
    rows = con.execute(
        """
        SELECT concept_id, confidence
        FROM semantic_concept_tags
        WHERE target_type = ? AND target_id = ?
          AND concept_id LIKE 'concept.dimension.%'
        ORDER BY confidence DESC, concept_id ASC
        """,
        (target_type, target_id),
    ).fetchall()
    for row in rows:
        cid = row["concept_id"]
        if cid in name_map:
            conf = row["confidence"]
            return name_map[cid], float(conf) if conf is not None else 0.4
    return None


def _choose_axis_target_tag(
    con: sqlite3.Connection, axis_id: str, name_map: dict[str, str]
) -> tuple[str, float] | None:
    """Per-axis dimension tag (unambiguous): the axis names itself."""
    return _choose_dimension_tag_at(con, "axis", axis_id, name_map)


def _choose_table_target_tag(
    con: sqlite3.Connection, table_id: str, name_map: dict[str, str]
) -> tuple[str, float] | None:
    """Table-level dimension tag (used only when one generic axis remains)."""
    return _choose_dimension_tag_at(con, "table", table_id, name_map)


def produce_named_time_axis(ir_db_path: str) -> dict[str, Any] | None:
    """
    B1 (Gap 1) producer: emit a single named time-axis dict read directly from the
    IR rank=1 time_index_candidates row.

    The axis NAME stays the deterministic literal 'time' (Q3: no LLM naming in v1).
    samples_json holds the actual axis cells; time_index_binding_id carries the
    identity so the consumer can join. Returns None when no rank=1 time index exists.

    Selection: the named time axis is the index the workbook's series actually EVOLVE
    over — rank=1 candidates are ordered by graph-backed dependent-series support
    (binding_time_annotations at confidence >= TIME_GRAPH_CONFIDENCE_MIN, shape-gated),
    then total gated series, then candidate confidence, then address. The old
    (confidence DESC, address ASC)-only order let a confidence-tied MIS-FLAGGED index
    (an age column on an alphabetically-early sheet, zero series depending on it)
    out-sort the real period axis every projection series graph-depends on.
    Deterministic: every tie-break level is pinned.
    """
    from xl_marinade.core.time_dependence import TIME_GRAPH_CONFIDENCE_MIN

    con = sqlite3.connect(ir_db_path)
    con.row_factory = sqlite3.Row
    try:
        have_tic = (
            con.execute("""
            SELECT name FROM sqlite_master
            WHERE type='table' AND name='time_index_candidates'
        """).fetchone()
            is not None
        )
        if not have_tic:
            return None

        candidates = con.execute("""
            SELECT t.binding_id, t.confidence, b.address_a1, s.sheet_name
            FROM time_index_candidates t
            JOIN bindings b ON t.binding_id = b.binding_id
            JOIN sheets s ON b.sheet_id = s.sheet_id
            WHERE t.rank = 1
            ORDER BY t.confidence DESC, b.address_a1 ASC
        """).fetchall()
        if not candidates:
            return None

        # Dependent-series support per time index (graph-backed first). Same shape
        # gate as link_time_axis so 1x1 scalars that merely read the index don't vote.
        support: dict[str, tuple[int, int]] = {}
        have_bta = (
            con.execute("""
            SELECT name FROM sqlite_master
            WHERE type='table' AND name='binding_time_annotations'
        """).fetchone()
            is not None
        )
        if have_bta:
            for row in con.execute(
                """
                SELECT a.time_index_binding_id AS tib,
                       COUNT(*) AS n_series,
                       SUM(CASE WHEN a.confidence >= ? THEN 1 ELSE 0 END) AS n_graph
                FROM binding_time_annotations a
                JOIN bindings b ON a.binding_id = b.binding_id
                WHERE a.is_time_dependent = 1
                  AND MAX(b.shape_rows, b.shape_cols) > 1
                GROUP BY a.time_index_binding_id
            """,
                (TIME_GRAPH_CONFIDENCE_MIN,),
            ):
                support[row["tib"]] = (int(row["n_graph"] or 0), int(row["n_series"] or 0))

        def _rank(c) -> tuple:
            n_graph, n_series = support.get(c["binding_id"], (0, 0))
            conf = float(c["confidence"]) if c["confidence"] is not None else 0.0
            return (-n_graph, -n_series, -conf, c["address_a1"])

        chosen = min(candidates, key=_rank)
        binding_id = chosen["binding_id"]
        sheet = chosen["sheet_name"]
        address = chosen["address_a1"]
        header_range_a1 = address if "!" in address else f"{sheet}!{address}"
        samples = _load_axis_samples_ordered(ir_db_path, binding_id, max_samples=8)
        # A vertical time axis (A7:A607) is a row orientation; otherwise column.
        bounds = _parse_range_bounds(sheet, address)
        orientation = "row"
        if bounds is not None:
            _, r1, c1, r2, c2 = bounds
            orientation = "row" if (r2 - r1) >= (c2 - c1) else "column"

        return {
            "orientation": orientation,
            "header_range_a1": header_range_a1,
            "axis_name": TIME_AXIS_NAME,
            "samples_json": json.dumps(samples),
            "confidence": float(chosen["confidence"]) if chosen["confidence"] is not None else None,
            "binding_id": binding_id,
            "time_index_binding_id": binding_id,
            "axis_kind": TIME_AXIS_NAME,
        }
    finally:
        con.close()


def _get_binding_range(ir_db_path: str, binding_id: str) -> str:
    """Get range_a1 for a binding."""
    con = sqlite3.connect(ir_db_path)
    try:
        row = con.execute(
            "SELECT address_a1 FROM bindings WHERE binding_id = ?", (binding_id,)
        ).fetchone()
        return row[0] if row else "UNKNOWN"
    finally:
        con.close()


def _load_axis_samples(
    ir_db_path: str,
    binding_id: str,
    max_samples: int = 5,
    orientation: str = "row",
) -> list[Any]:
    """
    Load sample values from an axis binding.

    Args:
        ir_db_path: Path to IR database
        binding_id: Binding ID to load samples from
        max_samples: Maximum number of samples to return
        orientation: "row" or "column". B3 (Gap 5) sample-contamination fix —
            a row-oriented axis binding starts at the HEADER row (e.g.
            'Mortality table'!B1:B101 where B1='MNS'), so the first spatial cell
            is the column header, not a numeric axis value. For row axes we drop
            the top (header) cell so value sampling starts one row below. Column
            axes legitimately ARE the header row, so they are not skipped.

    Returns:
        List of sample values (up to max_samples)
    """
    con = sqlite3.connect(ir_db_path)
    con.row_factory = sqlite3.Row

    try:
        # Try agent_cells view first (new schema). Fetch ALL cells (no SQL LIMIT)
        # so we can sort spatially and skip the header deterministically — a
        # lexicographic cell_address LIMIT would both mis-order (A101<A2) and
        # truncate before the skip.
        try:
            rows = con.execute(
                """
                SELECT ac.cell_address, ac.value
                FROM cell_to_binding ctb
                JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
                WHERE ctb.binding_id = ?
            """,
                (binding_id,),
            ).fetchall()
        except sqlite3.OperationalError:
            # Fall back to old schema
            try:
                rows = con.execute(
                    """
                    SELECT cell_address_a1, evaluated_value
                    FROM cells
                    WHERE binding_id = ?
                """,
                    (binding_id,),
                ).fetchall()
            except sqlite3.OperationalError:
                # No cells table (test fixture or incomplete IR)
                logger.debug(f"No cells table found for binding {binding_id}")
                return []

        def _key(r):
            addr = r[0] or ""
            pure = addr.split("!")[-1]
            parsed = parse_cell_address(addr if "!" in addr else f"X!{pure}")
            return (int(parsed.get("row", 0) or 0), int(parsed.get("col", 0) or 0))

        ordered = sorted(rows, key=_key)
        # B3 contamination fix: drop the top (header) cell for row-oriented axes.
        if orientation == "row" and ordered:
            ordered = ordered[1:]

        samples = []
        for row in ordered:
            value = row[1]  # Second column is value in both schemas
            if value is not None:
                samples.append(value)
            if len(samples) >= max_samples:
                break

        return samples

    finally:
        con.close()


def _load_axis_samples_ordered(ir_db_path: str, binding_id: str, max_samples: int = 8) -> list[Any]:
    """Load a binding's cells as samples in true spatial order, JSON-unwrapped.

    Unlike _load_axis_samples (which orders cell_address lexicographically — so
    A100 sorts before A99 and the LIMIT then grabs the wrong cells) this fetches
    ALL the binding's cells, sorts by parsed (row, col), truncates, and unwraps the
    stored JSON value. The result reads t=0,1,2,... and matches the clean string
    format every other axis emits (the agent_cells.value is JSON-encoded as
    '"3"'; the legacy cells.evaluated_value is the raw '3'). Used only by the B1
    named-time-axis producer so the existing extract_axes sample format is untouched.
    """
    con = sqlite3.connect(ir_db_path)
    con.row_factory = sqlite3.Row
    try:
        try:
            rows = con.execute(
                """
                SELECT ac.cell_address, ac.value
                FROM cell_to_binding ctb
                JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
                WHERE ctb.binding_id = ?
            """,
                (binding_id,),
            ).fetchall()
        except sqlite3.OperationalError:
            try:
                rows = con.execute(
                    """
                    SELECT cell_address_a1, evaluated_value
                    FROM cells
                    WHERE binding_id = ?
                """,
                    (binding_id,),
                ).fetchall()
            except sqlite3.OperationalError:
                return []

        def _key(r):
            addr = r[0] or ""
            pure = addr.split("!")[-1]
            parsed = parse_cell_address(addr if "!" in addr else f"X!{pure}")
            return (int(parsed.get("row", 0) or 0), int(parsed.get("col", 0) or 0))

        ordered = sorted((r for r in rows if r[1] is not None), key=_key)[:max_samples]
        samples: list[Any] = []
        for r in ordered:
            v = r[1]
            try:
                samples.append(json.loads(v))
            except (ValueError, TypeError):
                samples.append(v)
        return samples
    finally:
        con.close()


def _parse_range_bounds(sheet: str, range_a1: str) -> tuple[str, int, int, int, int] | None:
    if not range_a1:
        return None
    parsed = parse_cell_address(range_a1 if "!" in range_a1 else f"{sheet}!{range_a1}")
    row = int(parsed.get("row", 0) or 0)
    col = int(parsed.get("col", 0) or 0)
    height = int(parsed.get("height", 1) or 1)
    width = int(parsed.get("width", 1) or 1)
    if row <= 0 or col <= 0:
        return None
    return parsed.get("sheet", sheet) or sheet, row, col, row + height - 1, col + width - 1


def _load_axis_samples_by_range(
    ir_db_path: str,
    sheet: str,
    row_start: int,
    col_start: int,
    row_end: int,
    col_end: int,
    orientation: str,
    max_samples: int,
) -> list[Any]:
    con = sqlite3.connect(ir_db_path)
    con.row_factory = sqlite3.Row
    try:
        # Prefer agent_cells view (new schema)
        try:
            if orientation == "row":
                rows = con.execute(
                    """
                    SELECT cell_address, value
                    FROM agent_cells
                    WHERE sheet = ? AND col = ? AND row BETWEEN ? AND ?
                    ORDER BY row
                    LIMIT ?
                    """,
                    (sheet, col_start, row_start, row_end, max_samples),
                ).fetchall()
            else:
                rows = con.execute(
                    """
                    SELECT cell_address, value
                    FROM agent_cells
                    WHERE sheet = ? AND row = ? AND col BETWEEN ? AND ?
                    ORDER BY col
                    LIMIT ?
                    """,
                    (sheet, row_start, col_start, col_end, max_samples),
                ).fetchall()
        except sqlite3.OperationalError:
            return []

        samples: list[Any] = []
        for r in rows:
            v = r[1]
            if v is not None:
                try:
                    samples.append(json.loads(v))
                except Exception:
                    samples.append(v)
        return samples
    finally:
        con.close()


def _infer_axes_from_range(ir_db_path: str, table: dict[str, Any]) -> list[dict[str, Any]]:
    """
    Infer axes from table range.

    Story 8 enhancement: Respect table_kind from provenance to ensure vector tables
    get exactly one axis (not two).
    """
    sheet = table.get("sheet") or ""
    range_a1 = table.get("range_a1") or ""
    bounds = _parse_range_bounds(sheet, range_a1)
    if not bounds:
        return []
    sheet, r1, c1, r2, c2 = bounds

    # Check if this is a vector table (Story 8)
    provenance = {}
    try:
        provenance = json.loads(table.get("provenance_json") or "{}")
    except Exception:
        provenance = {}

    table_kind = provenance.get("table_kind", "grid")

    axes: list[dict[str, Any]] = []

    # For vector tables, only create one axis (the longer dimension)
    if table_kind == "vector":
        rows = r2 - r1 + 1
        cols = c2 - c1 + 1

        if rows > cols:
            # Row vector (vertical)
            col_letter = col_num_to_letter(c1)
            header_range = f"{sheet}!{col_letter}{r1}:{col_letter}{r2}"
            samples = _load_axis_samples_by_range(ir_db_path, sheet, r1, c1, r2, c1, "row", 6)
            axes.append(
                {
                    "orientation": "row",
                    "header_range_a1": header_range,
                    "axis_name": "row_index",
                    "samples_json": json.dumps(samples),
                    "confidence": 0.5,  # Higher confidence for vector (explicit from IR)
                    "binding_id": None,
                }
            )
        else:
            # Column vector (horizontal)
            row_number = r1
            start_col = col_num_to_letter(c1)
            end_col = col_num_to_letter(c2)
            header_range = f"{sheet}!{start_col}{row_number}:{end_col}{row_number}"
            samples = _load_axis_samples_by_range(ir_db_path, sheet, r1, c1, r1, c2, "column", 8)
            axes.append(
                {
                    "orientation": "column",
                    "header_range_a1": header_range,
                    "axis_name": "column_headers",
                    "samples_json": json.dumps(samples),
                    "confidence": 0.5,
                    "binding_id": None,
                }
            )
        return axes

    # For grid tables, create both axes (existing behavior)
    if r2 > r1:
        col_letter = col_num_to_letter(c1)
        header_range = f"{sheet}!{col_letter}{r1}:{col_letter}{r2}"
        # B3 (Gap 5) contamination fix: the grid's row-axis index runs down the
        # left column whose row 1 is a string header (e.g. A1='age-last' / 'MNS').
        # Start value sampling at r1+1 so the header value stops contaminating the
        # numeric row samples. The column axis below correctly starts at r1 (its
        # samples ARE the row-1 headers), so only the row axis shifts.
        row_sample_start = min(r1 + 1, r2)
        samples = _load_axis_samples_by_range(
            ir_db_path, sheet, row_sample_start, c1, r2, c1, "row", 6
        )
        axes.append(
            {
                "orientation": "row",
                "header_range_a1": header_range,
                "axis_name": "row_index",
                "samples_json": json.dumps(samples),
                "confidence": 0.35,
                "binding_id": None,
            }
        )
    if c2 > c1:
        row_number = r1
        start_col = col_num_to_letter(c1)
        end_col = col_num_to_letter(c2)
        header_range = f"{sheet}!{start_col}{row_number}:{end_col}{row_number}"
        samples = _load_axis_samples_by_range(ir_db_path, sheet, r1, c1, r1, c2, "column", 8)
        axes.append(
            {
                "orientation": "column",
                "header_range_a1": header_range,
                "axis_name": "column_headers",
                "samples_json": json.dumps(samples),
                "confidence": 0.35,
                "binding_id": None,
            }
        )
    return axes


def infer_axes_from_table_range(ir_db_path: str, table: dict[str, Any]) -> list[dict[str, Any]]:
    """Public wrapper to infer axes directly from a table range."""
    return _infer_axes_from_range(ir_db_path, table)
