-- ABOUTME: Schema v2.0 for memory-efficient IR extraction with staging heap tables and normalized final tables.
-- ABOUTME: Staging tables (raw_*) are standard heap tables with NO PK/UNIQUE/indexes for fast bulk loading.

-- ============================================================================
-- STAGING TABLES (build-time only; heap tables with no PK/UNIQUE/indexes)
-- ============================================================================

CREATE TABLE raw_edges_internal (
    from_cell_id INTEGER NOT NULL,
    to_cell_id INTEGER NOT NULL
);

CREATE TABLE raw_edges_range (
    from_cell_id INTEGER NOT NULL,
    to_sheet_id INTEGER NOT NULL,
    to_r1 INTEGER NOT NULL,
    to_c1 INTEGER NOT NULL,
    to_r2 INTEGER NOT NULL,
    to_c2 INTEGER NOT NULL,
    to_range_a1 TEXT NOT NULL,
    cell_count INTEGER NOT NULL,
    provenance TEXT
);

CREATE TABLE raw_edges_external (
    from_cell_id INTEGER NOT NULL,
    external_ref TEXT NOT NULL
);

CREATE TABLE raw_formulas (
    formula_r1c1 TEXT NOT NULL,
    formula_a1_example TEXT,
    source_sheet_id INTEGER,
    source_row INTEGER,
    source_col INTEGER
);

CREATE TABLE raw_json_blobs (
    sha256 TEXT NOT NULL,
    json TEXT NOT NULL
);

CREATE TABLE raw_cells (
    cell_id INTEGER NOT NULL,
    sheet_id INTEGER NOT NULL,
    row INTEGER NOT NULL,
    col INTEGER NOT NULL,
    a1 TEXT NOT NULL,
    formula_r1c1 TEXT,
    formula_a1 TEXT,
    value_sha256 TEXT,
    format_sha256 TEXT,
    data_type TEXT,
    is_array_formula BOOLEAN DEFAULT 0,
    is_spilled BOOLEAN DEFAULT 0,
    spilled_from_cell_id INTEGER
);

-- ============================================================================
-- FINAL TABLES (normalized storage layer)
-- ============================================================================

-- Sheet catalog
CREATE TABLE sheets (
    sheet_id INTEGER PRIMARY KEY,
    sheet_name TEXT NOT NULL UNIQUE
);

-- Deduplicated formulas
-- formula_canonical_a1 (Cycle 17 #312-B): position-invariant canonical
-- form for cross-position family grouping. NULL when the formula fails
-- the meaningfulness gate (no "(" in canonical). Sits alongside
-- formula_r1c1; existing R1C1 consumers are unchanged.
CREATE TABLE formulas (
    formula_id INTEGER PRIMARY KEY,
    formula_r1c1 TEXT NOT NULL UNIQUE,
    formula_a1_example TEXT,
    formula_canonical_a1 TEXT
);
CREATE INDEX idx_formulas_canonical_a1 ON formulas(formula_canonical_a1)
    WHERE formula_canonical_a1 IS NOT NULL;

-- Deduplicated JSON blobs
CREATE TABLE json_blobs (
    blob_id INTEGER PRIMARY KEY,
    sha256 TEXT NOT NULL UNIQUE,
    json TEXT NOT NULL
);

-- Cells
CREATE TABLE cells (
    cell_id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL REFERENCES sheets(sheet_id),
    row INTEGER NOT NULL,
    col INTEGER NOT NULL,
    a1 TEXT NOT NULL,
    formula_id INTEGER REFERENCES formulas(formula_id),
    formula_a1 TEXT,
    format_blob_id INTEGER REFERENCES json_blobs(blob_id),
    value_blob_id INTEGER REFERENCES json_blobs(blob_id),
    data_type TEXT,
    is_array_formula BOOLEAN DEFAULT 0,
    is_spilled BOOLEAN DEFAULT 0,
    spilled_from_cell_id INTEGER,
    UNIQUE(sheet_id, row, col)
);

-- Internal Cell Edges (WITHOUT ROWID for compact storage)
CREATE TABLE cell_edges_internal (
    from_cell_id INTEGER NOT NULL,
    to_cell_id INTEGER NOT NULL,
    PRIMARY KEY (from_cell_id, to_cell_id)
) WITHOUT ROWID;
-- Speedup: to_cell_id lookups for binding edge collapse (PK has from_cell_id first)
CREATE INDEX idx_cell_edges_internal_to ON cell_edges_internal(to_cell_id);

-- External Cell Edges
CREATE TABLE cell_edges_external (
    from_cell_id INTEGER NOT NULL,
    external_ref TEXT NOT NULL,  -- "[Book.xlsx]Sheet!A1" or "UNRESOLVED:<token>" or "DYNAMIC:<token>"
    PRIMARY KEY (from_cell_id, external_ref)
) WITHOUT ROWID;

-- Range Edges
CREATE TABLE range_edges (
    from_cell_id INTEGER NOT NULL,
    to_sheet_id INTEGER NOT NULL,
    to_r1 INTEGER NOT NULL,
    to_c1 INTEGER NOT NULL,
    to_r2 INTEGER NOT NULL,
    to_c2 INTEGER NOT NULL,
    to_range_a1 TEXT NOT NULL, -- Canonical A1 (no sheet), stored explicitly to handle >Z columns
    cell_count INTEGER NOT NULL,
    provenance TEXT, -- NULL/'static' = ordinary range ref; 'resolved_from_cache' = Issue #1 by-value INDIRECT/OFFSET (snapshot-specific)
    PRIMARY KEY (from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2)
) WITHOUT ROWID;

-- Bindings (Unchanged logic, updated schema refs)
CREATE TABLE bindings (
    binding_id TEXT PRIMARY KEY,
    sheet_id INTEGER NOT NULL REFERENCES sheets(sheet_id),
    address_a1 TEXT NOT NULL,
    top_left_cell_id INTEGER NOT NULL,
    shape_rows INTEGER NOT NULL,
    shape_cols INTEGER NOT NULL,
    binding_type TEXT NOT NULL,
    formula_id INTEGER REFERENCES formulas(formula_id),
    label TEXT,
    classification TEXT,
    confidence REAL,
    is_orphan BOOLEAN DEFAULT 0,
    extraction_source TEXT,
    evidence_blob_id INTEGER REFERENCES json_blobs(blob_id),
    spatial_candidates_blob_id INTEGER REFERENCES json_blobs(blob_id),
    -- Ordered header hierarchy [{role,text,cell,distance,span}, ...] derived from the
    -- label candidates (primary == label, then group/axis/title/sheet). Populated by
    -- backfill_binding_labels. See documentation_agent/header_context.py.
    header_context_json TEXT
);

-- Cell to binding
CREATE TABLE cell_to_binding (
    cell_id INTEGER NOT NULL,
    binding_id TEXT NOT NULL,
    PRIMARY KEY (cell_id, binding_id)
) WITHOUT ROWID;

-- Materialized label-candidate cells (address-level evidence without JSON parsing)
CREATE TABLE binding_label_candidate_cells (
    binding_id TEXT NOT NULL REFERENCES bindings(binding_id),
    candidate_type TEXT NOT NULL,
    candidate_address TEXT NOT NULL,
    cell_address TEXT NOT NULL,
    sheet_id INTEGER NOT NULL REFERENCES sheets(sheet_id),
    row INTEGER NOT NULL,
    col INTEGER NOT NULL,
    value_text TEXT,
    PRIMARY KEY (binding_id, candidate_type, candidate_address, cell_address)
) WITHOUT ROWID;

CREATE INDEX idx_binding_label_candidate_cells_cell_address
    ON binding_label_candidate_cells (cell_address);
CREATE INDEX idx_binding_label_candidate_cells_sheet_row_col
    ON binding_label_candidate_cells (sheet_id, row, col);

-- Speedup: binding_id lookups on cell_to_binding (PK has cell_id first)
CREATE INDEX idx_cell_to_binding_binding_id ON cell_to_binding(binding_id);
-- Speedup: sheet-scoped binding queries
CREATE INDEX idx_bindings_sheet_id ON bindings(sheet_id);
-- Speedup: evidence cache bounding-box queries (prereq for WI-7)
CREATE INDEX idx_cells_sheet_row_col ON cells(sheet_id, row, col);

-- Binding edges
-- kind discriminates the dependency mechanism. 'formula' is the default for
-- precedent edges synthesised from formula DAG walking; 'via_vba_paste' is
-- the edge kind emitted by the VBA paste-edge synthesiser (R21) for
-- target-cell-block ← source-template-row dependencies that flow through a
-- VBA PasteSpecial / .Value=.Value statement. provenance_proc (when set)
-- names the VBA procedure (module::name) responsible for synthesising the
-- edge, so writer trace prose can surface "propagated by VBA paste from
-- {procedure}".
CREATE TABLE binding_edges (
    from_binding_id TEXT NOT NULL,
    to_binding_id TEXT NOT NULL,
    edge_count INTEGER NOT NULL DEFAULT 1,
    kind TEXT NOT NULL DEFAULT 'formula',
    provenance_proc TEXT,
    PRIMARY KEY (from_binding_id, to_binding_id)
) WITHOUT ROWID;

-- User roots (Story 15 compatibility for dependency traversal)
CREATE TABLE user_roots (
    root_id INTEGER PRIMARY KEY,
    sheet TEXT NOT NULL,
    range_a1 TEXT NOT NULL,
    label_hint TEXT
);

-- Resolution metrics (ADR-041 verification for fast schema)
CREATE TABLE resolution_metrics (
    function_name TEXT NOT NULL,
    status TEXT NOT NULL,
    count INTEGER NOT NULL,
    PRIMARY KEY (function_name, status)
);

-- Sprint 6 deterministic outputs (time axis + time dependence)
CREATE TABLE time_index_candidates (
    sheet_id INTEGER NOT NULL REFERENCES sheets(sheet_id),
    binding_id TEXT NOT NULL REFERENCES bindings(binding_id),
    rank INTEGER NOT NULL,
    confidence REAL NOT NULL,
    reasons_top3_json TEXT NOT NULL,
    PRIMARY KEY (sheet_id, rank)
);

CREATE TABLE binding_time_annotations (
    binding_id TEXT PRIMARY KEY REFERENCES bindings(binding_id),
    time_index_binding_id TEXT NOT NULL REFERENCES bindings(binding_id),
    is_time_dependent BOOLEAN NOT NULL,
    confidence REAL NOT NULL,
    reasons_top3_json TEXT NOT NULL,
    evidence_flags_json TEXT NOT NULL
);

-- ============================================================================
-- AGENT-FRIENDLY VIEWS (backward compatibility layer)
-- ============================================================================
-- These views provide a stable query interface for the documentation agent,
-- abstracting the normalized storage schema. Agents MUST access data via views,
-- not raw tables, to ensure compatibility across schema versions.

-- Cells View (full, backward compatible)
-- Exposes all cell data including JSON blobs (value/format/extras).
-- Use agent_cells_light for safer default without large JSON payloads.
CREATE VIEW agent_cells AS
SELECT
    s.sheet_name || '!' || c.a1 AS cell_address,
    s.sheet_name AS sheet,
    c.row,
    c.col,
    -- Cell-local A1 formula, NOT formulas.formula_a1_example: the example is
    -- rendered at one representative cell of the shared R1C1 formula, so for
    -- every other member cell it is textually wrong for that address (e.g.
    -- E14 whose formula is =E13 shows as =A2). Consumers pair this column
    -- with the cell address, so it must be the cell's own rendering.
    c.formula_a1 AS formula,
    f.formula_r1c1,
    jv.json AS value,
    jf.json AS format,
    c.data_type,
    c.is_array_formula,
    c.is_spilled,
    c.cell_id
FROM cells c
JOIN sheets s ON c.sheet_id = s.sheet_id
LEFT JOIN formulas f ON c.formula_id = f.formula_id
LEFT JOIN json_blobs jv ON c.value_blob_id = jv.blob_id
LEFT JOIN json_blobs jf ON c.format_blob_id = jf.blob_id;

-- Cells View (light, safe default for agents)
-- Excludes JSON blob columns (value/format/extras) for better performance
-- and reduced memory footprint. Recommended for most agent queries.
CREATE VIEW agent_cells_light AS
SELECT
    s.sheet_name || '!' || c.a1 AS cell_address,
    s.sheet_name AS sheet,
    c.row,
    c.col,
    -- Cell-local A1 formula (see agent_cells comment above).
    c.formula_a1 AS formula,
    f.formula_r1c1,
    c.data_type,
    c.is_array_formula,
    c.is_spilled,
    c.cell_id
FROM cells c
JOIN sheets s ON c.sheet_id = s.sheet_id
LEFT JOIN formulas f ON c.formula_id = f.formula_id;

-- Cells View (explicit heavy payload)
-- Alias for agent_cells. Use when you explicitly need JSON blobs.
CREATE VIEW agent_cells_full AS
SELECT * FROM agent_cells;

-- Bindings View
-- Exposes bindings with sheet names and denormalized JSON evidence.
-- Includes formula patterns, labels, and classification metadata.
CREATE VIEW agent_bindings AS
SELECT 
    b.binding_id,
    s.sheet_name AS sheet,
    s.sheet_name || '!' || b.address_a1 AS address,
    b.shape_rows,
    b.shape_cols,
    b.binding_type,
    f.formula_r1c1 AS formula_pattern,
    b.label,
    b.classification,
    b.confidence,
    b.is_orphan,
    b.extraction_source,
    jsc.json AS spatial_candidates,
    jev.json AS evidence
FROM bindings b
JOIN sheets s ON b.sheet_id = s.sheet_id
LEFT JOIN formulas f ON b.formula_id = f.formula_id
LEFT JOIN json_blobs jsc ON b.spatial_candidates_blob_id = jsc.blob_id
LEFT JOIN json_blobs jev ON b.evidence_blob_id = jev.blob_id;

-- Phase 5: Unified node view — cell bindings + VBA procedures in one queryable surface.
-- Consumers needing cell-specific fields (shape, address) keep reading agent_bindings directly.
-- Consumers needing only node identity (node_id, label, kind) read atlas_nodes.
--
-- VBA node_id format: 'vba::<module>::<name>::<kind>' for unconditional
-- procedures, with '::<compile_branch>' appended when the procedure sits
-- inside a #If/#Else block. The kind suffix prevents Property Get/Let/Set
-- accessors with the same name from collapsing onto a single row, and the
-- optional compile_branch suffix preserves both halves of #If Win64/#Else
-- twin pairs. Consumers should split on '::' and not assume a fixed length.
CREATE VIEW atlas_nodes AS
SELECT
    b.binding_id AS node_id,
    'cell' AS node_kind,
    COALESCE(b.label, b.sheet || '!' || b.address) AS display_name,
    b.sheet,
    b.address,
    b.binding_type,
    b.classification,
    b.confidence
FROM agent_bindings b
UNION ALL
SELECT
    'vba::' || m.name || '::' || p.name || '::' || p.kind ||
        CASE WHEN p.compile_branch != '' THEN '::' || p.compile_branch ELSE '' END
        AS node_id,
    'procedure' AS node_kind,
    m.name || '::' || p.name AS display_name,
    NULL AS sheet,
    NULL AS address,
    p.kind AS binding_type,
    CASE WHEN p.is_event_handler = 1 THEN 'event_handler'
         WHEN p.kind = 'function' AND p.is_public = 1 THEN 'udf'
         ELSE 'orchestrator' END AS classification,
    1.0 AS confidence
FROM vba_procedures p
JOIN vba_modules m ON p.module_id = m.module_id;

CREATE VIEW agent_binding_label_candidate_cells AS
SELECT
    blc.binding_id,
    s.sheet_name AS sheet,
    blc.cell_address,
    blc.candidate_type,
    blc.candidate_address,
    blc.row,
    blc.col,
    blc.value_text
FROM binding_label_candidate_cells blc
JOIN sheets s ON blc.sheet_id = s.sheet_id;

-- Consolidated Dependencies View
-- UNION ALL of cell_edges_internal, range_edges, and cell_edges_external.
-- Provides uniform interface for all dependency types with metadata:
-- - dependency_type: 'cell' | 'range' | 'external'
-- - Range edges include: cell_count, to_r1, to_c1, to_r2, to_c2
-- - Whole-column/row refs appear as expanded bounds (e.g., A1:A1048576)
CREATE VIEW agent_dependencies AS
-- Internal Cell Dependencies
SELECT 
    s_from.sheet_name || '!' || c_from.a1 AS from_cell,
    s_to.sheet_name || '!' || c_to.a1 AS to_cell,
    'cell' AS dependency_type,
    NULL AS cell_count,
    NULL AS to_sheet_id,
    NULL AS to_r1,
    NULL AS to_c1,
    NULL AS to_r2,
    NULL AS to_c2
FROM cell_edges_internal e
JOIN cells c_from ON e.from_cell_id = c_from.cell_id
JOIN sheets s_from ON c_from.sheet_id = s_from.sheet_id
JOIN cells c_to ON e.to_cell_id = c_to.cell_id
JOIN sheets s_to ON c_to.sheet_id = s_to.sheet_id
UNION ALL
-- Range Dependencies
SELECT 
    s_from.sheet_name || '!' || c_from.a1 AS from_cell,
    s_to.sheet_name || '!' || re.to_range_a1 AS to_cell,
    'range' AS dependency_type,
    re.cell_count AS cell_count,
    re.to_sheet_id AS to_sheet_id,
    re.to_r1 AS to_r1,
    re.to_c1 AS to_c1,
    re.to_r2 AS to_r2,
    re.to_c2 AS to_c2
FROM range_edges re
JOIN cells c_from ON re.from_cell_id = c_from.cell_id
JOIN sheets s_from ON c_from.sheet_id = s_from.sheet_id
JOIN sheets s_to ON re.to_sheet_id = s_to.sheet_id
UNION ALL
-- External Dependencies
SELECT 
    s_from.sheet_name || '!' || c_from.a1 AS from_cell,
    ext.external_ref AS to_cell,
    'external' AS dependency_type,
    NULL AS cell_count,
    NULL AS to_sheet_id,
    NULL AS to_r1,
    NULL AS to_c1,
    NULL AS to_r2,
    NULL AS to_c2
FROM cell_edges_external ext
JOIN cells c_from ON ext.from_cell_id = c_from.cell_id
JOIN sheets s_from ON c_from.sheet_id = s_from.sheet_id;

-- Binding Dependencies View
-- Exposes binding-to-binding edges with labels and addresses for both endpoints.
-- kind / provenance_proc surfaced so consumers (writer trace prose) can
-- distinguish formula precedent edges from VBA-paste-synthesised edges.
CREATE VIEW agent_binding_dependencies AS
SELECT
    COALESCE(b_from.binding_id, be.from_binding_id) AS from_binding,
    COALESCE(s_from.sheet_name || '!' || b_from.address_a1, be.from_binding_id) AS from_address,
    COALESCE(b_from.label, be.from_binding_id) AS from_label,
    COALESCE(b_to.binding_id, be.to_binding_id) AS to_binding,
    COALESCE(s_to.sheet_name || '!' || b_to.address_a1, be.to_binding_id) AS to_address,
    COALESCE(b_to.label, be.to_binding_id) AS to_label,
    be.edge_count,
    be.kind,
    be.provenance_proc
FROM binding_edges be
LEFT JOIN bindings b_from ON be.from_binding_id = b_from.binding_id
LEFT JOIN sheets s_from ON b_from.sheet_id = s_from.sheet_id
LEFT JOIN bindings b_to ON be.to_binding_id = b_to.binding_id
LEFT JOIN sheets s_to ON b_to.sheet_id = s_to.sheet_id;

-- Time Index Candidates View
CREATE VIEW agent_time_index_candidates AS
SELECT
    s.sheet_name AS sheet,
    t.binding_id,
    b.address_a1 AS address_a1,
    t.rank,
    t.confidence,
    t.reasons_top3_json
FROM time_index_candidates t
JOIN sheets s ON t.sheet_id = s.sheet_id
JOIN bindings b ON t.binding_id = b.binding_id;

-- Binding Time Annotations View
CREATE VIEW agent_binding_time_annotations AS
SELECT
    b.binding_id,
    s.sheet_name AS sheet,
    b.address_a1 AS address_a1,
    t.time_index_binding_id,
    t.is_time_dependent,
    t.confidence,
    t.reasons_top3_json,
    t.evidence_flags_json
FROM binding_time_annotations t
JOIN bindings b ON t.binding_id = b.binding_id
JOIN sheets s ON b.sheet_id = s.sheet_id;

-- Defined names (Excel named ranges and named formulas)
CREATE TABLE defined_names (
    name TEXT NOT NULL,
    scope TEXT NOT NULL,        -- 'workbook' or sheet name for sheet-scoped names
    destinations TEXT NOT NULL,  -- JSON list of A1 refs, e.g. '["Sheet1!A1:A10"]'
    is_external INTEGER NOT NULL DEFAULT 0
);

-- Sprint 10 Story 7: Table candidates from 1D bindings
CREATE TABLE table_candidates (
    candidate_id TEXT PRIMARY KEY,
    sheet_id INTEGER NOT NULL REFERENCES sheets(sheet_id),
    kind TEXT NOT NULL CHECK (kind IN ('vector','grid')),
    r1 INTEGER NOT NULL,
    c1 INTEGER NOT NULL,
    r2 INTEGER NOT NULL,
    c2 INTEGER NOT NULL,
    range_a1 TEXT NOT NULL,
    confidence REAL NOT NULL,
    reasons_top3_json TEXT NOT NULL CHECK (json_valid(reasons_top3_json))
);

CREATE TABLE table_candidate_members (
    candidate_id TEXT NOT NULL REFERENCES table_candidates(candidate_id),
    ordinal INTEGER NOT NULL,
    binding_id TEXT NOT NULL REFERENCES bindings(binding_id),
    role_hint TEXT,
    PRIMARY KEY (candidate_id, ordinal),
    UNIQUE (candidate_id, binding_id)
) WITHOUT ROWID;

-- Table Candidates View (agent-friendly)
CREATE VIEW agent_table_candidates AS
SELECT
    tc.candidate_id,
    s.sheet_name AS sheet,
    tc.range_a1,
    tc.kind,
    tc.confidence,
    tc.r1,
    tc.c1,
    tc.r2,
    tc.c2,
    tc.reasons_top3_json
FROM table_candidates tc
JOIN sheets s ON tc.sheet_id = s.sheet_id;

-- Formula Families: groups of bindings sharing the same R1C1 on the same sheet
CREATE TABLE formula_families (
    family_id TEXT PRIMARY KEY,
    sheet_id INTEGER NOT NULL REFERENCES sheets(sheet_id),
    formula_id INTEGER NOT NULL REFERENCES formulas(formula_id),
    member_count INTEGER NOT NULL,
    representative_binding_id TEXT NOT NULL REFERENCES bindings(binding_id),
    UNIQUE (sheet_id, formula_id)
);

CREATE TABLE formula_family_members (
    family_id TEXT NOT NULL REFERENCES formula_families(family_id),
    ordinal INTEGER NOT NULL,
    binding_id TEXT NOT NULL REFERENCES bindings(binding_id),
    PRIMARY KEY (family_id, ordinal),
    UNIQUE (family_id, binding_id)
) WITHOUT ROWID;

-- Formula Families View (agent-friendly)
CREATE VIEW agent_formula_families AS
SELECT
    ff.family_id,
    s.sheet_name AS sheet,
    f.formula_r1c1,
    ff.member_count,
    ff.representative_binding_id
FROM formula_families ff
JOIN sheets s ON ff.sheet_id = s.sheet_id
JOIN formulas f ON ff.formula_id = f.formula_id;

-- VBA User-Defined Functions
-- Phase 0: UDF-only persistence. Populated via ir_extractor.vba_parser.extract_udfs_from_path()
-- after the cells finalize stage. Phase 2+ supersedes with vba_procedures covering Subs,
-- event handlers, and class modules.
CREATE TABLE udfs (
    udf_id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    module TEXT NOT NULL,
    param_count INTEGER NOT NULL,
    param_names_json TEXT NOT NULL,
    declared_volatile BOOLEAN NOT NULL,
    source_text TEXT NOT NULL,
    source_hash TEXT NOT NULL,
    UNIQUE(name, module)
);
CREATE INDEX idx_udfs_name ON udfs(name);

-- ==========================================================================
-- Phase 3: Full VBA schema — modules, procedures, edges, cell refs, chunks
-- ==========================================================================

-- VBA Modules (standard, class, form, document)
CREATE TABLE vba_modules (
    module_id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    kind TEXT NOT NULL CHECK(kind IN ('standard', 'class', 'form', 'document')),
    source_sha256 TEXT NOT NULL,
    source_text TEXT NOT NULL,
    security_findings_json TEXT NOT NULL DEFAULT '[]',
    UNIQUE(name)
);

-- VBA Procedures — supersedes udfs table for all procedure kinds.
-- The udfs table remains populated for Phase 1 backward compat.
--
-- compile_branch captures the surrounding #If condition stack so that
-- conditional-compilation twins (e.g. the 32-bit and 64-bit variants of an
-- API-bound function inside #If Win64 / #Else) can both be stored. It is the
-- empty string for unconditional procedures (the common case). Stored as
-- TEXT NOT NULL with a default of '' rather than NULL because SQLite treats
-- NULLs as distinct in UNIQUE, which would silently allow real duplicates.
CREATE TABLE vba_procedures (
    procedure_id INTEGER PRIMARY KEY,
    module_id INTEGER NOT NULL REFERENCES vba_modules(module_id),
    name TEXT NOT NULL,
    kind TEXT NOT NULL CHECK(kind IN ('sub', 'function', 'property_get', 'property_let', 'property_set')),
    signature TEXT NOT NULL,
    parameters_json TEXT NOT NULL DEFAULT '[]',
    return_type TEXT,
    body TEXT NOT NULL,
    normalized_body_hash TEXT NOT NULL,
    is_public INTEGER NOT NULL DEFAULT 1,
    is_event_handler INTEGER NOT NULL DEFAULT 0,
    event_trigger_spec_json TEXT,
    line_start INTEGER,
    line_end INTEGER,
    compile_branch TEXT NOT NULL DEFAULT '',
    UNIQUE(module_id, name, kind, compile_branch)
);
CREATE INDEX idx_vba_procedures_name ON vba_procedures(name);
CREATE INDEX idx_vba_procedures_module ON vba_procedures(module_id);
CREATE INDEX idx_vba_procedures_kind ON vba_procedures(kind);
CREATE INDEX idx_vba_procedures_event ON vba_procedures(is_event_handler) WHERE is_event_handler = 1;

-- VBA Procedure Call Graph
-- edge_kind: 'calls' (direct call), 'reads_property', 'writes_property'
-- precision: 'exact' (static name match), 'static_only' (compile-time resolvable),
--            'dynamic' (runtime dispatch, e.g. CallByName, Application.Run)
CREATE TABLE vba_procedure_edges (
    edge_id INTEGER PRIMARY KEY,
    from_procedure_id INTEGER NOT NULL REFERENCES vba_procedures(procedure_id),
    to_procedure_id INTEGER NOT NULL REFERENCES vba_procedures(procedure_id),
    edge_kind TEXT NOT NULL CHECK(edge_kind IN ('calls', 'reads_property', 'writes_property')),
    precision TEXT NOT NULL CHECK(precision IN ('exact', 'static_only', 'dynamic')),
    UNIQUE(from_procedure_id, to_procedure_id, edge_kind)
);

-- VBA Procedure Cell References (cells that a procedure reads/writes)
-- target_kind discriminator prevents downstream consumers from pattern-matching strings.
CREATE TABLE vba_procedure_cell_refs (
    ref_id INTEGER PRIMARY KEY,
    procedure_id INTEGER NOT NULL REFERENCES vba_procedures(procedure_id),
    ref_kind TEXT NOT NULL CHECK(ref_kind IN ('read', 'write', 'read_write')),
    target_kind TEXT NOT NULL CHECK(target_kind IN ('cell_range', 'named_range', 'external_ref', 'table_ref', 'row_col')),
    target TEXT NOT NULL,
    precision TEXT NOT NULL CHECK(precision IN ('exact', 'static_only', 'dynamic'))
);
CREATE INDEX idx_vba_cell_refs_proc ON vba_procedure_cell_refs(procedure_id);
CREATE INDEX idx_vba_cell_refs_target ON vba_procedure_cell_refs(target);

-- vba_module_declarations removed (WI-5): populated but never read in production

-- VBA Procedure Chunks (for retrieval-style LLM prompting)
-- One chunk per procedure up to a size cap, then split at outermost block boundary.
-- Never split mid-statement. Metadata supports retrieval + reranking.
CREATE TABLE vba_chunks (
    chunk_id INTEGER PRIMARY KEY,
    procedure_id INTEGER NOT NULL REFERENCES vba_procedures(procedure_id),
    chunk_index INTEGER NOT NULL,
    chunk_text TEXT NOT NULL,
    line_start INTEGER NOT NULL,
    line_end INTEGER NOT NULL,
    enclosing_block_kind TEXT,
    identifier_tokens_json TEXT NOT NULL DEFAULT '[]',
    comment_tokens_json TEXT NOT NULL DEFAULT '[]',
    referenced_cells_json TEXT,
    called_procedures_json TEXT
);
CREATE INDEX idx_vba_chunks_proc ON vba_chunks(procedure_id);

-- Reverse index: cells → UDFs that they call (via formula text scan)
-- Phase 3 optimization of the Phase 1 query-time regex scan.
CREATE TABLE cell_udf_calls (
    cell_id INTEGER NOT NULL REFERENCES cells(cell_id),
    udf_id INTEGER NOT NULL REFERENCES udfs(udf_id),
    PRIMARY KEY (cell_id, udf_id)
) WITHOUT ROWID;
CREATE INDEX idx_cell_udf_calls_udf ON cell_udf_calls(udf_id);

-- ==========================================================================
-- Data-validation rules + cell comments (2026-07-05 scope-gap channels #1/#2)
-- ==========================================================================
-- Additive tables; older DBs lack them — consumers MUST capability-check
-- (sqlite_master) before querying. Verbatim storage: sqref / formulas are
-- stored as written; range/name-sourced lists are NOT resolved to values.

-- Modeller-declared input constraints. source discriminates the native
-- <dataValidations> form from the x14:dataValidation form inside <extLst>
-- (the corpus's scenario-selection wiring exists ONLY as x14).
CREATE TABLE data_validations (
    validation_id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL REFERENCES sheets(sheet_id),
    sqref TEXT NOT NULL,      -- verbatim; may contain multiple space-separated ranges
    val_type TEXT,            -- list / whole / decimal / date / textLength / custom / ...
    operator TEXT,            -- NULL = spec default ("between" for bounded types)
    formula1 TEXT,
    formula2 TEXT,
    -- JSON array parsed from a quoted literal-list formula1 like '"a, b, c"'
    -- (outer quotes stripped, comma-split, trimmed). NULL for range /
    -- defined-name / expression sources.
    literal_values_json TEXT,
    allow_blank INTEGER NOT NULL DEFAULT 0,
    prompt_title TEXT,
    prompt_text TEXT,
    source TEXT NOT NULL CHECK (source IN ('native', 'x14'))
);
CREATE INDEX idx_data_validations_sheet ON data_validations(sheet_id);

-- Cell notes: classic xl/comments*.xml + threaded xl/threadedComments/*.
-- Classic placeholder mirrors of threaded comments are dropped at extraction
-- (the threaded part is canonical). thread_order is the 0-based position
-- within a cell's thread (dT timestamp ordering); always 0 for classic.
CREATE TABLE cell_comments (
    comment_id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL REFERENCES sheets(sheet_id),
    a1 TEXT NOT NULL,
    author TEXT,
    text TEXT NOT NULL,
    kind TEXT NOT NULL CHECK (kind IN ('classic', 'threaded')),
    thread_order INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX idx_cell_comments_sheet ON cell_comments(sheet_id);
