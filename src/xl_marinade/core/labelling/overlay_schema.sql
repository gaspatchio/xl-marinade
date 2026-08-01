-- ABOUTME: SQLite schema for semantic overlay database v0.3
-- ABOUTME: Defines semantic_variables, mutation_log, metadata, and composite binding tables with indexes

-- Enable foreign key enforcement
PRAGMA foreign_keys = ON;

-- Table: semantic_variables
-- Stores semantic enrichment for each binding
-- Note: binding_id references Phase 1 IR bindings table (separate DB)
-- FK constraint not enforceable cross-database; validated via integrity checks
CREATE TABLE semantic_variables (
    binding_id TEXT PRIMARY KEY,
    label TEXT,
    label_source INTEGER REFERENCES mutation_log(mutation_id),
    actuarial_class TEXT,
    actuarial_class_reasoning TEXT,
    actuarial_class_confidence REAL CHECK (actuarial_class_confidence >= 0 AND actuarial_class_confidence <= 1 OR actuarial_class_confidence IS NULL),
    is_active BOOLEAN NOT NULL DEFAULT 1,
    is_composite BOOLEAN NOT NULL DEFAULT 0,
    superseded_by INTEGER REFERENCES mutation_log(mutation_id),
    
    -- Reconciliation classification (Story 23)
    reconciliation_required INTEGER DEFAULT 0,
    reconciliation_rationale TEXT,
    
    -- Confidence scores (Story 14)
    label_confidence REAL CHECK (label_confidence >= 0 AND label_confidence <= 1 OR label_confidence IS NULL),
    classification_confidence REAL CHECK (classification_confidence >= 0 AND classification_confidence <= 1 OR classification_confidence IS NULL),
    
    -- Orphan indicator (Story 15)
    is_orphan BOOLEAN NOT NULL DEFAULT 0
);

CREATE INDEX idx_semantic_variables_label ON semantic_variables(label);
CREATE INDEX idx_semantic_variables_label_source ON semantic_variables(label_source);
CREATE INDEX idx_semantic_variables_active ON semantic_variables(is_active);
CREATE INDEX idx_semantic_variables_actuarial_class ON semantic_variables(actuarial_class);
CREATE INDEX idx_semantic_variables_is_orphan ON semantic_variables(is_orphan);

-- Table: composite_bindings
-- Tracks which IR bindings compose each virtual/composite binding
CREATE TABLE composite_bindings (
    composite_id TEXT NOT NULL,
    ir_binding_id TEXT NOT NULL,
    ordinal INTEGER NOT NULL,
    PRIMARY KEY (composite_id, ir_binding_id),
    FOREIGN KEY (composite_id) REFERENCES semantic_variables(binding_id)
);

CREATE INDEX idx_composite_bindings_composite_id ON composite_bindings(composite_id);
CREATE INDEX idx_composite_bindings_ir_binding_id ON composite_bindings(ir_binding_id);

-- Table: mutation_log
-- Complete history of applied mutations
CREATE TABLE mutation_log (
    mutation_id INTEGER PRIMARY KEY,
    timestamp TEXT NOT NULL,
    action TEXT NOT NULL,
    parameters_json TEXT NOT NULL,
    metadata_json TEXT,
    
    CHECK (mutation_id > 0),
    CHECK (json_valid(parameters_json)),
    CHECK (metadata_json IS NULL OR json_valid(metadata_json))
);

CREATE INDEX idx_mutation_log_timestamp ON mutation_log(timestamp);
CREATE INDEX idx_mutation_log_action ON mutation_log(action);

-- Table: metadata
-- Overlay database metadata (provenance)
CREATE TABLE metadata (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
