# ABOUTME: Frozen dataclasses for the IR diff algorithm's in-memory model.
# ABOUTME: All entities use canonical textual keys (no database-local IDs).

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field

from xl_marinade.errors import MarinadeError

# ---------------------------------------------------------------------------
# Canonical entity keys
# ---------------------------------------------------------------------------


@dataclass(frozen=True, order=True)
class CellKey:
    """Canonical cell identity: (sheet_name, row, col)."""

    sheet: str
    row: int
    col: int


@dataclass(frozen=True, order=True)
class BindingKey:
    """Canonical binding position key."""

    sheet: str
    top_left_row: int
    top_left_col: int
    shape_rows: int
    shape_cols: int


# ---------------------------------------------------------------------------
# Cell-level data
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class CellSig:
    """Full cell attribute vector for comparison."""

    formula_r1c1: str | None
    value_sha256: str | None
    value_json: str | None
    format_sha256: str | None
    format_json: str | None
    data_type: str | None
    is_array_formula: bool
    is_spilled: bool
    spill_origin: CellKey | None


@dataclass(frozen=True)
class CellSigLite:
    """Lightweight cell signature for binding fingerprinting.

    Excludes absolute coordinates and value snapshots — only structural identity.
    """

    formula_r1c1: str | None
    data_type: str | None
    is_array_formula: bool
    is_spilled: bool


# ---------------------------------------------------------------------------
# Binding-level data
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class BindingDesc:
    """Full binding attribute vector."""

    key: BindingKey
    binding_type: str
    formula_r1c1: str | None
    label: str | None
    classification: str | None
    confidence: float | None
    is_orphan: bool
    extraction_source: str | None
    evidence_sha256: str | None
    evidence_json: str | None
    spatial_sha256: str | None
    spatial_json: str | None
    address_a1: str
    # member cells by (row_offset, col_offset) from top-left
    members_by_offset: dict[tuple[int, int], CellSigLite] = field(default_factory=dict)
    # original binding_id from the database (for lineage tracking, not matching)
    original_binding_id: str = ""

    @property
    def binding_fp_rel(self) -> str:
        """Relative fingerprint: hash of sorted member-offset + CellSigLite tuples."""
        items = sorted(
            (
                dr,
                dc,
                sig.formula_r1c1 or "",
                sig.data_type or "",
                sig.is_array_formula,
                sig.is_spilled,
            )
            for (dr, dc), sig in self.members_by_offset.items()
        )
        raw = json.dumps(items, separators=(",", ":"), sort_keys=True)
        return hashlib.sha256(raw.encode()).hexdigest()[:16]


# ---------------------------------------------------------------------------
# Named entities
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class NameDesc:
    """Defined name descriptor."""

    name: str
    scope: str
    destinations: str  # JSON list of A1 refs
    is_external: bool


@dataclass(frozen=True)
class TableDesc:
    """Table candidate descriptor."""

    candidate_id: str
    sheet: str
    kind: str  # 'vector' or 'grid'
    r1: int
    c1: int
    r2: int
    c2: int
    range_a1: str
    confidence: float
    reasons_json: str
    members: tuple[tuple[int, str, str | None], ...] = ()  # (ordinal, binding_key_repr, role_hint)


# ---------------------------------------------------------------------------
# Edge tuples
# ---------------------------------------------------------------------------


@dataclass(frozen=True, order=True)
class CellEdge:
    """Internal cell-to-cell dependency edge."""

    from_key: CellKey
    to_key: CellKey


@dataclass(frozen=True, order=True)
class ExternalEdge:
    """Cell-to-external dependency edge."""

    from_key: CellKey
    external_ref: str


@dataclass(frozen=True, order=True)
class RangeEdge:
    """Cell-to-range dependency edge."""

    from_key: CellKey
    to_sheet: str
    to_r1: int
    to_c1: int
    to_r2: int
    to_c2: int
    to_range_a1: str
    cell_count: int


@dataclass(frozen=True, order=True)
class BindingEdgeTuple:
    """Binding-to-binding dependency edge (matched by position key, not ID)."""

    from_key: BindingKey
    to_key: BindingKey
    edge_count: int


# ---------------------------------------------------------------------------
# Formula families
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class FamilyDesc:
    """Formula family descriptor."""

    sheet: str
    formula_r1c1: str
    member_count: int
    representative_binding_key: BindingKey
    member_binding_keys: tuple[BindingKey, ...] = ()
    original_family_id: str = ""


# ---------------------------------------------------------------------------
# Label evidence, time annotations, resolution metrics
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class LabelEvidence:
    """Binding label candidate cell evidence row."""

    binding_key: BindingKey
    candidate_type: str
    candidate_address: str
    cell_address: str
    sheet: str
    row: int
    col: int
    value_text: str | None


@dataclass(frozen=True)
class TimeIndexCandidate:
    """Time index candidate row."""

    sheet: str
    binding_key: BindingKey
    rank: int
    confidence: float
    reasons_json: str


@dataclass(frozen=True)
class BindingTimeAnnotation:
    """Binding time annotation row."""

    binding_key: BindingKey
    time_index_binding_key: BindingKey
    is_time_dependent: bool
    confidence: float
    reasons_json: str
    evidence_flags_json: str


@dataclass(frozen=True)
class ResolutionMetric:
    """Resolution metric row."""

    function_name: str
    status: str
    count: int


# ---------------------------------------------------------------------------
# Root and metadata
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class UserRoot:
    """User root specification."""

    sheet: str
    range_a1: str
    label_hint: str | None


# ---------------------------------------------------------------------------
# The complete in-memory IR model
# ---------------------------------------------------------------------------


@dataclass
class IRModel:
    """Complete in-memory representation of one IR database."""

    # Metadata
    metadata: dict[str, str] = field(default_factory=dict)
    roots: list[UserRoot] = field(default_factory=list)

    # Sheets
    sheet_names: list[str] = field(default_factory=list)

    # Cells: CellKey -> CellSig
    cells: dict[CellKey, CellSig] = field(default_factory=dict)

    # Bindings: BindingKey -> BindingDesc
    bindings: dict[BindingKey, BindingDesc] = field(default_factory=dict)

    # Cell-to-binding membership: CellKey -> list of BindingKey
    cell_to_binding: dict[CellKey, list[BindingKey]] = field(default_factory=dict)

    # Edges
    cell_edges: set[CellEdge] = field(default_factory=set)
    external_edges: set[ExternalEdge] = field(default_factory=set)
    range_edges: set[RangeEdge] = field(default_factory=set)
    binding_edges: set[BindingEdgeTuple] = field(default_factory=set)

    # Named entities
    names: dict[tuple[str, str], NameDesc] = field(
        default_factory=dict
    )  # (name, scope) -> NameDesc
    tables: dict[str, TableDesc] = field(default_factory=dict)  # candidate_id -> TableDesc

    # Families: (sheet, formula_r1c1) -> FamilyDesc
    families: dict[tuple[str, str], FamilyDesc] = field(default_factory=dict)

    # Label evidence: frozenset of LabelEvidence
    label_evidence: set[LabelEvidence] = field(default_factory=set)

    # Time annotations
    time_index_candidates: list[TimeIndexCandidate] = field(default_factory=list)
    binding_time_annotations: dict[BindingKey, BindingTimeAnnotation] = field(default_factory=dict)

    # Resolution metrics: (function_name, status) -> count
    resolution_metrics: dict[tuple[str, str], int] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Matching results
# ---------------------------------------------------------------------------


@dataclass
class SheetMatch:
    """Result of sheet matching phase."""

    # mu: A sheet name -> B sheet name (matched pairs)
    matched: dict[str, str] = field(default_factory=dict)
    # Renamed pairs: (a_name, b_name, score_bind, score_cell, score_coord)
    renamed: list[tuple[str, str, float, float, float]] = field(default_factory=list)
    # Unmatched sheets
    removed: list[str] = field(default_factory=list)  # in A only
    added: list[str] = field(default_factory=list)  # in B only


@dataclass
class AxisMap:
    """Row/column maps for one matched sheet pair."""

    # rho: A_row -> B_row (None if row deleted)
    row_map: dict[int, int | None] = field(default_factory=dict)
    # kappa: A_col -> B_col (None if col deleted)
    col_map: dict[int, int | None] = field(default_factory=dict)
    # Detected row insertions: list of (at_row_in_B, count)
    rows_inserted: list[tuple[int, int]] = field(default_factory=list)
    rows_deleted: list[tuple[int, int]] = field(default_factory=list)
    cols_inserted: list[tuple[int, int]] = field(default_factory=list)
    cols_deleted: list[tuple[int, int]] = field(default_factory=list)


@dataclass
class BindingMatch:
    """Result of binding matching phase."""

    # Matched binding pairs: A BindingKey -> B BindingKey
    matched: dict[BindingKey, BindingKey] = field(default_factory=dict)
    # Classification of match type
    match_type: dict[BindingKey, str] = field(
        default_factory=dict
    )  # A key -> 'exact'|'moved'|'overlap'
    # Unmatched bindings
    removed: list[BindingKey] = field(default_factory=list)  # in A only
    added: list[BindingKey] = field(default_factory=list)  # in B only
    # Lineage metadata for splits/merges
    lineage: dict[BindingKey, dict] = field(default_factory=dict)


@dataclass
class CellMatch:
    """Result of cell matching phase."""

    # Matched cell pairs: A CellKey -> B CellKey
    matched: dict[CellKey, CellKey] = field(default_factory=dict)
    # Unmatched cells (not reported individually — covered by binding events)
    removed: list[CellKey] = field(default_factory=list)
    added: list[CellKey] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Change record
# ---------------------------------------------------------------------------


@dataclass
class Change:
    """A single change in the diff output."""

    type: str
    # Canonical sort key for deterministic ordering within a tier
    sort_key: tuple = ()
    # Human-readable fields (vary by change type)
    details: dict = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------


class DiffVerificationError(MarinadeError):
    """Raised when the verification pass detects an invariant violation."""

    pass
