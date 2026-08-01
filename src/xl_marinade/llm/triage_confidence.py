# ABOUTME: Deterministic triage confidence scorer for LLM targeting (Sprint 7)
# ABOUTME: Conservative scoring to identify bindings needing enrichment - NOT post-hoc quality assessment

"""
Triage Confidence Scoring Module (Sprint 7)
============================================

This module implements a **deterministic triage confidence scorer** that produces
targeting signals for LLM enrichment. This is fundamentally different from the
existing post-hoc confidence assessment in confidence_scorer.py.

## Purpose

The triage scorer answers: "Should we spend LLM tokens enriching this binding?"
- Low score → likely wrong/ambiguous → enrich
- High score → likely correct → skip enrichment

## Conservative Philosophy (Sprint 7 requirement)

High confidence is **hard to earn**:
- We only assign high confidence when we have strong, corroborating deterministic evidence
- Absence of evidence is treated as uncertainty (lower confidence), not as confidence
- Default behavior is pessimistic: start low, add confidence only when evidence supports correctness

## Confidence Thresholds

- score >= 0.85: "high confidence" (rare; strong evidence; skip enrichment)
- 0.60 <= score < 0.85: "medium confidence" (review/enrich optional)
- score < 0.60: "low confidence" (eligible for enrichment)

## Two-Score Model

We maintain two separate conservative scores because labels and structure fail for different reasons:

1. **Label confidence** (label_score):
   - High confidence requires multiple corroborating signals:
     - Header text strongly matches actuarial concept AND
     - Formula token patterns align with that concept AND
     - Label is not generic/boilerplate AND
     - (optional) Neighborhood labels are coherent
   - Penalties: generic labels, very short labels, conflicting cues, heterogeneous formulas

2. **Structural/binding confidence** (structural_score):
   - High confidence requires strong layout/geometry evidence:
     - Clean rectangle with high density
     - Consistent formula "shape" across binding
     - No overlaps with neighbors
     - Headers/axes are consistent
   - Penalties: ragged edges, islands, multiple formula clusters, overlaps, single-cell bindings

## Reason Codes (v1 contract)

Reason codes are a stable contract used by:
- Integration tests (exact selection expectations)
- enrichment_audit.jsonl (evaluation)
- Future tuning without losing explainability

**Ordering contract (deterministic):**
- Return reason_codes[] sorted by:
  1) descending absolute impact on score (|Δ|), then
  2) stable lexicographic order by code

**Format contract:**
- Upper snake case strings (e.g., GENERIC_LABEL_TOKEN)
- Codes are versioned implicitly by this story; changes require fixture + test updates

## Integration

This module is called during Sprint 7 enrichment pipeline:
1. Extract IR (fast mode)
2. Run deterministic labeling/classification
3. **Triage scoring** (this module) → select bindings for enrichment
4. LLM enrichment (only for selected bindings)
5. Replay mutations → overlay
6. Downstream outputs

See: docs/phase2_documentation_agent/backlog/sprint7/SPRINT7_PLAN.md
"""

import json
import logging
import sqlite3
from dataclasses import dataclass

logger = logging.getLogger(__name__)

# Generic label tokens (penalty)
GENERIC_TOKENS = {
    "input",
    "value",
    "total",
    "date",
    "calc",
    "result",
    "data",
    "output",
    "number",
    "amount",
    "field",
}

# Minimum label length (non-space chars)
MIN_LABEL_LENGTH = 3

# Thresholds for confidence bands
HIGH_CONFIDENCE_THRESHOLD = 0.85
MEDIUM_CONFIDENCE_THRESHOLD = 0.60


@dataclass
class TriageConfidence:
    """Triage confidence for a single binding."""

    binding_id: str
    score: float  # Overall triage score [0.0, 1.0]
    label_score: float  # Label-specific score [0.0, 1.0]
    structural_score: float  # Structural-specific score [0.0, 1.0]
    reasons: list[str]  # Ordered reason codes (deterministic)


def _fetch_cell_positions(binding_id: str, ir_conn: sqlite3.Connection) -> list[tuple[int, int]]:
    """
    Fetch (row, col) positions for cells in a binding across supported IR schemas.

    Tries:
    - agent_cells(row_idx, col_idx) + cell_to_binding
    - cells(row, col) + cell_to_binding (fast schema)
    """
    try:
        rows = ir_conn.execute(
            """
            SELECT ac.row_idx, ac.col_idx
            FROM cell_to_binding ctb
            JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
            WHERE ctb.binding_id = ?
            """,
            (binding_id,),
        ).fetchall()
        if rows:
            return [(int(r[0]), int(r[1])) for r in rows if r[0] is not None and r[1] is not None]
    except sqlite3.OperationalError:
        pass

    try:
        rows = ir_conn.execute(
            """
            SELECT c.row, c.col
            FROM cell_to_binding ctb
            JOIN cells c ON ctb.cell_id = c.cell_id
            WHERE ctb.binding_id = ?
            """,
            (binding_id,),
        ).fetchall()
        return [(int(r[0]), int(r[1])) for r in rows if r[0] is not None and r[1] is not None]
    except sqlite3.OperationalError:
        return []


def _fetch_r1c1_patterns(
    binding_id: str, ir_conn: sqlite3.Connection, limit: int = 50
) -> list[str]:
    """
    Fetch R1C1 formulas/patterns for a binding across supported IR schemas.

    Tries:
    - agent_cells(formula_r1c1)
    - formulas(formula_r1c1) via cells(formula_id) (fast schema)
    """
    try:
        rows = ir_conn.execute(
            """
            SELECT ac.formula_r1c1
            FROM cell_to_binding ctb
            JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
            WHERE ctb.binding_id = ? AND ac.formula_r1c1 IS NOT NULL
            LIMIT ?
            """,
            (binding_id, limit),
        ).fetchall()
        if rows:
            return [r[0] for r in rows if r and r[0]]
    except sqlite3.OperationalError:
        pass

    try:
        rows = ir_conn.execute(
            """
            SELECT f.formula_r1c1
            FROM cell_to_binding ctb
            JOIN cells c ON ctb.cell_id = c.cell_id
            JOIN formulas f ON c.formula_id = f.formula_id
            WHERE ctb.binding_id = ? AND f.formula_r1c1 IS NOT NULL
            LIMIT ?
            """,
            (binding_id, limit),
        ).fetchall()
        return [r[0] for r in rows if r and r[0]]
    except sqlite3.OperationalError:
        return []


def _fetch_formula_examples(
    binding_id: str, ir_conn: sqlite3.Connection, limit: int = 10
) -> list[str]:
    """
    Fetch formula examples for function-token heuristics across supported IR schemas.

    Tries:
    - agent_cells(formula)
    - formulas(formula_a1_example) via cells(formula_id) (fast schema)
    """
    try:
        rows = ir_conn.execute(
            """
            SELECT DISTINCT ac.formula
            FROM cell_to_binding ctb
            JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
            WHERE ctb.binding_id = ? AND ac.formula IS NOT NULL
            LIMIT ?
            """,
            (binding_id, limit),
        ).fetchall()
        if rows:
            return [r[0] for r in rows if r and r[0]]
    except sqlite3.OperationalError:
        pass

    try:
        rows = ir_conn.execute(
            """
            SELECT DISTINCT f.formula_a1_example
            FROM cell_to_binding ctb
            JOIN cells c ON ctb.cell_id = c.cell_id
            JOIN formulas f ON c.formula_id = f.formula_id
            WHERE ctb.binding_id = ? AND f.formula_a1_example IS NOT NULL
            LIMIT ?
            """,
            (binding_id, limit),
        ).fetchall()
        return [r[0] for r in rows if r and r[0]]
    except sqlite3.OperationalError:
        return []


# ============================================================================
# Label Confidence Heuristics
# ============================================================================


def check_generic_label_token(label: str | None) -> tuple[bool, float]:
    """
    Check if label contains generic tokens.

    Args:
        label: Binding label

    Returns:
        (has_penalty, impact) tuple
    """
    if not label:
        return False, 0.0

    label_lower = label.lower()
    tokens = set(label_lower.replace("_", " ").split())

    if tokens & GENERIC_TOKENS:
        return True, -0.25  # Significant penalty
    return False, 0.0


def check_label_too_short(label: str | None) -> tuple[bool, float]:
    """
    Check if label is too short.

    Args:
        label: Binding label

    Returns:
        (has_penalty, impact) tuple
    """
    if not label:
        return False, 0.0

    non_space_chars = len(label.replace(" ", "").replace("_", ""))
    if non_space_chars < MIN_LABEL_LENGTH:
        return True, -0.15
    return False, 0.0


def check_header_match_strong(
    binding_id: str, label: str | None, ir_conn: sqlite3.Connection
) -> tuple[bool, float]:
    """
    Check if header text strongly matches label.

    Args:
        binding_id: Binding ID
        label: Binding label
        ir_conn: IR database connection

    Returns:
        (has_boost, impact) tuple
    """
    if not label:
        return False, 0.0

    try:
        # Get spatial candidates from IR
        candidates_row = ir_conn.execute(
            """
            SELECT spatial_candidates
            FROM agent_bindings
            WHERE binding_id = ?
        """,
            (binding_id,),
        ).fetchone()

        if not candidates_row or not candidates_row[0]:
            return False, 0.0

        candidates_data = json.loads(candidates_row[0])
        if not isinstance(candidates_data, dict):
            return False, 0.0

        label_candidates = candidates_data.get("label_candidates", [])
        if not label_candidates:
            return False, 0.0

        # Check if label exactly matches a high-scoring candidate
        for cand in label_candidates:
            if isinstance(cand, dict):
                cand_label = cand.get("label", "")
                cand_score = cand.get("score", 0.0)

                if cand_label == label and cand_score >= 0.8:
                    # Strong header match
                    return True, +0.20

        return False, 0.0

    except Exception as e:
        logger.warning(f"Error checking header match for {binding_id}: {e}")
        return False, 0.0


def check_formula_tokens_coherent(
    binding_id: str, ir_conn: sqlite3.Connection
) -> tuple[bool, float]:
    """
    Check if formula function tokens are coherent across binding.

    Args:
        binding_id: Binding ID
        ir_conn: IR database connection

    Returns:
        (has_boost, impact) tuple
    """
    try:
        formulas = _fetch_formula_examples(binding_id, ir_conn, limit=10)
        if not formulas or len(formulas) < 2:
            return False, 0.0  # Need multiple formulas to check coherence

        # Extract function tokens from each formula
        function_sets = []
        for formula in formulas:
            upper = formula.upper()
            # Simple extraction: look for common functions
            functions = set()
            for func in ["SUM", "AVERAGE", "IF", "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "XLOOKUP"]:
                if func in upper:
                    functions.add(func)
            function_sets.append(functions)

        if not any(function_sets):
            return False, 0.0  # No functions found

        # Check if all formulas use similar function families
        # (e.g., all use SUM, or all use lookups)
        common_functions = set.intersection(*[s for s in function_sets if s])
        if common_functions:
            return True, +0.15  # Coherent function usage

        return False, 0.0

    except Exception as e:
        logger.warning(f"Error checking formula coherence for {binding_id}: {e}")
        return False, 0.0


def check_header_missing_or_weak(
    binding_id: str, ir_conn: sqlite3.Connection
) -> tuple[bool, float]:
    """
    Check if header candidates are missing or weak.

    Args:
        binding_id: Binding ID
        ir_conn: IR database connection

    Returns:
        (has_penalty, impact) tuple
    """
    try:
        candidates_row = ir_conn.execute(
            """
            SELECT spatial_candidates
            FROM agent_bindings
            WHERE binding_id = ?
        """,
            (binding_id,),
        ).fetchone()

        if not candidates_row or not candidates_row[0]:
            return True, -0.20  # No candidates at all

        candidates_data = json.loads(candidates_row[0])
        if not isinstance(candidates_data, dict):
            return True, -0.20

        label_candidates = candidates_data.get("label_candidates", [])
        if not label_candidates:
            return True, -0.20  # No label candidates

        # Check if best candidate has weak score
        best_score = max(
            (c.get("score", 0.0) for c in label_candidates if isinstance(c, dict)), default=0.0
        )

        if best_score < 0.5:
            return True, -0.15  # Weak candidates

        return False, 0.0

    except Exception as e:
        logger.warning(f"Error checking header quality for {binding_id}: {e}")
        return True, -0.20


def check_heterogeneous_formula_tokens(
    binding_id: str, ir_conn: sqlite3.Connection
) -> tuple[bool, float]:
    """
    Check if formulas show multiple distinct token families.

    Args:
        binding_id: Binding ID
        ir_conn: IR database connection

    Returns:
        (has_penalty, impact) tuple
    """
    try:
        formulas = ir_conn.execute(
            """
            SELECT DISTINCT ac.formula
            FROM cell_to_binding ctb
            JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
            WHERE ctb.binding_id = ? AND ac.formula IS NOT NULL
            LIMIT 20
        """,
            (binding_id,),
        ).fetchall()

        if not formulas or len(formulas) < 3:
            return False, 0.0  # Need multiple formulas to detect heterogeneity

        # Categorize formulas into families
        lookup_count = 0
        aggregation_count = 0
        arithmetic_count = 0

        for formula_row in formulas:
            formula = formula_row[0].upper()

            if any(f in formula for f in ["VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "XLOOKUP"]):
                lookup_count += 1
            elif any(f in formula for f in ["SUM", "AVERAGE", "SUMPRODUCT"]):
                aggregation_count += 1
            elif any(op in formula for op in ["+", "-", "*", "/"]):
                arithmetic_count += 1

        # Check if multiple families are present
        families_present = sum([lookup_count > 0, aggregation_count > 0, arithmetic_count > 0])

        if families_present >= 2:
            return True, -0.15  # Heterogeneous formulas

        return False, 0.0

    except Exception as e:
        logger.warning(f"Error checking formula heterogeneity for {binding_id}: {e}")
        return False, 0.0


# ============================================================================
# Structural Confidence Heuristics
# ============================================================================


def check_single_cell_binding(binding_id: str, ir_conn: sqlite3.Connection) -> tuple[bool, float]:
    """
    Check if binding is a single cell.

    Args:
        binding_id: Binding ID
        ir_conn: IR database connection

    Returns:
        (has_penalty, impact) tuple
    """
    try:
        cell_count = ir_conn.execute(
            """
            SELECT COUNT(*)
            FROM cell_to_binding
            WHERE binding_id = ?
        """,
            (binding_id,),
        ).fetchone()[0]

        if cell_count == 1:
            return True, -0.20  # Single-cell bindings are often ambiguous

        return False, 0.0

    except Exception as e:
        logger.warning(f"Error checking cell count for {binding_id}: {e}")
        return False, 0.0


def check_rectangle_dense(binding_id: str, ir_conn: sqlite3.Connection) -> tuple[bool, float]:
    """
    Check if binding is a dense rectangle.

    Args:
        binding_id: Binding ID
        ir_conn: IR database connection

    Returns:
        (has_boost, impact) tuple
    """
    try:
        cells = _fetch_cell_positions(binding_id, ir_conn)
        if len(cells) < 2:
            return False, 0.0
        rows = [r for r, _ in cells]
        cols = [c for _, c in cells]
        min_row, max_row = min(rows), max(rows)
        min_col, max_col = min(cols), max(cols)

        bbox_cells = (max_row - min_row + 1) * (max_col - min_col + 1)
        if bbox_cells <= 0:
            return False, 0.0

        density = len(cells) / bbox_cells

        if density >= 0.9:
            return True, +0.15  # High density rectangle

        return False, 0.0

    except Exception as e:
        logger.warning(f"Error checking density for {binding_id}: {e}")
        return False, 0.0


def check_formula_shape_uniform(binding_id: str, ir_conn: sqlite3.Connection) -> tuple[bool, float]:
    """
    Check if R1C1/formula patterns are uniform across binding.

    Args:
        binding_id: Binding ID
        ir_conn: IR database connection

    Returns:
        (has_boost, impact) tuple
    """
    try:
        r1c1_formulas = _fetch_r1c1_patterns(binding_id, ir_conn, limit=20)
        if not r1c1_formulas or len(r1c1_formulas) < 2:
            return False, 0.0  # Need multiple formulas to check uniformity

        # Check if all R1C1 formulas are identical (perfect uniformity)
        unique_r1c1 = set(r1c1_formulas)

        if len(unique_r1c1) == 1:
            return True, +0.15  # Perfect uniformity
        elif len(unique_r1c1) <= 2:
            return True, +0.10  # Near uniformity (e.g., edge cases)

        return False, 0.0

    except Exception as e:
        logger.warning(f"Error checking formula uniformity for {binding_id}: {e}")
        return False, 0.0


def check_ragged_edges(binding_id: str, ir_conn: sqlite3.Connection) -> tuple[bool, float]:
    """
    Check for irregular boundary shape / islands.

    Args:
        binding_id: Binding ID
        ir_conn: IR database connection

    Returns:
        (has_penalty, impact) tuple
    """
    try:
        cells = _fetch_cell_positions(binding_id, ir_conn)

        if len(cells) < 4:
            return False, 0.0  # Too small to detect ragged edges

        # Get bounding box
        rows = [c[0] for c in cells]
        cols = [c[1] for c in cells]
        min_row, max_row = min(rows), max(rows)
        min_col, max_col = min(cols), max(cols)

        # Calculate density
        bbox_cells = (max_row - min_row + 1) * (max_col - min_col + 1)
        actual_cells = len(cells)
        density = actual_cells / bbox_cells if bbox_cells > 0 else 0.0

        if density < 0.6:
            return True, -0.15  # Low density suggests ragged edges

        return False, 0.0

    except Exception as e:
        logger.warning(f"Error checking ragged edges for {binding_id}: {e}")
        return False, 0.0


def check_multiple_formula_clusters(
    binding_id: str, ir_conn: sqlite3.Connection
) -> tuple[bool, float]:
    """
    Check if binding contains multiple distinct formula pattern clusters.

    Args:
        binding_id: Binding ID
        ir_conn: IR database connection

    Returns:
        (has_penalty, impact) tuple
    """
    try:
        r1c1_patterns = _fetch_r1c1_patterns(binding_id, ir_conn, limit=50)
        if not r1c1_patterns:
            return False, 0.0
        unique_patterns = len(set(r1c1_patterns))

        # If more than 3 distinct patterns, likely multiple clusters
        if unique_patterns > 3:
            return True, -0.15

        return False, 0.0

    except Exception as e:
        logger.warning(f"Error checking formula clusters for {binding_id}: {e}")
        return False, 0.0


# ============================================================================
# Triage Confidence Calculation
# ============================================================================


def calculate_triage_confidence(
    binding_id: str, label: str | None, ir_conn: sqlite3.Connection
) -> TriageConfidence:
    """
    Calculate deterministic triage confidence for a binding.

    This is a conservative scorer: high confidence is hard to earn.

    Args:
        binding_id: Binding ID
        label: Current label (may be None)
        ir_conn: IR database connection

    Returns:
        TriageConfidence with score and reason codes
    """
    # Start from baseline (pessimistic)
    label_score = 0.50  # Neutral baseline
    structural_score = 0.50  # Neutral baseline

    # Track reasons with their impacts
    reason_impacts: list[tuple[str, float]] = []

    # ========================================================================
    # Label Confidence Checks
    # ========================================================================

    # Penalties
    has_generic, impact = check_generic_label_token(label)
    if has_generic:
        label_score += impact
        reason_impacts.append(("GENERIC_LABEL_TOKEN", impact))

    has_short, impact = check_label_too_short(label)
    if has_short:
        label_score += impact
        reason_impacts.append(("LABEL_TOO_SHORT", impact))

    has_weak_header, impact = check_header_missing_or_weak(binding_id, ir_conn)
    if has_weak_header:
        label_score += impact
        reason_impacts.append(("HEADER_MISSING_OR_WEAK", impact))

    has_hetero, impact = check_heterogeneous_formula_tokens(binding_id, ir_conn)
    if has_hetero:
        label_score += impact
        reason_impacts.append(("HETEROGENEOUS_FORMULA_TOKENS", impact))

    # Boosts
    has_strong_header, impact = check_header_match_strong(binding_id, label, ir_conn)
    if has_strong_header:
        label_score += impact
        reason_impacts.append(("HEADER_MATCH_STRONG", impact))

    has_coherent, impact = check_formula_tokens_coherent(binding_id, ir_conn)
    if has_coherent:
        label_score += impact
        reason_impacts.append(("FORMULA_TOKENS_COHERENT", impact))

    # ========================================================================
    # Structural Confidence Checks
    # ========================================================================

    # Penalties
    is_single, impact = check_single_cell_binding(binding_id, ir_conn)
    if is_single:
        structural_score += impact
        reason_impacts.append(("SINGLE_CELL_BINDING", impact))

    has_ragged, impact = check_ragged_edges(binding_id, ir_conn)
    if has_ragged:
        structural_score += impact
        reason_impacts.append(("RAGGED_EDGES", impact))

    has_clusters, impact = check_multiple_formula_clusters(binding_id, ir_conn)
    if has_clusters:
        structural_score += impact
        reason_impacts.append(("MULTIPLE_FORMULA_CLUSTERS", impact))

    # Boosts
    is_dense, impact = check_rectangle_dense(binding_id, ir_conn)
    if is_dense:
        structural_score += impact
        reason_impacts.append(("RECTANGLE_DENSE", impact))

    is_uniform, impact = check_formula_shape_uniform(binding_id, ir_conn)
    if is_uniform:
        structural_score += impact
        reason_impacts.append(("FORMULA_SHAPE_UNIFORM", impact))

    # ========================================================================
    # Combine Scores
    # ========================================================================

    # Overall score is the minimum of label and structural scores (conservative)
    overall_score = min(label_score, structural_score)

    # Clamp to [0.0, 1.0]
    overall_score = max(0.0, min(1.0, overall_score))
    label_score = max(0.0, min(1.0, label_score))
    structural_score = max(0.0, min(1.0, structural_score))

    # ========================================================================
    # Sort Reason Codes (Deterministic Ordering)
    # ========================================================================

    # Sort by: 1) descending absolute impact, 2) lexicographic order
    reason_impacts.sort(key=lambda x: (-abs(x[1]), x[0]))
    reason_codes = [code for code, _ in reason_impacts]

    return TriageConfidence(
        binding_id=binding_id,
        score=overall_score,
        label_score=label_score,
        structural_score=structural_score,
        reasons=reason_codes,
    )


def should_enrich(
    confidence: TriageConfidence, threshold: float = MEDIUM_CONFIDENCE_THRESHOLD
) -> bool:
    """
    Decide whether to enrich a binding based on triage confidence.

    Args:
        confidence: TriageConfidence for the binding
        threshold: Score threshold (default 0.60)

    Returns:
        True if binding should be enriched (score < threshold)
    """
    return confidence.score < threshold


def assess_all_triage_confidence(ir_db_path: str, overlay_db_path: str) -> list[TriageConfidence]:
    """
    Assess triage confidence for all active bindings.

    Args:
        ir_db_path: Path to IR database
        overlay_db_path: Path to semantic overlay database

    Returns:
        List of TriageConfidence for each active binding
    """
    results = []

    ir_conn = None
    overlay_conn = None

    try:
        ir_conn = sqlite3.connect(ir_db_path)
        overlay_conn = sqlite3.connect(overlay_db_path)

        # Get all active bindings with labels
        bindings = overlay_conn.execute("""
            SELECT binding_id, label
            FROM semantic_variables
            WHERE is_active = 1
        """).fetchall()

        logger.info(f"Assessing triage confidence for {len(bindings)} active bindings")

        for binding_id, label in bindings:
            confidence = calculate_triage_confidence(binding_id, label, ir_conn)
            results.append(confidence)

        return results

    except Exception as e:
        logger.error(f"Error assessing triage confidence: {e}")
        return []
    finally:
        if ir_conn:
            ir_conn.close()
        if overlay_conn:
            overlay_conn.close()
