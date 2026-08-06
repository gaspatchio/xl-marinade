# ABOUTME: Post-hoc confidence assessment for labeling and classification decisions
# ABOUTME: Implements weighted scoring algorithm from Phase 2 design
"""
Confidence Scoring Module
=========================

This module implements a post-hoc confidence assessment system for label and
classification decisions made by the documentation agent. It uses a weighted
scoring algorithm to evaluate decision quality across multiple dimensions.

## Algorithm Philosophy

The confidence scoring system evaluates decisions through complementary signals:
- **Formula-semantic match**: Does the label fit the formula's purpose?
- **Graph context**: Does the label/classification fit parent/child relationships?
- **Domain knowledge**: Does RAG find supporting actuarial terminology?
- **Candidate clarity**: Was there a clear winner or ambiguous choices?

## Weight Configuration

Default weights are based on actuarial domain expertise:

**Label Confidence (LABEL_CONFIDENCE_WEIGHTS):**
- formula_label_match: 0.3 - Formula semantics are primary signal
- parent_context_fit: 0.2 - Parent context validates label choice
- child_context_fit: 0.2 - Child context validates label choice
- rag_match_strength: 0.2 - Domain knowledge confirms terminology
- candidate_ambiguity: 0.1 - Tie-breaking signal for close calls

**Classification Confidence (CLASSIFICATION_CONFIDENCE_WEIGHTS):**
- formula_pattern_clarity: 0.4 - Formula patterns strongly indicate role
- tree_position_fit: 0.3 - Graph position reveals computational role
- role_coherence: 0.3 - Role should fit parent/child roles

## Tuning Guidance

To adjust weights based on empirical testing:

1. Run confidence assessment on test set with known ground truth
2. Calculate precision/recall for low-confidence flagging at various thresholds
3. Identify which components have poor signal-to-noise ratio
4. Adjust weights to emphasize reliable signals, de-emphasize noisy ones
5. Re-run and verify improvement in precision/recall

Example: If parent_context_fit has high false-positive rate (flags good labels
as low-confidence), reduce its weight from 0.2 to 0.1 and redistribute to
more reliable signals.

## Integration

This module is called during Pass 4 of the two-pass labeling engine:
1. Pass 1-3: Generate labels and classifications deterministically
2. Pass 4: Assess confidence, persist scores, flag low-confidence cases
3. A reviewer works through the flagged cases and makes improvements

See: docs/phase2_documentation_agent/backlog/sprint4/story-14-agent-confidence-decisions.md
"""

import json
import logging
import sqlite3
from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    pass  # RAG removed (NEW 2)

from xl_marinade.docs.utils.ir_schema import detect_dependency_edges

logger = logging.getLogger(__name__)

# Configurable weight constants - can be tuned based on empirical testing
# These are DEFAULT values derived from actuarial domain expertise
LABEL_CONFIDENCE_WEIGHTS = {
    "formula_label_match": 0.3,  # Does label fit formula semantics?
    "parent_context_fit": 0.2,  # Does label fit parent labels?
    "child_context_fit": 0.2,  # Does label fit child labels?
    "rag_match_strength": 0.2,  # Does RAG find similar terms?
    "candidate_ambiguity": 0.1,  # Were candidates clear or ambiguous?
}

CLASSIFICATION_CONFIDENCE_WEIGHTS = {
    "formula_pattern_clarity": 0.4,  # Does formula clearly indicate role?
    "tree_position_fit": 0.3,  # Does role fit tree position?
    "role_coherence": 0.3,  # Does role fit parent/child roles?
}

# Verify weights sum to 1.0
assert abs(sum(LABEL_CONFIDENCE_WEIGHTS.values()) - 1.0) < 0.001, "Label weights must sum to 1.0"
assert abs(sum(CLASSIFICATION_CONFIDENCE_WEIGHTS.values()) - 1.0) < 0.001, (
    "Classification weights must sum to 1.0"
)


@dataclass
class ConfidenceScores:
    """Confidence scores for a single binding."""

    binding_id: str
    label_confidence: float
    classification_confidence: float
    label_components: dict[str, float]  # Sub-component scores for debugging
    classification_components: dict[str, float]


def formula_label_match(formula: str | None, label: str) -> float:
    """
    Check if label keywords appear in formula.

    Returns 0.0-1.0 based on how well the label matches formula semantics.

    Args:
        formula: Formula text (A1 notation)
        label: Variable label

    Returns:
        Match score 0.0-1.0
    """
    if not formula or not label:
        return 0.5  # Neutral

    # Extract keywords from label (split on spaces, underscores)
    label_keywords = set(label.lower().replace("_", " ").split())
    formula_lower = formula.lower()

    # Count how many label keywords appear in formula
    matches = sum(1 for keyword in label_keywords if keyword in formula_lower)

    if len(label_keywords) == 0:
        return 0.5

    # Score based on match ratio
    match_ratio = matches / len(label_keywords)

    if match_ratio >= 0.7:
        return 0.9  # Strong match
    elif match_ratio >= 0.4:
        return 0.7  # Moderate match
    elif match_ratio > 0:
        return 0.5  # Weak match
    else:
        return 0.3  # No match


def parent_context_fit(
    binding_id: str, label: str, ir_conn: sqlite3.Connection, overlay_conn: sqlite3.Connection
) -> float:
    """
    Check if label fits parent (dependent) labels.

    Returns 0.0-1.0 based on how well label fits parent context.

    Args:
        binding_id: Binding ID to check
        label: Proposed label
        ir_conn: Open connection to IR database (reused for efficiency)
        overlay_conn: Open connection to overlay database (reused for efficiency)

    Returns:
        Context fit score 0.0-1.0
    """
    try:
        edges = detect_dependency_edges(ir_conn)
        if edges is None:
            return 0.5

        # Get parent bindings (what uses this binding)
        # Edge semantics: from_binding_id depends on to_binding_id
        # So parents are bindings where this binding is the to_binding_id
        parents = ir_conn.execute(
            f"""
            SELECT DISTINCT {edges.from_col}
            FROM {edges.table}
            WHERE {edges.to_col} = ?
            """,
            (binding_id,),
        ).fetchall()

        if not parents:
            return 0.5  # No parents - neutral score

        # Get parent labels from semantic overlay
        parent_ids = [p[0] for p in parents]
        placeholders = ",".join("?" * len(parent_ids))
        parent_labels = overlay_conn.execute(
            f"""
            SELECT label
            FROM semantic_variables
            WHERE binding_id IN ({placeholders})
            AND label IS NOT NULL
        """,
            parent_ids,
        ).fetchall()

        if not parent_labels:
            return 0.5  # No parent labels - neutral score

        # Check label similarity with parent labels
        # Strategy: Look for keyword overlap between this label and parent labels
        label_keywords = set(label.lower().replace("_", " ").split())

        max_similarity = 0.0
        for parent_row in parent_labels:
            parent_label = parent_row[0]
            parent_keywords = set(parent_label.lower().replace("_", " ").split())

            # Calculate Jaccard similarity
            if len(label_keywords) > 0 and len(parent_keywords) > 0:
                intersection = len(label_keywords & parent_keywords)
                union = len(label_keywords | parent_keywords)
                similarity = intersection / union if union > 0 else 0.0
                max_similarity = max(max_similarity, similarity)

        # Score based on best parent similarity
        if max_similarity >= 0.5:
            return 0.9  # Strong contextual fit
        elif max_similarity >= 0.3:
            return 0.7  # Moderate fit
        elif max_similarity >= 0.1:
            return 0.6  # Weak fit
        else:
            return 0.4  # Poor fit

    except Exception as e:
        logger.warning(f"Error checking parent context: {e}")
        return 0.5


def child_context_fit(
    binding_id: str, label: str, ir_conn: sqlite3.Connection, overlay_conn: sqlite3.Connection
) -> float:
    """
    Check if label fits child (precedent) labels.

    Returns 0.0-1.0 based on how well label fits child context.

    Args:
        binding_id: Binding ID to check
        label: Proposed label
        ir_conn: Open connection to IR database (reused for efficiency)
        overlay_conn: Open connection to overlay database (reused for efficiency)

    Returns:
        Context fit score 0.0-1.0
    """
    try:
        edges = detect_dependency_edges(ir_conn)
        if edges is None:
            return 0.5

        # Get child bindings (what this binding uses)
        # Edge semantics: from_binding_id depends on to_binding_id
        # So children are bindings where this binding is the from_binding_id
        children = ir_conn.execute(
            f"""
            SELECT DISTINCT {edges.to_col}
            FROM {edges.table}
            WHERE {edges.from_col} = ?
            """,
            (binding_id,),
        ).fetchall()

        if not children:
            return 0.5  # No children - neutral score

        # Get child labels from semantic overlay
        child_ids = [c[0] for c in children]
        placeholders = ",".join("?" * len(child_ids))
        child_labels = overlay_conn.execute(
            f"""
            SELECT label
            FROM semantic_variables
            WHERE binding_id IN ({placeholders})
            AND label IS NOT NULL
        """,
            child_ids,
        ).fetchall()

        if not child_labels:
            return 0.5  # No child labels - neutral score

        # Check label similarity with child labels
        # Strategy: Look for keyword overlap between this label and child labels
        label_keywords = set(label.lower().replace("_", " ").split())

        max_similarity = 0.0
        for child_row in child_labels:
            child_label = child_row[0]
            child_keywords = set(child_label.lower().replace("_", " ").split())

            # Calculate Jaccard similarity
            if len(label_keywords) > 0 and len(child_keywords) > 0:
                intersection = len(label_keywords & child_keywords)
                union = len(label_keywords | child_keywords)
                similarity = intersection / union if union > 0 else 0.0
                max_similarity = max(max_similarity, similarity)

        # Score based on best child similarity
        if max_similarity >= 0.5:
            return 0.9  # Strong contextual fit
        elif max_similarity >= 0.3:
            return 0.7  # Moderate fit
        elif max_similarity >= 0.1:
            return 0.6  # Weak fit
        else:
            return 0.4  # Poor fit

    except Exception as e:
        logger.warning(f"Error checking child context: {e}")
        return 0.5


def rag_match_strength(rag_matches: list[dict]) -> float:
    """
    Extract RAG similarity score.

    Args:
        rag_matches: List of RAG match dictionaries with 'score' field

    Returns:
        RAG match score 0.0-1.0 (or 0.5 if no matches)
    """
    if not rag_matches:
        return 0.5  # No RAG - neutral score

    # Get best match score
    best_score = max(match.get("score", 0.0) for match in rag_matches)
    return best_score


def candidate_ambiguity(candidates: list[dict], selected_score: float) -> float:
    """
    Measure candidate score spread.

    Returns 0.0-1.0 based on how clear the winner was.
    High score = clear winner (low ambiguity)
    Low score = multiple similar candidates (high ambiguity)

    Args:
        candidates: List of candidate dictionaries with 'score' field
        selected_score: Score of the selected candidate

    Returns:
        Clarity score 0.0-1.0
    """
    if not candidates:
        return 0.0  # No candidates - very low confidence

    if len(candidates) == 1:
        return 1.0  # Only one candidate - no ambiguity

    # Sort candidates by score descending
    sorted_scores = sorted([c.get("score", 0.0) for c in candidates], reverse=True)

    if len(sorted_scores) < 2:
        return 1.0  # Only one score - no ambiguity

    # Calculate score difference between top 2 candidates
    score_diff = sorted_scores[0] - sorted_scores[1]

    # Normalize: larger diff = higher confidence
    # If diff >= 0.3, very clear winner (score 1.0)
    # If diff <= 0.05, very ambiguous (score 0.0)
    if score_diff >= 0.3:
        return 1.0
    elif score_diff <= 0.05:
        return 0.0
    else:
        # Linear interpolation between 0.05 and 0.3
        return (score_diff - 0.05) / (0.3 - 0.05)


def formula_pattern_clarity(formula: str | None, actuarial_class: str) -> float:
    """
    Check if formula clearly indicates actuarial role.

    Args:
        formula: Formula text (A1 notation)
        actuarial_class: Proposed classification

    Returns:
        Clarity score 0.0-1.0
    """
    if not formula:
        # No formula - should be Assumption or Policyholder Data
        if actuarial_class in ("Assumption", "Policyholder Data"):
            return 0.9  # High confidence
        else:
            return 0.3  # Mismatch: no formula but classified as calculation

    formula_lower = formula.lower()

    # Check for clear patterns
    if actuarial_class == "Calculation":
        # Calculations have formulas with transformations
        if any(func in formula_lower for func in ["sum", "average", "if", "vlookup", "index"]):
            return 0.9  # Strong signal
        else:
            return 0.6  # Has formula, probably calculation

    elif actuarial_class == "Result":
        # Results are often aggregations
        if any(func in formula_lower for func in ["sum", "sumproduct", "total"]):
            return 0.8  # Good signal
        else:
            return 0.5  # Uncertain

    elif actuarial_class in ("Assumption", "Policyholder Data"):
        # Should not have formula
        return 0.2  # Mismatch

    else:
        return 0.5  # Uncertain


def tree_position_fit(binding_id: str, actuarial_class: str, ir_conn: sqlite3.Connection) -> float:
    """
    Check if role fits graph position.

    Args:
        binding_id: Binding ID to check
        actuarial_class: Proposed classification
        ir_conn: Open connection to IR database (reused for efficiency)

    Returns:
        Position fit score 0.0-1.0
    """
    try:
        # Check if sink (no outgoing edges)
        edges = detect_dependency_edges(ir_conn)
        if edges is None:
            return 0.5

        outgoing = ir_conn.execute(
            f"SELECT COUNT(*) FROM {edges.table} WHERE {edges.from_col} = ?",
            (binding_id,),
        ).fetchone()[0]

        # Check if source (no incoming edges)
        incoming = ir_conn.execute(
            f"SELECT COUNT(*) FROM {edges.table} WHERE {edges.to_col} = ?",
            (binding_id,),
        ).fetchone()[0]

        is_sink = outgoing == 0
        is_source = incoming == 0

        if is_sink and actuarial_class == "Result":
            return 0.9  # Sinks are usually Results
        elif is_source and actuarial_class in ("Assumption", "Policyholder Data"):
            return 0.9  # Sources are usually Assumptions/Policyholder Data
        elif not is_sink and not is_source and actuarial_class == "Calculation":
            return 0.8  # Middle nodes are usually Calculations
        else:
            return 0.5  # Uncertain

    except Exception as e:
        logger.warning(f"Error checking tree position: {e}")
        return 0.5


def role_coherence(
    binding_id: str,
    actuarial_class: str,
    ir_conn: sqlite3.Connection,
    overlay_conn: sqlite3.Connection,
) -> float:
    """
    Check if role fits parent/child roles.

    Expected flow: Assumptions/Policyholder Data → Calculations → Results

    Args:
        binding_id: Binding ID to check
        actuarial_class: Proposed classification
        ir_conn: Open connection to IR database (reused for efficiency)
        overlay_conn: Open connection to overlay database (reused for efficiency)

    Returns:
        Coherence score 0.0-1.0
    """
    try:
        edges = detect_dependency_edges(ir_conn)
        if edges is None:
            return 0.5

        # Get parent bindings (what uses this binding) and their classifications
        # Edge semantics: from_binding_id depends on to_binding_id
        # So parents are bindings where this binding is the to_binding_id
        parents = ir_conn.execute(
            f"""
            SELECT DISTINCT {edges.from_col}
            FROM {edges.table}
            WHERE {edges.to_col} = ?
            """,
            (binding_id,),
        ).fetchall()

        # Get child bindings (what this binding uses) and their classifications
        # Edge semantics: from_binding_id depends on to_binding_id
        # So children are bindings where this binding is the from_binding_id
        children = ir_conn.execute(
            f"""
            SELECT DISTINCT {edges.to_col}
            FROM {edges.table}
            WHERE {edges.from_col} = ?
            """,
            (binding_id,),
        ).fetchall()

        if not parents and not children:
            return 0.5  # Isolated node - neutral

        # Get parent classifications
        parent_classes = []
        if parents:
            parent_ids = [p[0] for p in parents]
            placeholders = ",".join("?" * len(parent_ids))
            parent_rows = overlay_conn.execute(
                f"""
                SELECT actuarial_class
                FROM semantic_variables
                WHERE binding_id IN ({placeholders})
                AND actuarial_class IS NOT NULL
            """,
                parent_ids,
            ).fetchall()
            parent_classes = [r[0] for r in parent_rows]

        # Get child classifications
        child_classes = []
        if children:
            child_ids = [c[0] for c in children]
            placeholders = ",".join("?" * len(child_ids))
            child_rows = overlay_conn.execute(
                f"""
                SELECT actuarial_class
                FROM semantic_variables
                WHERE binding_id IN ({placeholders})
                AND actuarial_class IS NOT NULL
            """,
                child_ids,
            ).fetchall()
            child_classes = [r[0] for r in child_rows]

        # Check role coherence based on expected flow
        # Assumptions/Policyholder Data → Calculations → Results

        violations = 0
        checks = 0

        # Check parent coherence
        for parent_class in parent_classes:
            checks += 1
            if actuarial_class in ("Assumption", "Policyholder Data"):
                # Assumptions shouldn't be used by other assumptions
                if parent_class in ("Assumption", "Policyholder Data"):
                    violations += 1
                # It's OK for assumptions to be used by Calculations or Results
            elif actuarial_class == "Calculation":
                # Calculations can be used by other Calculations or Results (OK)
                # But shouldn't be used by Assumptions
                if parent_class in ("Assumption", "Policyholder Data"):
                    violations += 1
            elif actuarial_class == "Result":
                # Results shouldn't be used by Assumptions
                if parent_class in ("Assumption", "Policyholder Data"):
                    violations += 1
                # Results being used by Calculations is unusual but not wrong

        # Check child coherence
        for child_class in child_classes:
            checks += 1
            if actuarial_class in ("Assumption", "Policyholder Data"):
                # Assumptions shouldn't depend on Calculations or Results
                if child_class in ("Calculation", "Result"):
                    violations += 1
            elif actuarial_class == "Calculation":
                # Calculations can depend on Assumptions or other Calculations (OK)
                # But shouldn't depend on Results
                if child_class == "Result":
                    violations += 1
            elif actuarial_class == "Result":
                # Results can depend on anything (OK)
                pass

        if checks == 0:
            return 0.5  # No classifications to check - neutral

        # Calculate coherence score
        coherence_ratio = 1.0 - (violations / checks)

        if coherence_ratio >= 0.95:
            return 0.9  # Excellent coherence
        elif coherence_ratio >= 0.8:
            return 0.7  # Good coherence
        elif coherence_ratio >= 0.6:
            return 0.5  # Acceptable coherence
        else:
            return 0.3  # Poor coherence

    except Exception as e:
        logger.warning(f"Error checking role coherence: {e}")
        return 0.5


def calculate_label_confidence(
    binding_id: str,
    label: str,
    candidates: list[dict],
    selected_score: float,
    formula: str | None,
    rag_matches: list[dict],
    ir_conn: sqlite3.Connection,
    overlay_conn: sqlite3.Connection,
) -> tuple[float, dict[str, float]]:
    """
    Calculate overall label confidence using weighted algorithm.

    Args:
        binding_id: Binding ID
        label: Selected label
        candidates: List of candidate dictionaries
        selected_score: Score of selected candidate
        formula: Formula text
        rag_matches: RAG match results
        ir_conn: Open connection to IR database (reused for efficiency)
        overlay_conn: Open connection to overlay database (reused for efficiency)

    Returns:
        (overall_confidence, component_scores) tuple
    """
    components = {
        "formula_label_match": formula_label_match(formula, label),
        "parent_context_fit": parent_context_fit(binding_id, label, ir_conn, overlay_conn),
        "child_context_fit": child_context_fit(binding_id, label, ir_conn, overlay_conn),
        "rag_match_strength": rag_match_strength(rag_matches),
        "candidate_ambiguity": candidate_ambiguity(candidates, selected_score),
    }

    # Weighted average
    overall = sum(components[name] * LABEL_CONFIDENCE_WEIGHTS[name] for name in components)

    return overall, components


def calculate_classification_confidence(
    binding_id: str,
    actuarial_class: str,
    formula: str | None,
    ir_conn: sqlite3.Connection,
    overlay_conn: sqlite3.Connection,
) -> tuple[float, dict[str, float]]:
    """
    Calculate overall classification confidence using weighted algorithm.

    Args:
        binding_id: Binding ID
        actuarial_class: Proposed classification
        formula: Formula text
        ir_conn: Open connection to IR database (reused for efficiency)
        overlay_conn: Open connection to overlay database (reused for efficiency)

    Returns:
        (overall_confidence, component_scores) tuple
    """
    components = {
        "formula_pattern_clarity": formula_pattern_clarity(formula, actuarial_class),
        "tree_position_fit": tree_position_fit(binding_id, actuarial_class, ir_conn),
        "role_coherence": role_coherence(binding_id, actuarial_class, ir_conn, overlay_conn),
    }

    # Weighted average
    overall = sum(components[name] * CLASSIFICATION_CONFIDENCE_WEIGHTS[name] for name in components)

    return overall, components


def assess_all_confidence(
    ir_db_path: str,
    overlay_db_path: str,
) -> list[ConfidenceScores]:
    """
    Assess confidence for all bindings in the overlay.

    Args:
        ir_db_path: Path to IR database
        overlay_db_path: Path to semantic overlay database

    Returns:
        List of ConfidenceScores for each binding
    """
    results = []

    # Initialize connections to None for safe cleanup
    ir_conn = None
    overlay_conn = None

    try:
        # Connect to both databases
        ir_conn = sqlite3.connect(ir_db_path)
        overlay_conn = sqlite3.connect(overlay_db_path)

        # Get all bindings with labels and classifications
        bindings = overlay_conn.execute("""
            SELECT binding_id, label, actuarial_class
            FROM semantic_variables
            WHERE is_active = 1
        """).fetchall()

        logger.info(f"Assessing confidence for {len(bindings)} bindings")

        for binding_id, label, actuarial_class in bindings:
            # Get formula from IR (support both fast and legacy schemas)
            try:
                formula = ir_conn.execute(
                    """
                    SELECT ac.formula
                    FROM cell_to_binding ctb
                    JOIN agent_cells ac ON ctb.cell_id = ac.cell_id
                    WHERE ctb.binding_id = ? AND ac.formula IS NOT NULL AND ac.formula != ''
                    LIMIT 1
                """,
                    (binding_id,),
                ).fetchone()
            except sqlite3.OperationalError:
                formula = ir_conn.execute(
                    """
                    SELECT formula_a1
                    FROM cells
                    WHERE binding_id = ? AND formula_a1 IS NOT NULL AND formula_a1 != ''
                    LIMIT 1
                """,
                    (binding_id,),
                ).fetchone()

            formula_text = formula[0] if formula else None

            # Get candidates from IR bindings (support both fast and legacy schemas)
            try:
                candidates_row = ir_conn.execute(
                    """
                    SELECT spatial_candidates
                    FROM agent_bindings
                    WHERE binding_id = ?
                """,
                    (binding_id,),
                ).fetchone()
            except sqlite3.OperationalError:
                candidates_row = ir_conn.execute(
                    """
                    SELECT label_candidates_json
                    FROM bindings
                    WHERE binding_id = ?
                """,
                    (binding_id,),
                ).fetchone()

            candidates = []
            selected_score = 0.8  # Default if no candidates
            if candidates_row and candidates_row[0]:
                try:
                    candidates_data = json.loads(candidates_row[0])
                    # Handle case where JSON is "null" or not a dict
                    if isinstance(candidates_data, dict):
                        candidates = candidates_data.get("label_candidates", [])
                    elif isinstance(candidates_data, list):
                        # Backwards compatibility: if it's already a list
                        candidates = candidates_data
                    else:
                        candidates = []

                    # Find selected candidate score
                    if label and candidates:
                        for cand in candidates:
                            if isinstance(cand, dict) and cand.get("label") == label:
                                selected_score = cand.get("score", 0.8)
                                break
                except (json.JSONDecodeError, TypeError, AttributeError) as e:
                    logger.warning(f"Failed to parse candidates JSON for {binding_id}: {e}")

            # RAG removed (NEW 2): no external matches; deterministic scoring only.
            rag_matches: list = []

            # Calculate label confidence (pass connections for efficiency)
            label_conf, label_comps = calculate_label_confidence(
                binding_id,
                label or "",
                candidates,
                selected_score,
                formula_text,
                rag_matches,
                ir_conn,
                overlay_conn,
            )

            # Calculate classification confidence (pass connections for efficiency)
            class_conf, class_comps = calculate_classification_confidence(
                binding_id, actuarial_class or "Unknown", formula_text, ir_conn, overlay_conn
            )

            results.append(
                ConfidenceScores(
                    binding_id=binding_id,
                    label_confidence=label_conf,
                    classification_confidence=class_conf,
                    label_components=label_comps,
                    classification_components=class_comps,
                )
            )

        return results

    except Exception as e:
        logger.error(f"Error assessing confidence: {e}")
        return []
    finally:
        # Safe cleanup: only close if connections were created
        if ir_conn:
            ir_conn.close()
        if overlay_conn:
            overlay_conn.close()
