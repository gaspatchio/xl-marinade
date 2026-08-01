# ABOUTME: Actuarial variable taxonomy classifier
# ABOUTME: Classifies bindings as Assumption, Input, Calculation, or Result

import json
import logging
import re
import sqlite3
from dataclasses import dataclass, field
from enum import Enum

# Use TYPE_CHECKING to avoid runtime import issues if client not available
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    pass  # RAG removed (NEW 2)

from xl_marinade.docs.utils.ir_schema import detect_dependency_edges

logger = logging.getLogger(__name__)


class ActuarialType(str, Enum):
    """Actuarial variable classification types."""

    ASSUMPTION = "Assumption"
    POLICYHOLDER_DATA = "Policyholder Data"
    CALCULATION = "Calculation"
    RESULT = "Result"
    INDEX_LOOKUP = "Index Lookup"
    UNCLASSIFIED = "Unclassified"


@dataclass
class ClassificationResult:
    """Result of actuarial classification."""

    actuarial_type: ActuarialType
    reasoning: str
    confidence: float  # 0.0 to 1.0


@dataclass
class ActuarialEvidence:
    """Evidence gathered for classification."""

    label: str
    sheet: str
    has_formula: bool
    parent_count: int
    child_count: int
    is_sink: bool
    is_source: bool = False
    is_trivial_formula: bool = False
    sample_values: list[Any] = None
    rag_tags: list[str] = field(default_factory=list)

    @property
    def has_parents(self) -> bool:
        """Check if binding has parents (dependents)."""
        return self.parent_count > 0

    @property
    def has_children(self) -> bool:
        """Check if binding has children (dependencies)."""
        return self.child_count > 0


class ActuarialClassifier:
    """
    Classify bindings into actuarial types based on semantic context.

    Classification logic (per design doc):
    1. Assumption: Judgmental view of future uncertain experience or management behaviour
    2. Policyholder Data: Policy-level attributes that vary by insured life
    3. Calculation: Formula transforming inputs/assumptions into intermediate values
    4. Result: Value intended for external reporting or decision making

    Uses heuristics based on:
    - Label semantic meaning (keywords, patterns)
    - Formula presence
    - Dependency structure (parents/children)
    - Sheet name context
    """

    # Keyword patterns for Assumptions (future experience, judgmental views)
    ASSUMPTION_KEYWORDS = [
        # Rates and experience (strong indicators)
        r"\bmortality\b",
        r"\bmorbidity\b",
        r"\blapse\b",
        r"\bsurrender\b",
        r"\bexpense\b",
        r"\bsalary\b",
        r"\binflation\b",
        # Discount and yield
        r"\bdiscount\b",
        r"\byield\b",
        r"\binterest rate\b",
        # Bonus and crediting
        r"\bbonus\b",
        r"\bcrediting rate\b",
        r"\bprofit.?sharing\b",
        # Management actions
        r"\bmanagement action\b",
        r"\bstrategy\b",
        # Assumptions sheet
        r"\bassumptions?\b",
        # Rate-like patterns (strong assumption indicators)
        r"\brate\b",
        r"\bfactor\b",
        r"\bpercentage\b",
        r"\b%\b",
        # Actuarial assumption terms
        r"\bstress\b",
        r"\bscenario\b",
        r"\bbest.?estimate\b",
        r"\badverse\b",
        r"\boptimistic\b",
        r"\bpessimistic\b",
        # Lookup tables (Story 24 - actuarial judgment inputs)
        r"\bcso\b",
        r"\btable\b",
    ]

    # Keyword patterns for Policyholder Data (policy-level attributes)
    POLICYHOLDER_KEYWORDS = [
        # Policy attributes - core policyholder data
        r"\bpolicy\b",
        r"\bage\b",
        r"\bdate of birth\b",
        r"\bdob\b",
        r"\bsex\b",
        r"\bgender\b",
        r"\bpremium\b",
        r"\bsum assured\b",
        r"\bproduct code\b",
        r"\bduration\b",
    ]

    # Keyword patterns for Results (outputs, reporting metrics)
    RESULT_KEYWORDS = [
        # Actuarial metrics
        r"\bbel\b",
        r"\bbest estimate\b",
        r"\breserve\b",
        r"\bliability\b",
        r"\bscr\b",
        r"\bcapital\b",
        r"\bpvfp\b",
        r"\bembedded value\b",
        r"\bprofit\b",
        # Reporting
        r"\btotal\b",
        r"\bsummary\b",
        r"\baggregat\b",
        r"\boutput\b",
        r"\bresult\b",
    ]

    # Sheet name patterns
    ASSUMPTION_SHEETS = [
        r"\bassumptions?\b",
        r"\brates\b",
        r"\bparameters\b",
        r"\btables?\b",
        r"\binterest\b",
        r"\bexpenses?\b",
        # Added patterns for common actuarial sheets (Observation 17 fix)
        r"\bdiscount",
        r"\bcurve\b",
        r"\bscenario",
    ]
    POLICYHOLDER_SHEETS = [r"\bpolicyholder\b", r"\bpolicy\b", r"\bdata\b"]
    RESULT_SHEETS = [r"\bresults?\b", r"\boutputs?\b", r"\breport\b", r"\bsummary\b"]

    # Scoring Weights
    WEIGHT_LABEL_STRONG = 0.6  # Strong keyword match
    WEIGHT_SHEET_MEDIUM = 0.3  # Sheet context (standard)
    WEIGHT_SHEET_WEAK = 0.2  # Sheet context (weak)

    # Structure Weights
    WEIGHT_STRUCT_SINK = 0.5  # Is a sink (Result indicator)
    WEIGHT_STRUCT_NOT_SINK = 0.4  # Is not a sink (Calc indicator)
    WEIGHT_STRUCT_FORMULA = 0.4  # Has formula (Calc indicator)
    WEIGHT_STRUCT_NO_FORMULA = 0.1  # No formula (Input/Assumption indicator)
    WEIGHT_STRUCT_DEPENDENCIES = 0.2  # Has children/dependencies
    WEIGHT_RESULT_FORMULA = 0.3  # Result usually has formula

    # RAG Weights
    WEIGHT_RAG_CONFIRMATION = 0.5  # Strong signal if RAG confirms type

    # Value-based weights
    WEIGHT_VALUE_INPUT_LIKE = 0.4  # Values look like input data (integers, dates, unique IDs)
    WEIGHT_VALUE_ASSUMPTION_LIKE = 0.4  # Values look like rates (0-1)

    # Penalties
    PENALTY_FORMULA = -0.5  # Penalty for Input having formula

    # Minimum confidence threshold - below this, classify as Unclassified
    # Per design doc (actuarial_taxonomy.md Req 4.4): "IF scores are ambiguous
    # (e.g. low confidence), the system SHALL mark the variable as Unclassified"
    MINIMUM_CONFIDENCE_THRESHOLD = 0.2

    def __init__(
        self,
        overlay_conn: sqlite3.Connection,
        ir_db_path: str,
    ):
        """
        Initialize classifier.

        Args:
            overlay_conn: Connection to overlay database (with IR attached)
            ir_db_path: Path to Phase 1 IR database
        """
        self.overlay_conn = overlay_conn
        self.ir_db_path = ir_db_path

        # Attach IR if not already attached (don't assume a specific IR schema/table).
        # Note: Some unit tests pass a mocked connection; skip auto-attach in that case.
        if isinstance(self.overlay_conn, sqlite3.Connection):
            try:
                attached_names = {
                    row[1] for row in self.overlay_conn.execute("PRAGMA database_list")
                }
            except sqlite3.OperationalError:
                attached_names = set()

            if "ir" not in attached_names:
                from xl_marinade.core.labelling.overlay_database import attach_ir_to_overlay

                attach_ir_to_overlay(overlay_conn, ir_db_path)

        self._ir_cells_has_binding_id = False
        self._ir_has_agent_cells = False
        if isinstance(self.overlay_conn, sqlite3.Connection):
            try:
                columns = [
                    row[1] for row in self.overlay_conn.execute("PRAGMA ir.table_info(cells)")
                ]
                self._ir_cells_has_binding_id = "binding_id" in columns
            except sqlite3.OperationalError:
                self._ir_cells_has_binding_id = False

            try:
                self.overlay_conn.execute("SELECT 1 FROM ir.agent_cells LIMIT 1")
                self._ir_has_agent_cells = True
            except sqlite3.OperationalError:
                self._ir_has_agent_cells = False

            self._ir_dependency_edges = detect_dependency_edges(self.overlay_conn, schema="ir")
        else:
            self._ir_dependency_edges = None

        self._precomputed_has_formula_ids: set[str] | None = None
        self._precomputed_parent_counts: dict[str, int] | None = None
        self._precomputed_child_counts: dict[str, int] | None = None
        self._precomputed_sample_values: dict[str, list[Any]] | None = None
        self._precomputed_trivial_formula: dict[str, bool] | None = None

    def set_precomputed_maps(
        self,
        has_formula_ids: set[str] | None = None,
        parent_counts: dict[str, int] | None = None,
        child_counts: dict[str, int] | None = None,
        sample_values: dict[str, list[Any]] | None = None,
        trivial_formula: dict[str, bool] | None = None,
    ) -> None:
        self._precomputed_has_formula_ids = has_formula_ids
        self._precomputed_parent_counts = parent_counts
        self._precomputed_child_counts = child_counts
        self._precomputed_sample_values = sample_values
        self._precomputed_trivial_formula = trivial_formula

    def classify(self, binding_id: str, label: str, sheet: str) -> ClassificationResult:
        """
        Classify a binding into actuarial type.

        Args:
            binding_id: Binding ID to classify
            label: Semantic label for binding
            sheet: Sheet name where binding is located

        Returns:
            ClassificationResult with type, reasoning, and confidence
        """
        evidence = self._gather_evidence(binding_id, label, sheet)
        return self._classify_by_score(evidence)

    def _classify_by_score(self, evidence: ActuarialEvidence) -> ClassificationResult:
        """
        Classify based on scoring definitions.

        Args:
            evidence: Gathered evidence

        Returns:
            ClassificationResult
        """
        scores = self.score_definitions(evidence)

        # Sort by score descending
        sorted_scores = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        winner_type, winner_score = sorted_scores[0]

        # Cap score at 1.0
        winner_score = min(1.0, winner_score)

        # Per design doc (actuarial_taxonomy.md Req 4.4): If scores are ambiguous
        # (e.g. low confidence), return Unclassified rather than forcing a weak match.
        # This prevents empty cells from being classified as Assumption/Input
        # with minimal evidence.
        if winner_score < self.MINIMUM_CONFIDENCE_THRESHOLD:
            return ClassificationResult(
                actuarial_type=ActuarialType.UNCLASSIFIED,
                reasoning=(
                    f"Insufficient evidence for classification (score {winner_score:.2f} "
                    f"< threshold {self.MINIMUM_CONFIDENCE_THRESHOLD}). "
                    f"Best candidate was {winner_type.value}."
                ),
                confidence=winner_score,
            )

        # Build reasoning string
        reasoning = self._generate_reasoning(winner_type, winner_score, evidence)

        return ClassificationResult(
            actuarial_type=winner_type, reasoning=reasoning, confidence=winner_score
        )

    def score_definitions(self, evidence: ActuarialEvidence) -> dict[ActuarialType, float]:
        """
        Score all actuarial types against the evidence.

        Args:
            evidence: Gathered evidence

        Returns:
            Dict mapping ActuarialType to score (0.0 to 1.0)
        """
        scores = {}
        scores[ActuarialType.ASSUMPTION] = self._score_assumption(evidence)
        scores[ActuarialType.POLICYHOLDER_DATA] = self._score_policyholder_data(evidence)
        scores[ActuarialType.CALCULATION] = self._score_calculation(evidence)
        scores[ActuarialType.RESULT] = self._score_result(evidence)
        return scores

    def _gather_evidence(self, binding_id: str, label: str, sheet: str) -> ActuarialEvidence:
        """Gather all evidence for a binding."""
        has_formula = self._has_formula(binding_id)
        is_trivial = self._is_trivial_formula(binding_id) if has_formula else False

        # NOTE: Edge direction in binding_level_edges is:
        # Dependent -> Dependency (Consumer -> Provider)
        # parent_count (to=me) = bindings that depend on me (Consumers/Downstream)
        # child_count (from=me) = bindings I depend on (Providers/Upstream)

        consumer_count = self._get_parent_count(binding_id)
        provider_count = self._get_child_count(binding_id)

        # Source (Input): I depend on nothing (0 Providers)
        is_source = provider_count == 0

        # Sink (Result): Nothing depends on me (0 Consumers)
        is_sink = consumer_count == 0

        sample_values = self._get_sample_values(binding_id)

        # RAG removed (NEW 2): no external tags; deterministic evidence only.
        rag_tags: list = []

        return ActuarialEvidence(
            label=label,
            sheet=sheet,
            has_formula=has_formula,
            parent_count=consumer_count,
            child_count=provider_count,
            is_sink=is_sink,
            is_trivial_formula=is_trivial,
            sample_values=sample_values,
            is_source=is_source,
            rag_tags=rag_tags,
        )

    def _score_assumption(self, evidence: ActuarialEvidence) -> float:
        """Score for Assumption candidate."""
        score = 0.0

        # 0. RAG Confirmation (Strong signal)
        if "Assumption" in evidence.rag_tags:
            score += self.WEIGHT_RAG_CONFIRMATION

        # 1. Label keywords (Strongest signal)
        if self._matches_patterns(evidence.label, self.ASSUMPTION_KEYWORDS):
            score += self.WEIGHT_LABEL_STRONG

            # Bonus: If label contains "rate", "factor", or similar (very strong assumption signal)
            if re.search(
                r"\b(rate|factor|percentage|%|stress|scenario)\b", evidence.label, re.IGNORECASE
            ):
                score += 0.15  # Additional boost for rate-like patterns

        # 2. Sheet context (Medium signal)
        if self._matches_patterns(evidence.sheet, self.ASSUMPTION_SHEETS):
            score += self.WEIGHT_SHEET_MEDIUM

        # 3. Structure (Small signal)
        if evidence.is_source:
            score += self.WEIGHT_STRUCT_NO_FORMULA
        elif evidence.is_trivial_formula:
            # Trivial formula (link/constant) is acceptable for Assumption
            # Reduced penalty or even neutral
            score += self.WEIGHT_STRUCT_NO_FORMULA * 0.5

        # 4. Lookup table heuristic (Story 24 - Medium signal)
        # Non-formulaic bindings that are heavily referenced are likely lookup tables
        # (e.g., CSO Table, Mortality Rates) which are assumptions, not calculations
        if not evidence.has_formula and evidence.parent_count >= 5:
            score += 0.3  # Boost for lookup table pattern
            logger.debug(
                f"Lookup table pattern detected for '{evidence.label}': "
                f"no formula, {evidence.parent_count} consumers"
            )

        # 5. Value analysis
        if self._values_look_like_assumption(evidence.sample_values):
            score += self.WEIGHT_VALUE_ASSUMPTION_LIKE

        return score

    def _score_policyholder_data(self, evidence: ActuarialEvidence) -> float:
        """Score for Policyholder Data candidate."""
        score = 0.0

        # 0. RAG Confirmation
        if "Policyholder Data" in evidence.rag_tags:
            score += self.WEIGHT_RAG_CONFIRMATION

        # 1. Label keywords (Strongest signal)
        if self._matches_patterns(evidence.label, self.POLICYHOLDER_KEYWORDS):
            score += self.WEIGHT_LABEL_STRONG

            # Penalty: If label contains rate-like patterns, likely an assumption not policyholder data
            # E.g., "Premium Rate" is assumption, but "Premium" is policyholder data
            if re.search(
                r"\b(rate|factor|percentage|%|stress|scenario)\b", evidence.label, re.IGNORECASE
            ):
                score -= 0.25  # Reduce policyholder data score for rate-like patterns

        # 2. Sheet context (Medium signal)
        if self._matches_patterns(evidence.sheet, self.POLICYHOLDER_SHEETS):
            score += self.WEIGHT_SHEET_MEDIUM

        # 3. Structure (Small signal)
        if evidence.is_source:
            score += self.WEIGHT_STRUCT_NO_FORMULA

        # 4. Value analysis
        if self._values_look_like_input(evidence.sample_values):
            score += self.WEIGHT_VALUE_INPUT_LIKE

        # Penalty: Has formula (Strong negative for Policyholder Data)
        # BUT: If it is a source (root), ignore formula penalty (could be a link or hardcoded calc)
        if evidence.has_formula and not evidence.is_source:
            if evidence.is_trivial_formula:
                # Trivial formula (link/constant) - small penalty
                score += self.PENALTY_FORMULA * 0.2
            else:
                # Calculation formula - full penalty
                score += self.PENALTY_FORMULA

        return max(0.0, score)

    def _score_calculation(self, evidence: ActuarialEvidence) -> float:
        """Score for Calculation candidate."""
        # Rule: If NO formula, it CANNOT be a calculation
        # (unless it's a copy/paste value result, which is edge case)
        # Exception: Trivial formulas are allowed but score low
        if not evidence.has_formula:
            return 0.0

        # Rule: If it is a SOURCE (no parents), it is unlikely to be
        # a calculation (it doesn't calculate FROM anything)
        # Exception: RAND(), NOW(), or hardcoded tables with formulas
        if evidence.is_source:
            return 0.1  # Very low score for source calculations

        score = 0.0

        # 0. RAG Confirmation
        if "Calculation" in evidence.rag_tags:
            score += self.WEIGHT_RAG_CONFIRMATION

        # 1. Structure: Has Formula (Strongest signal)
        if evidence.has_formula:
            if evidence.is_trivial_formula:
                # Trivial formula - weak signal for Calculation
                score += self.WEIGHT_STRUCT_FORMULA * 0.3
            else:
                # Complex formula - strong signal
                score += self.WEIGHT_STRUCT_FORMULA

        # 2. Structure: Not a sink (Medium signal)
        if not evidence.is_sink:
            score += self.WEIGHT_STRUCT_NOT_SINK

        # 3. Structure: Has dependencies (Small signal)
        if evidence.has_children:
            score += self.WEIGHT_STRUCT_DEPENDENCIES

        return score

    def _score_result(self, evidence: ActuarialEvidence) -> float:
        """Score for Result candidate."""
        score = 0.0

        # 0. RAG Confirmation
        if "Result" in evidence.rag_tags or "Output" in evidence.rag_tags:
            score += self.WEIGHT_RAG_CONFIRMATION

        # 1. Label keywords (Strongest signal)
        if self._matches_patterns(evidence.label, self.RESULT_KEYWORDS):
            score += self.WEIGHT_LABEL_STRONG

        # 2. Sheet context (Medium signal)
        if self._matches_patterns(evidence.sheet, self.RESULT_SHEETS):
            score += self.WEIGHT_SHEET_WEAK

        # 3. Structure: Is Sink (Medium signal)
        if evidence.is_sink:
            score += self.WEIGHT_STRUCT_SINK

        # 4. Structure: Has formula (Results usually calculated)
        # Re-using formula weight but scaled down slightly or use constant?
        # Previous code used 0.3. Let's add a specific constant for Result Formula if needed,
        # or reuse WEIGHT_SHEET_MEDIUM (0.3) if appropriate, but better to be explicit.
        # Let's stick to the previous value 0.3.
        # I'll define a WEIGHT_RESULT_FORMULA = 0.3
        if evidence.has_formula:
            score += self.WEIGHT_RESULT_FORMULA

        return score

    def _generate_reasoning(
        self, winner: ActuarialType, score: float, evidence: ActuarialEvidence
    ) -> str:
        """Generate human-readable reasoning."""
        reasons = []

        # Add specific reasons based on winner type
        if winner == ActuarialType.ASSUMPTION:
            if "Assumption" in evidence.rag_tags:
                reasons.append("RAG confirmed as Assumption")
            if self._matches_patterns(evidence.label, self.ASSUMPTION_KEYWORDS):
                reasons.append("Label matches assumption keywords")
            if self._matches_patterns(evidence.sheet, self.ASSUMPTION_SHEETS):
                reasons.append("Located on assumption sheet")
            if self._values_look_like_assumption(evidence.sample_values):
                reasons.append("Values look like rates/assumptions")

        elif winner == ActuarialType.POLICYHOLDER_DATA:
            if "Policyholder Data" in evidence.rag_tags:
                reasons.append("RAG confirmed as Policyholder Data")
            if self._matches_patterns(evidence.label, self.POLICYHOLDER_KEYWORDS):
                reasons.append("Label matches policyholder keywords")
            if self._matches_patterns(evidence.sheet, self.POLICYHOLDER_SHEETS):
                reasons.append("Located on policyholder sheet")
            if self._values_look_like_input(evidence.sample_values):
                reasons.append("Values look like policyholder data")
            if evidence.has_formula:
                reasons.append("WARNING: Has formula but classified as Policyholder Data")

        elif winner == ActuarialType.CALCULATION:
            if "Calculation" in evidence.rag_tags:
                reasons.append("RAG confirmed as Calculation")
            if evidence.has_formula:
                reasons.append("Has formula")
            if not evidence.is_sink:
                reasons.append("Is intermediate calculation (not a sink)")

        elif winner == ActuarialType.RESULT:
            if "Result" in evidence.rag_tags or "Output" in evidence.rag_tags:
                reasons.append("RAG confirmed as Result")
            if self._matches_patterns(evidence.label, self.RESULT_KEYWORDS):
                reasons.append("Label matches result keywords")
            if evidence.is_sink:
                reasons.append("Is a calculation sink (no dependents)")

        # Fallback if no specific reasons found (e.g. low score default)
        if not reasons:
            reasons.append("Best fit based on available evidence")

        return f"Classified as {winner.value} ({score:.2f}): {', '.join(reasons)}"

    def _get_sample_values(self, binding_id: str) -> list[Any]:
        """Get a sample of evaluated values from the binding."""
        if self._precomputed_sample_values is not None:
            return self._precomputed_sample_values.get(binding_id, [])

        if self._ir_cells_has_binding_id:
            rows = self.overlay_conn.execute(
                """
                SELECT evaluated_value
                FROM ir.cells
                WHERE binding_id = ? AND evaluated_value IS NOT NULL
                LIMIT 5
            """,
                (binding_id,),
            ).fetchall()
            return [row[0] for row in rows]
        if self._ir_has_agent_cells:
            rows = self.overlay_conn.execute(
                """
                SELECT ac.value
                FROM ir.agent_cells ac
                JOIN ir.cell_to_binding ctb ON ac.cell_id = ctb.cell_id
                WHERE ctb.binding_id = ? AND ac.value IS NOT NULL
                LIMIT 5
            """,
                (binding_id,),
            ).fetchall()
            values: list[Any] = []
            for (raw_value,) in rows:
                if raw_value is None:
                    continue
                try:
                    values.append(json.loads(raw_value))
                except (TypeError, json.JSONDecodeError):
                    values.append(raw_value)
            return values
        return []

    def _values_look_like_assumption(self, values: list[Any]) -> bool:
        """Check if values look like assumptions (rates, factors)."""
        if not values:
            return False

        # Assumptions are often small decimals (rates) or factors
        # But NOT integers like ages (0-100 is ambiguous) or years (2020)

        match_count = 0
        for val in values:
            try:
                # Rate: 0.0 < x < 1.0 (strict)
                if isinstance(val, (int, float)) and 0.0 <= float(val) <= 1.0:
                    match_count += 1
            except (ValueError, TypeError):
                continue

        # If > 50% are rates, it's assumption-like
        return (match_count / len(values)) > 0.5

    def _values_look_like_input(self, values: list[Any]) -> bool:
        """Check if values look like policyholder input data.

        Policyholder data characteristics:
        - Short codes: 'M', 'F' (gender), 'Y', 'N' (flags)
        - Policy IDs: alphanumeric like 'POL123', '12345'
        - Integers in typical ranges (age 0-120, policy term 1-50)

        NOT policyholder data:
        - Scenario names: 'Yield_Up', 'Yield_Down', 'Base_Case'
        - Table references: '2015 VBT', 'SOA 2017'
        - Headers/labels: 'Acquisition Cost', 'Maintenance Expense'
        - Configuration strings: 'Discounting', 'Monthly'
        - Date strings used as control inputs: '2024-12-31'

        Updated in Observation 17 fix to prevent over-classification.
        """
        if not values:
            return False

        match_count = 0
        for val in values:
            if isinstance(val, str):
                # Only short strings (1-3 chars) are likely policyholder codes
                # Examples: 'M', 'F', 'Y', 'N', '01', 'ABC'
                if len(val) <= 3 and val.isalnum():
                    match_count += 1
                    continue

                # Strings with underscores are scenario/config names, NOT policyholder data
                # Examples: 'Yield_Up', 'Base_Case', 'Stress_1'
                if "_" in val:
                    continue

                # Multi-word strings are likely headers/labels, NOT policyholder data
                # Examples: 'Acquisition Cost', '2015 VBT', 'Best Estimate'
                if " " in val:
                    continue

                # Date-like strings are likely control inputs, NOT policyholder data
                # Examples: '2024-12-31', '2024/12/31'
                if "-" in val or "/" in val:
                    continue

                # Single words > 3 chars are likely config/scenario names
                # Examples: 'Discounting', 'Monthly', 'Yield'
                # Don't count these as policyholder data
                continue

            try:
                num_val = float(val)
                # Integers in policyholder-like ranges (age, term, count)
                # Age: 0-120, Policy term: 1-50, Sum assured: > 1000
                if num_val.is_integer() and 1 < num_val <= 120:
                    match_count += 1
                elif num_val > 1000:
                    # Large integers are likely sum assured, premium amounts
                    match_count += 1
            except (ValueError, TypeError):
                continue

        return (match_count / len(values)) > 0.5 if values else False

    def _has_formula(self, binding_id: str) -> bool:
        """Check if binding contains any formula cells."""
        if self._precomputed_has_formula_ids is not None:
            return binding_id in self._precomputed_has_formula_ids

        if self._ir_cells_has_binding_id:
            result = self.overlay_conn.execute(
                """
                SELECT COUNT(*) FROM ir.cells
                WHERE binding_id = ? AND formula_a1 IS NOT NULL AND formula_a1 != ''
            """,
                (binding_id,),
            ).fetchone()
        elif self._ir_has_agent_cells:
            result = self.overlay_conn.execute(
                """
                SELECT COUNT(*)
                FROM ir.agent_cells ac
                JOIN ir.cell_to_binding ctb ON ac.cell_id = ctb.cell_id
                WHERE ctb.binding_id = ? AND ac.formula IS NOT NULL AND ac.formula != ''
            """,
                (binding_id,),
            ).fetchone()
        else:
            result = None
        return result[0] > 0 if result else False

    def _is_trivial_formula(self, binding_id: str) -> bool:
        """
        Check if binding formulas are trivial (direct links or constants).

        Trivial means:
        - Starts with '='
        - Contains only references, sheet names, !, $, or numbers
        - No function calls (no '(')
        """
        if self._precomputed_trivial_formula is not None:
            return self._precomputed_trivial_formula.get(binding_id, False)

        if self._ir_cells_has_binding_id:
            rows = self.overlay_conn.execute(
                """
                SELECT formula_a1 FROM ir.cells
                WHERE binding_id = ? AND formula_a1 IS NOT NULL AND formula_a1 != ''
                LIMIT 10
            """,
                (binding_id,),
            ).fetchall()
            formulas = [row[0] for row in rows]
        elif self._ir_has_agent_cells:
            rows = self.overlay_conn.execute(
                """
                SELECT ac.formula
                FROM ir.agent_cells ac
                JOIN ir.cell_to_binding ctb ON ac.cell_id = ctb.cell_id
                WHERE ctb.binding_id = ? AND ac.formula IS NOT NULL AND ac.formula != ''
                LIMIT 10
            """,
                (binding_id,),
            ).fetchall()
            formulas = [row[0] for row in rows]
        else:
            formulas = []

        if not formulas:
            return False

        for formula in formulas:
            # Check for function calls (parentheses)
            if "(" in formula:
                return False
            # Allow simple math? Maybe not for now. strict link check.
            # Just check if it looks like =Sheet!A1 or =A1 or =123
            if not re.match(r"^=?[A-Za-z0-9\.\s!$:]+$", formula):
                return False

        return True

    def _get_parent_count(self, binding_id: str) -> int:
        """Get number of parent bindings (bindings that depend on this one)."""
        if self._precomputed_parent_counts is not None:
            return self._precomputed_parent_counts.get(binding_id, 0)
        if self._ir_dependency_edges is None:
            return 0

        # Check if composite binding
        composite_members = self.overlay_conn.execute(
            """
            SELECT ir_binding_id FROM composite_bindings
            WHERE composite_id = ?
        """,
            (binding_id,),
        ).fetchall()

        binding_ids = [row[0] for row in composite_members] if composite_members else [binding_id]

        # Count unique parents across all member bindings
        parent_ids: set[str] = set()
        for bid in binding_ids:
            edges = self._ir_dependency_edges
            parents = self.overlay_conn.execute(
                f"""
                SELECT DISTINCT {edges.from_col}
                FROM ir.{edges.table}
                WHERE {edges.to_col} = ?
                """,
                (bid,),
            ).fetchall()
            parent_ids.update(row[0] for row in parents)

        return len(parent_ids)

    def _get_child_count(self, binding_id: str) -> int:
        """Get number of child bindings (bindings this one depends on)."""
        if self._precomputed_child_counts is not None:
            return self._precomputed_child_counts.get(binding_id, 0)
        if self._ir_dependency_edges is None:
            return 0

        # Check if composite binding
        composite_members = self.overlay_conn.execute(
            """
            SELECT ir_binding_id FROM composite_bindings
            WHERE composite_id = ?
        """,
            (binding_id,),
        ).fetchall()

        binding_ids = [row[0] for row in composite_members] if composite_members else [binding_id]

        # Count unique children across all member bindings
        child_ids: set[str] = set()
        for bid in binding_ids:
            edges = self._ir_dependency_edges
            children = self.overlay_conn.execute(
                f"""
                SELECT DISTINCT {edges.to_col}
                FROM ir.{edges.table}
                WHERE {edges.from_col} = ?
                """,
                (bid,),
            ).fetchall()
            child_ids.update(row[0] for row in children)

        return len(child_ids)

    def _matches_patterns(self, text: str, patterns: list[str]) -> bool:
        """Check if text matches any of the regex patterns."""
        return any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns)


def _precompute_composite_sheets(overlay_conn: sqlite3.Connection) -> dict[str, str]:
    try:
        rows = overlay_conn.execute("""
            SELECT cb.composite_id, b.sheet
            FROM composite_bindings cb
            JOIN semantic_variables sv ON sv.binding_id = cb.composite_id
            LEFT JOIN ir.agent_bindings b ON cb.ir_binding_id = b.binding_id
            WHERE sv.label IS NOT NULL AND sv.is_active = 1
            ORDER BY cb.composite_id, cb.ordinal
        """).fetchall()
    except sqlite3.OperationalError:
        rows = overlay_conn.execute("""
            SELECT cb.composite_id, b.sheet
            FROM composite_bindings cb
            JOIN semantic_variables sv ON sv.binding_id = cb.composite_id
            LEFT JOIN ir.bindings b ON cb.ir_binding_id = b.binding_id
            WHERE sv.label IS NOT NULL AND sv.is_active = 1
            ORDER BY cb.composite_id, cb.ordinal
        """).fetchall()
    sheets: dict[str, str] = {}
    for composite_id, sheet in rows:
        if composite_id not in sheets and sheet is not None:
            sheets[composite_id] = sheet
    return sheets


def _precompute_has_formula_ids(
    overlay_conn: sqlite3.Connection,
    has_binding_id_column: bool,
    has_agent_cells_view: bool,
) -> set[str]:
    if has_binding_id_column:
        rows = overlay_conn.execute("""
            SELECT DISTINCT c.binding_id
            FROM ir.cells c
            JOIN semantic_variables sv ON sv.binding_id = c.binding_id
            WHERE sv.label IS NOT NULL AND sv.is_active = 1
              AND c.formula_a1 IS NOT NULL AND c.formula_a1 != ''
        """).fetchall()
        formula_ids = {row[0] for row in rows}
        composite_rows = overlay_conn.execute("""
            SELECT DISTINCT cb.composite_id
            FROM composite_bindings cb
            JOIN semantic_variables sv ON sv.binding_id = cb.composite_id
            JOIN ir.cells c ON c.binding_id = cb.ir_binding_id
            WHERE sv.label IS NOT NULL AND sv.is_active = 1
              AND c.formula_a1 IS NOT NULL AND c.formula_a1 != ''
        """).fetchall()
        formula_ids.update(row[0] for row in composite_rows)
        return formula_ids

    if has_agent_cells_view:
        rows = overlay_conn.execute("""
            SELECT DISTINCT ctb.binding_id
            FROM ir.agent_cells ac
            JOIN ir.cell_to_binding ctb ON ac.cell_id = ctb.cell_id
            JOIN semantic_variables sv ON sv.binding_id = ctb.binding_id
            WHERE sv.label IS NOT NULL AND sv.is_active = 1
              AND ac.formula IS NOT NULL AND ac.formula != ''
        """).fetchall()
        formula_ids = {row[0] for row in rows}
        composite_rows = overlay_conn.execute("""
            SELECT DISTINCT cb.composite_id
            FROM composite_bindings cb
            JOIN semantic_variables sv ON sv.binding_id = cb.composite_id
            JOIN ir.cell_to_binding ctb ON ctb.binding_id = cb.ir_binding_id
            JOIN ir.agent_cells ac ON ac.cell_id = ctb.cell_id
            WHERE sv.label IS NOT NULL AND sv.is_active = 1
              AND ac.formula IS NOT NULL AND ac.formula != ''
        """).fetchall()
        formula_ids.update(row[0] for row in composite_rows)
        return formula_ids

    return set()


def _precompute_dependency_counts(
    overlay_conn: sqlite3.Connection,
) -> tuple[dict[str, int], dict[str, int]]:
    edges = detect_dependency_edges(overlay_conn, schema="ir")
    if edges is None:
        return {}, {}

    parent_rows = overlay_conn.execute(
        f"""
        SELECT abd.{edges.to_col}, COUNT(DISTINCT abd.{edges.from_col})
        FROM ir.{edges.table} abd
        JOIN semantic_variables sv ON sv.binding_id = abd.{edges.to_col}
        WHERE sv.label IS NOT NULL AND sv.is_active = 1
        GROUP BY abd.{edges.to_col}
        """
    ).fetchall()
    parent_counts = {row[0]: row[1] for row in parent_rows}

    child_rows = overlay_conn.execute(
        f"""
        SELECT abd.{edges.from_col}, COUNT(DISTINCT abd.{edges.to_col})
        FROM ir.{edges.table} abd
        JOIN semantic_variables sv ON sv.binding_id = abd.{edges.from_col}
        WHERE sv.label IS NOT NULL AND sv.is_active = 1
        GROUP BY abd.{edges.from_col}
        """
    ).fetchall()
    child_counts = {row[0]: row[1] for row in child_rows}

    composite_parent_rows = overlay_conn.execute(
        f"""
        SELECT cb.composite_id, COUNT(DISTINCT abd.{edges.from_col})
        FROM composite_bindings cb
        JOIN semantic_variables sv ON sv.binding_id = cb.composite_id
        JOIN ir.{edges.table} abd ON abd.{edges.to_col} = cb.ir_binding_id
        WHERE sv.label IS NOT NULL AND sv.is_active = 1
        GROUP BY cb.composite_id
        """
    ).fetchall()
    for composite_id, count in composite_parent_rows:
        parent_counts[composite_id] = count

    composite_child_rows = overlay_conn.execute(
        f"""
        SELECT cb.composite_id, COUNT(DISTINCT abd.{edges.to_col})
        FROM composite_bindings cb
        JOIN semantic_variables sv ON sv.binding_id = cb.composite_id
        JOIN ir.{edges.table} abd ON abd.{edges.from_col} = cb.ir_binding_id
        WHERE sv.label IS NOT NULL AND sv.is_active = 1
        GROUP BY cb.composite_id
        """
    ).fetchall()
    for composite_id, count in composite_child_rows:
        child_counts[composite_id] = count

    return parent_counts, child_counts


def _precompute_sample_values(
    overlay_conn: sqlite3.Connection,
    has_binding_id_column: bool,
    has_agent_cells_view: bool,
) -> dict[str, list[Any]] | None:
    try:
        if has_binding_id_column:
            rows = overlay_conn.execute("""
                WITH ordered AS (
                    SELECT c.binding_id,
                           c.evaluated_value,
                           ROW_NUMBER() OVER (
                               PARTITION BY c.binding_id
                               ORDER BY c.cell_address_a1
                           ) AS rn
                    FROM ir.cells c
                    JOIN semantic_variables sv ON sv.binding_id = c.binding_id
                    WHERE sv.label IS NOT NULL AND sv.is_active = 1
                      AND c.evaluated_value IS NOT NULL
                )
                SELECT binding_id, evaluated_value
                FROM ordered
                WHERE rn <= 5
            """).fetchall()
            values_map: dict[str, list[Any]] = {}
            for binding_id, value in rows:
                values_map.setdefault(binding_id, []).append(value)
            return values_map

        if has_agent_cells_view:
            rows = overlay_conn.execute("""
                WITH ordered AS (
                    SELECT ctb.binding_id,
                           ac.value,
                           ROW_NUMBER() OVER (
                               PARTITION BY ctb.binding_id
                               ORDER BY ac.cell_id
                           ) AS rn
                    FROM ir.agent_cells ac
                    JOIN ir.cell_to_binding ctb ON ac.cell_id = ctb.cell_id
                    JOIN semantic_variables sv ON sv.binding_id = ctb.binding_id
                    WHERE sv.label IS NOT NULL AND sv.is_active = 1
                      AND ac.value IS NOT NULL
                )
                SELECT binding_id, value
                FROM ordered
                WHERE rn <= 5
            """).fetchall()
            values_map = {}
            for binding_id, raw_value in rows:
                if raw_value is None:
                    continue
                try:
                    parsed = json.loads(raw_value)
                except (TypeError, json.JSONDecodeError):
                    parsed = raw_value
                values_map.setdefault(binding_id, []).append(parsed)
            return values_map
    except sqlite3.OperationalError:
        return None

    return {}


def _precompute_trivial_formulas(
    overlay_conn: sqlite3.Connection,
    has_binding_id_column: bool,
    has_agent_cells_view: bool,
) -> dict[str, bool] | None:
    try:
        if has_binding_id_column:
            rows = overlay_conn.execute("""
                WITH ordered AS (
                    SELECT c.binding_id,
                           c.formula_a1,
                           ROW_NUMBER() OVER (
                               PARTITION BY c.binding_id
                               ORDER BY c.cell_address_a1
                           ) AS rn
                    FROM ir.cells c
                    JOIN semantic_variables sv ON sv.binding_id = c.binding_id
                    WHERE sv.label IS NOT NULL AND sv.is_active = 1
                      AND c.formula_a1 IS NOT NULL AND c.formula_a1 != ''
                )
                SELECT binding_id, formula_a1
                FROM ordered
                WHERE rn <= 10
            """).fetchall()
        elif has_agent_cells_view:
            rows = overlay_conn.execute("""
                WITH ordered AS (
                    SELECT ctb.binding_id,
                           ac.formula,
                           ROW_NUMBER() OVER (
                               PARTITION BY ctb.binding_id
                               ORDER BY ac.cell_address
                           ) AS rn
                    FROM ir.agent_cells ac
                    JOIN ir.cell_to_binding ctb ON ac.cell_id = ctb.cell_id
                    JOIN semantic_variables sv ON sv.binding_id = ctb.binding_id
                    WHERE sv.label IS NOT NULL AND sv.is_active = 1
                      AND ac.formula IS NOT NULL AND ac.formula != ''
                )
                SELECT binding_id, formula
                FROM ordered
                WHERE rn <= 10
            """).fetchall()
        else:
            return {}
    except sqlite3.OperationalError:
        return None

    formulas_by_binding: dict[str, list[str]] = {}
    for binding_id, formula in rows:
        if formula:
            formulas_by_binding.setdefault(binding_id, []).append(formula)

    trivial_map: dict[str, bool] = {}
    trivial_pattern = re.compile(r"^=?[A-Za-z0-9\\.\\s!$:]+$")
    for binding_id, formulas in formulas_by_binding.items():
        is_trivial = True
        for formula in formulas:
            if "(" in formula:
                is_trivial = False
                break
            if not trivial_pattern.match(formula):
                is_trivial = False
                break
        trivial_map[binding_id] = is_trivial if formulas else False

    return trivial_map


def classify_all_bindings(
    overlay_conn: sqlite3.Connection,
    ir_db_path: str,
) -> dict[str, ClassificationResult]:
    """
    Classify all labelled bindings in overlay.

    Args:
        overlay_conn: Connection to overlay database (with IR attached)
        ir_db_path: Path to Phase 1 IR database

    Returns:
        Dict mapping binding_id -> ClassificationResult
    """
    classifier = ActuarialClassifier(overlay_conn, ir_db_path)

    # Get all active, labelled bindings
    try:
        rows = overlay_conn.execute("""
            SELECT sv.binding_id, sv.label, b.sheet
            FROM semantic_variables sv
            LEFT JOIN ir.agent_bindings b ON sv.binding_id = b.binding_id
            WHERE sv.label IS NOT NULL AND sv.is_active = 1
        """).fetchall()
    except sqlite3.OperationalError:
        rows = overlay_conn.execute("""
            SELECT sv.binding_id, sv.label, b.sheet
            FROM semantic_variables sv
            LEFT JOIN ir.bindings b ON sv.binding_id = b.binding_id
            WHERE sv.label IS NOT NULL AND sv.is_active = 1
        """).fetchall()

    sheet_map = {binding_id: sheet for binding_id, _, sheet in rows if sheet is not None}
    if len(sheet_map) < len(rows):
        sheet_map.update(_precompute_composite_sheets(overlay_conn))

    has_formula_ids = _precompute_has_formula_ids(
        overlay_conn,
        classifier._ir_cells_has_binding_id,
        classifier._ir_has_agent_cells,
    )
    parent_counts, child_counts = _precompute_dependency_counts(overlay_conn)
    sample_values = _precompute_sample_values(
        overlay_conn,
        classifier._ir_cells_has_binding_id,
        classifier._ir_has_agent_cells,
    )
    trivial_formula = _precompute_trivial_formulas(
        overlay_conn,
        classifier._ir_cells_has_binding_id,
        classifier._ir_has_agent_cells,
    )
    classifier.set_precomputed_maps(
        has_formula_ids,
        parent_counts,
        child_counts,
        sample_values=sample_values,
        trivial_formula=trivial_formula,
    )

    results = {}
    for binding_id, label, sheet in rows:
        sheet = sheet_map.get(binding_id, sheet)
        if sheet is None:
            sheet = "Unknown"

        result = classifier.classify(binding_id, label, sheet)
        results[binding_id] = result

        logger.debug(
            f"Classified '{label}' ({binding_id[:8]}...) as {result.actuarial_type.value}: "
            f"{result.reasoning} (confidence: {result.confidence:.2f})"
        )

    return results
