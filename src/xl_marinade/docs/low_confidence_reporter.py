# ABOUTME: Generate review reports for evaluating low-confidence labeling/classification decisions
# ABOUTME: Creates human-readable Markdown reports with RAG context and confidence score breakdowns

import logging
from dataclasses import dataclass

from xl_marinade.docs.confidence_scorer import ConfidenceScores

logger = logging.getLogger(__name__)


@dataclass
class LowConfidenceCase:
    """
    Complete context for a low-confidence binding requiring agent review.

    Grouped into logical sections:
    - Binding identification (binding_id, sheet, address, formula)
    - Current state (current_label, current_classification)
    - Label assessment (label_details)
    - Classification assessment (classification_details)
    - Supporting data (candidates, context, rag_matches)
    """

    # Binding identification
    binding_id: str
    sheet: str
    address: str
    formula: str | None

    # Current state
    current_label: str | None
    current_classification: str | None

    # Confidence assessments
    label_confidence: float
    classification_confidence: float
    label_components: dict[str, float]
    classification_components: dict[str, float]

    # Supporting data
    candidates: list[dict]  # Label candidates with scores
    parent_labels: list[str]  # Labels of parent bindings
    child_labels: list[str]  # Labels of child bindings
    rag_matches: list[dict]  # (unused; RAG removed, NEW 2)


def generate_review_report(
    cases: list[LowConfidenceCase], output_path: str, threshold: float = 0.7
) -> None:
    """
    Generate Markdown review report for a reviewer.

    Args:
        cases: List of low-confidence cases to review
        output_path: Path to save Markdown report
        threshold: Confidence threshold (for context in report)
    """
    logger.info(f"Generating review report with {len(cases)} low-confidence cases")

    # Sort cases by lowest confidence first
    cases_sorted = sorted(cases, key=lambda c: min(c.label_confidence, c.classification_confidence))

    with open(output_path, "w", encoding="utf-8", newline="\n") as f:
        f.write("# Low-Confidence Labeling & Classification Review\n\n")
        f.write(f"**Confidence Threshold:** {threshold:.2f}\n\n")
        f.write(f"**Total Cases Flagged:** {len(cases)}\n\n")

        # Summary statistics
        label_low = sum(1 for c in cases if c.label_confidence < threshold)
        class_low = sum(1 for c in cases if c.classification_confidence < threshold)
        both_low = sum(
            1
            for c in cases
            if c.label_confidence < threshold and c.classification_confidence < threshold
        )

        f.write("## Summary\n\n")
        f.write(f"- **Low Label Confidence:** {label_low} cases\n")
        f.write(f"- **Low Classification Confidence:** {class_low} cases\n")
        f.write(f"- **Both Low:** {both_low} cases\n\n")

        f.write("---\n\n")
        f.write("## Review Instructions\n\n")
        f.write("For each case below:\n\n")
        f.write(
            "1. **Review Context**: Examine binding details, formula, candidates, parent/child context\n"
        )
        f.write(
            "2. **Make Decision**: Decide if current label/classification is correct or needs improvement\n"
        )
        f.write("3. **Generate Mutation**: If change needed, create mutation with reasoning\n\n")
        f.write("**Example Mutation (label change):**\n")
        f.write("```python\n")
        f.write("from xl_marinade.core.labelling.mutation_engine import MutationLogger\n")
        f.write("logger = MutationLogger('mutations.json')\n")
        f.write("logger.set_label(\n")
        f.write("    binding_id='abc123',\n")
        f.write("    old='Ambiguous Label',\n")
        f.write("    new='Premium Income',\n")
        f.write(
            "    reasoning='RAG match confirms this is standard actuarial term for premium. Formula references PolicyCount*PremiumRate.'\n"
        )
        f.write(")\n")
        f.write("```\n\n")
        f.write("---\n\n")

        # Individual cases
        for idx, case in enumerate(cases_sorted, 1):
            f.write(f"## Case {idx}: {case.binding_id}\n\n")

            # Binding details
            f.write("### Binding Details\n\n")
            f.write("| Property | Value |\n")
            f.write("|----------|-------|\n")
            f.write(f"| **Sheet** | {case.sheet} |\n")
            f.write(f"| **Address** | {case.address} |\n")
            f.write(f"| **Formula** | `{case.formula or 'None'}` |\n")
            f.write(f"| **Current Label** | {case.current_label or 'UNSET'} |\n")
            f.write(
                f"| **Current Classification** | {case.current_classification or 'Unknown'} |\n\n"
            )

            # Confidence scores
            f.write("### Confidence Assessment\n\n")
            f.write("| Score Type | Confidence | Status |\n")
            f.write("|------------|------------|--------|\n")

            label_status = "⚠️ LOW" if case.label_confidence < threshold else "✅ OK"
            class_status = "⚠️ LOW" if case.classification_confidence < threshold else "✅ OK"

            f.write(f"| **Label** | {case.label_confidence:.3f} | {label_status} |\n")
            f.write(
                f"| **Classification** | {case.classification_confidence:.3f} | {class_status} |\n\n"
            )

            # Component breakdown
            if case.label_confidence < threshold:
                f.write("**Label Confidence Breakdown:**\n\n")
                for comp, score in case.label_components.items():
                    f.write(f"- `{comp}`: {score:.3f}\n")
                f.write("\n")

            if case.classification_confidence < threshold:
                f.write("**Classification Confidence Breakdown:**\n\n")
                for comp, score in case.classification_components.items():
                    f.write(f"- `{comp}`: {score:.3f}\n")
                f.write("\n")

            # Label candidates
            if case.candidates:
                f.write("### Label Candidates\n\n")
                f.write("| Candidate | Score | Type | Context |\n")
                f.write("|-----------|-------|------|----------|\n")
                for cand in sorted(case.candidates, key=lambda x: x.get("score", 0), reverse=True):
                    label = cand.get("label", "N/A")
                    score = cand.get("score", 0.0)
                    cand_type = cand.get("type", "unknown")
                    context = cand.get("context", "")
                    f.write(f"| {label} | {score:.3f} | {cand_type} | {context} |\n")
                f.write("\n")

            # Parent/child context
            if case.parent_labels or case.child_labels:
                f.write("### Dependency Context\n\n")
                if case.parent_labels:
                    f.write(f"**Parents (what uses this):** {', '.join(case.parent_labels)}\n\n")
                if case.child_labels:
                    f.write(f"**Children (what this uses):** {', '.join(case.child_labels)}\n\n")

            # RAG matches
            if case.rag_matches:
                f.write("### RAG Matches (Actuarial Knowledge)\n\n")
                f.write("| Term | Score | Definition |\n")
                f.write("|------|-------|------------|\n")
                for match in case.rag_matches[:5]:  # Top 5
                    term = match.get("term", "N/A")
                    score = match.get("score", 0.0)
                    definition = match.get("definition", "N/A")[:100]  # Truncate
                    f.write(f"| {term} | {score:.3f} | {definition}... |\n")
                f.write("\n")

            # Decision prompt
            f.write("### Agent Decision\n\n")
            f.write("**Your assessment:**\n\n")
            f.write("- [ ] Current label/classification is correct (no action needed)\n")
            f.write("- [ ] Improved label: `_______________`\n")
            f.write("- [ ] Improved classification: `_______________`\n")
            f.write("- [ ] Reasoning: _________________________________\n\n")

            f.write("---\n\n")

    logger.info(f"Review report saved to: {output_path}")


def filter_low_confidence_cases(
    all_confidence: list[ConfidenceScores], threshold: float = 0.7
) -> list[str]:
    """
    Filter binding IDs with confidence below threshold.

    Args:
        all_confidence: List of all confidence scores
        threshold: Confidence threshold

    Returns:
        List of binding IDs with low confidence
    """
    low_conf = []

    for conf in all_confidence:
        if conf.label_confidence < threshold or conf.classification_confidence < threshold:
            low_conf.append(conf.binding_id)

    logger.info(f"Found {len(low_conf)} bindings below confidence threshold {threshold}")

    return low_conf
