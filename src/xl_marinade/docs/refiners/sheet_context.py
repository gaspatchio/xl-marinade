# ABOUTME: Sheet-level context refinement for post-processing labelling results
# ABOUTME: Analyzes labels on sheet together for consistency and sheet-level semantics

import logging
import sqlite3
from collections import Counter
from pathlib import Path

logger = logging.getLogger(__name__)


class SheetContextRefiner:
    """
    Post-pass refinement that reviews all labels on a sheet together.

    Applies sheet-level heuristics:
    1. Dominant prefix consistency: If 90% of variables have "Mortality..." prefix,
       refine outliers to match.
    2. Sheet name classification: If sheet is "Assumptions", all variables are
       classified as Assumptions.
    3. Consistency enforcement: Apply shared prefixes from sheet context.
    """

    def __init__(self, ir_db_path: str):
        """
        Initialize sheet context refiner.

        Args:
            ir_db_path: Path to Phase 1 IR database
        """
        if not Path(ir_db_path).exists():
            raise FileNotFoundError(f"IR database not found: {ir_db_path}")

        self.ir_db_path = ir_db_path

    def refine(
        self, overlay_db_path: str, output_mutations_path: str, consistency_threshold: float = 0.9
    ) -> int:
        """
        Refine labels by analyzing sheet-level context.

        Args:
            overlay_db_path: Path to semantic_overlay.db (will be read)
            output_mutations_path: Path to append new mutations
            consistency_threshold: Fraction of variables that must share
                pattern (default: 0.9)

        Returns:
            Number of labels refined
        """
        # Connect to overlay database. uri=True is required so the ATTACH below
        # (which uses a file: URI to open the IR database read-only) is parsed
        # as a URI rather than a literal filename — matches the pattern used in
        # two_pass_labeller.py and core/labelling/overlay_database.py.
        overlay_conn = sqlite3.connect(overlay_db_path, uri=True)
        overlay_conn.execute(f"ATTACH DATABASE 'file:{self.ir_db_path}?mode=ro' AS ir")

        # Group bindings by sheet
        sheet_groups = self._group_bindings_by_sheet(overlay_conn)

        refinements = []

        # Analyze each sheet
        for sheet_name, bindings in sheet_groups.items():
            logger.info(f"Analyzing sheet '{sheet_name}' with {len(bindings)} bindings")

            # Skip sheets with too few bindings
            if len(bindings) < 3:
                logger.debug(f"Sheet '{sheet_name}' has < 3 bindings, skipping")
                continue

            # Identify dominant prefix
            dominant_prefix = self._find_dominant_prefix(bindings, consistency_threshold)

            if dominant_prefix:
                logger.info(f"Sheet '{sheet_name}': dominant prefix '{dominant_prefix}' detected")

                # Find outliers
                for binding_id, current_label in bindings:
                    if current_label and not current_label.startswith(dominant_prefix):
                        # Outlier detected - refine to include prefix
                        new_label = f"{dominant_prefix} {current_label}"
                        refinements.append(
                            {
                                "binding_id": binding_id,
                                "old_label": current_label,
                                "new_label": new_label,
                                "reasoning": (
                                    f"Sheet-level consistency: '{dominant_prefix}' is dominant "
                                    f"prefix on sheet '{sheet_name}'"
                                ),
                            }
                        )

            # Apply sheet-level classification
            classification = self._classify_sheet(sheet_name)

            if classification:
                logger.info(f"Sheet '{sheet_name}': applying classification '{classification}'")

                # Note: For now, we log classification but don't force it
                # In future sprints with taxonomy (Story 7), we'll apply variable_class
                # For now, we just use it in reasoning
                for _, _ in bindings:
                    # We could force classification here, but for now just log
                    pass

        overlay_conn.close()

        logger.info(f"Sheet-level refinement: {len(refinements)} refinements identified")

        # TODO: Apply refinements via mutation system
        # For now, just return count
        return len(refinements)

    def _group_bindings_by_sheet(
        self, overlay_conn: sqlite3.Connection
    ) -> dict[str, list[tuple[str, str]]]:
        """
        Group bindings by sheet name.

        Args:
            overlay_conn: Overlay database connection (with IR attached)

        Returns:
            Dict mapping sheet_name -> [(binding_id, label), ...]
        """
        try:
            rows = overlay_conn.execute("""
                SELECT b.sheet, sv.binding_id, sv.label
                FROM semantic_variables sv
                JOIN ir.agent_bindings b ON sv.binding_id = b.binding_id
                WHERE sv.label IS NOT NULL
                  AND sv.is_active = 1
                ORDER BY b.sheet, b.address
            """).fetchall()
        except sqlite3.OperationalError:
            rows = overlay_conn.execute("""
                SELECT b.sheet, sv.binding_id, sv.label
                FROM semantic_variables sv
                JOIN ir.bindings b ON sv.binding_id = b.binding_id
                WHERE sv.label IS NOT NULL
                  AND sv.is_active = 1
                ORDER BY b.sheet, b.address_a1
            """).fetchall()

        groups: dict[str, list[tuple[str, str]]] = {}
        for sheet, binding_id, label in rows:
            if sheet not in groups:
                groups[sheet] = []
            groups[sheet].append((binding_id, label))

        return groups

    def _find_dominant_prefix(
        self, bindings: list[tuple[str, str]], threshold: float
    ) -> str | None:
        """
        Identify dominant prefix in labels (e.g., "Mortality").

        Args:
            bindings: List of (binding_id, label) tuples
            threshold: Fraction of labels that must share prefix (0.0-1.0)

        Returns:
            Dominant prefix string, or None if no dominant pattern found
        """
        # Extract first word from each label
        prefixes = []
        for _, label in bindings:
            if not label:
                continue

            # Split on whitespace and take first word
            words = label.split()
            if words:
                prefixes.append(words[0])

        if not prefixes:
            return None

        # Count prefix frequencies
        prefix_counts = Counter(prefixes)

        # Find most common prefix
        most_common_prefix, count = prefix_counts.most_common(1)[0]

        # Check if it meets threshold
        fraction = count / len(prefixes)

        if fraction >= threshold:
            return most_common_prefix

        return None

    def _classify_sheet(self, sheet_name: str) -> str | None:
        """
        Classify sheet based on name.

        Args:
            sheet_name: Sheet name from workbook

        Returns:
            Classification string ("Assumption", "Input", "Calculation",
            "Result") or None if no classification applies
        """
        sheet_lower = sheet_name.lower()

        # Classification keywords
        if any(keyword in sheet_lower for keyword in ["assumption", "assumptions"]):
            return "Assumption"

        if any(keyword in sheet_lower for keyword in ["input", "inputs"]):
            return "Input"

        if any(keyword in sheet_lower for keyword in ["calculation", "calc"]):
            return "Calculation"

        if any(keyword in sheet_lower for keyword in ["result", "results", "output", "outputs"]):
            return "Result"

        return None
