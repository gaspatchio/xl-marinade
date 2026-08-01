# ABOUTME: Diagnostic collection and reporting infrastructure for semantic lookup resolution
# ABOUTME: Tracks unresolved bindings and conservative fallbacks per design doc §5

from typing import Any


class DiagnosticCollector:
    """
    Collects diagnostic events during semantic resolution.

    Per design doc §5, tracks:
    - Unresolved bindings (resolution failed)
    - Conservative fallbacks (resolution used broader range)
    - Resolution attempts and traces

    Attributes:
        events: List of diagnostic events
    """

    def __init__(self) -> None:
        """Initialize diagnostic collector."""
        self.events: list[dict[str, Any]] = []

    def record_unresolved(
        self,
        cell_address: str,
        formula: str,
        function: str,
        failure_reason: str,
        details: dict[str, Any] | None = None,
    ) -> None:
        """
        Record an unresolved binding event.

        Args:
            cell_address: Cell address (A1 notation)
            formula: Full formula text
            function: Function name (VLOOKUP, INDEX, etc.)
            failure_reason: Short code for failure (e.g., "col_index_resolution_failed")
            details: Additional diagnostic details

        Example:
            >>> collector.record_unresolved(
            ...     "Projection!A10",
            ...     "=VLOOKUP(A5, E6:G607, D1, FALSE)",
            ...     "VLOOKUP",
            ...     "col_index_resolution_failed",
            ...     {"failed_argument": "col_index_num", "failed_argument_text": "D1"}
            ... )
        """
        event = {
            "type": "unresolved",
            "cell_address": cell_address,
            "formula_full": formula,
            "function": function,
            "failure_reason": failure_reason,
            "details": details or {},
        }
        self.events.append(event)

    def record_conservative_fallback(
        self,
        cell_address: str,
        formula: str,
        function: str,
        syntactic_ref: str,
        fallback_ref: str,
        reason: str,
        impact: str | None = None,
    ) -> None:
        """
        Record a conservative fallback event.

        Args:
            cell_address: Cell address (A1 notation)
            formula: Full formula text
            function: Function name
            syntactic_ref: Original syntactic reference
            fallback_ref: Fallback reference used
            reason: Why fallback was needed
            impact: Description of impact (e.g., "Using all 3 columns instead of specific column")

        Example:
            >>> collector.record_conservative_fallback(
            ...     "Projection!A10",
            ...     "=VLOOKUP(A5, E6:G607, D1, FALSE)",
            ...     "VLOOKUP",
            ...     "E6:G607",
            ...     "E6:G607",
            ...     "col_index D1 contains IF formula - cannot resolve",
            ...     "Using all 3 columns instead of specific column"
            ... )
        """
        event = {
            "type": "conservative_fallback",
            "cell_address": cell_address,
            "formula": formula,
            "function": function,
            "syntactic_ref": syntactic_ref,
            "fallback_ref": fallback_ref,
            "reason": reason,
            "impact": impact or "",
        }
        self.events.append(event)

    def record_resolved(
        self,
        cell_address: str,
        formula: str,
        function: str,
        resolved_ref: str,
        resolution_source: str = "automatic",
    ) -> None:
        """
        Record a successful resolution event.

        Args:
            cell_address: Cell address (A1 notation)
            formula: Full formula text
            function: Function name
            resolved_ref: Resolved semantic reference
            resolution_source: "automatic" | "manual"

        Example:
            >>> collector.record_resolved(
            ...     "Projection!A10",
            ...     "=VLOOKUP(A5, E6:G607, 2, FALSE)",
            ...     "VLOOKUP",
            ...     "F6:F607",
            ...     "automatic"
            ... )
        """
        event = {
            "type": "resolved",
            "cell_address": cell_address,
            "formula": formula,
            "function": function,
            "resolved_ref": resolved_ref,
            "resolution_source": resolution_source,
        }
        self.events.append(event)

    def record_argument_resolution_failure(
        self,
        cell_address: str,
        argument_position: int,
        argument_type: str,
        failure_reason: str,
        attempted_strategies: list[str],
    ) -> None:
        """
        Record argument resolution failure for diagnostic analysis.

        Args:
            cell_address: Cell address containing the formula (A1 notation)
            argument_position: Position of argument (0-indexed)
            argument_type: Type of argument (e.g., "col_index", "row_index", "lookup_value")
            failure_reason: Human-readable reason for failure
            attempted_strategies: List of strategies that were attempted

        Example:
            >>> collector.record_argument_resolution_failure(
            ...     "Sheet1!A10",
            ...     2,
            ...     "col_index",
            ...     "Cell reference D1 contains complex formula",
            ...     ["literal", "cell_ref", "expression"]
            ... )
        """
        event = {
            "type": "argument_resolution_failure",
            "cell_address": cell_address,
            "argument_position": argument_position,
            "argument_type": argument_type,
            "failure_reason": failure_reason,
            "attempted_strategies": attempted_strategies,
        }
        self.events.append(event)

    def get_events_by_type(self, event_type: str) -> list[dict[str, Any]]:
        """
        Get all events of a specific type.

        Args:
            event_type: "unresolved" | "conservative_fallback" | "resolved"

        Returns:
            List of events matching the type
        """
        return [e for e in self.events if e.get("type") == event_type]

    def clear(self) -> None:
        """Clear all collected events."""
        self.events.clear()


def generate_unresolved_bindings_report(events: list[dict[str, Any]]) -> dict[str, Any]:
    """
    Generate unresolved bindings diagnostic report.

    Per design doc §5, produces JSON report with:
    - Summary statistics
    - Detailed unresolved cases with failure reasons
    - Resolution attempt traces (if available)

    Args:
        events: List of all diagnostic events

    Returns:
        Report dictionary (to be written to JSON by caller)

    Example output:
        {
          "summary": {
            "total_cells_with_lookups": 150,
            "resolved_automatically": 142,
            "resolved_manually": 5,
            "unresolved": 3,
            "resolution_rate": 0.98
          },
          "unresolved_details": [...]
        }
    """
    # Count events by type and resolution source
    resolved_events = [e for e in events if e.get("type") == "resolved"]
    unresolved_events = [e for e in events if e.get("type") == "unresolved"]

    resolved_auto = [e for e in resolved_events if e.get("resolution_source") == "automatic"]
    resolved_manual = [e for e in resolved_events if e.get("resolution_source") == "manual"]

    total_cells = len(resolved_events) + len(unresolved_events)
    resolution_rate = len(resolved_events) / total_cells if total_cells > 0 else 0.0

    # Build summary
    summary = {
        "total_cells_with_lookups": total_cells,
        "resolved_automatically": len(resolved_auto),
        "resolved_manually": len(resolved_manual),
        "unresolved": len(unresolved_events),
        "resolution_rate": round(resolution_rate, 3),
    }

    # Build unresolved details
    unresolved_details = []
    for event in unresolved_events:
        detail = {
            "cell_address": event.get("cell_address", ""),
            "formula_full": event.get("formula_full", ""),
            "function": event.get("function", ""),
            "failure_reason": event.get("failure_reason", ""),
        }

        # Add optional fields if present
        if "details" in event and event["details"]:
            detail.update(event["details"])

        unresolved_details.append(detail)

    return {"summary": summary, "unresolved_details": unresolved_details}


def generate_conservative_fallbacks_report(events: list[dict[str, Any]]) -> dict[str, Any]:
    """
    Generate conservative fallbacks diagnostic report.

    Per design doc §5, produces JSON report with:
    - Summary by function type
    - Detailed fallback cases with reasons and impact

    Args:
        events: List of all diagnostic events

    Returns:
        Report dictionary (to be written to JSON by caller)

    Example output:
        {
          "summary": {
            "total_fallbacks": 5,
            "by_function": {
              "VLOOKUP": 3,
              "INDEX": 2
            }
          },
          "details": [...]
        }
    """
    fallback_events = [e for e in events if e.get("type") == "conservative_fallback"]

    # Count by function type
    by_function: dict[str, int] = {}
    for event in fallback_events:
        func = event.get("function", "UNKNOWN")
        by_function[func] = by_function.get(func, 0) + 1

    # Build summary
    summary = {"total_fallbacks": len(fallback_events), "by_function": by_function}

    # Build details
    details = []
    for event in fallback_events:
        detail = {
            "cell_address": event.get("cell_address", ""),
            "formula": event.get("formula", ""),
            "function": event.get("function", ""),
            "syntactic_ref": event.get("syntactic_ref", ""),
            "fallback_ref": event.get("fallback_ref", ""),
            "reason": event.get("reason", ""),
            "impact": event.get("impact", ""),
        }
        details.append(detail)

    return {"summary": summary, "details": details}
