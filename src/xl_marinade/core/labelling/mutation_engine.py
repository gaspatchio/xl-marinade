# ABOUTME: Core mutation system - validation, replay, and logging
# ABOUTME: Handles set_label mutations for Sprint 1, extensible for future mutation types

import json
import sqlite3
from dataclasses import dataclass, field
from typing import Any

from loguru import logger

from . import structural_mutations
from .mutation_errors import (
    MutationConflictError,
    MutationSequenceError,
    MutationValidationError,
)

# Mutation ordering is carried by the monotonic mutation_id, so the persisted
# timestamp need not be a wall clock. Using a fixed value keeps the overlay's
# compared surface (mutation_log, metadata) byte-identical across runs of the
# same workbook (ADR-000 determinism; matches the IR's ADR-010 epoch constant).
DETERMINISTIC_TIMESTAMP = "1970-01-01T00:00:00Z"


@dataclass
class BindingOverlay:
    """Semantic data for one binding."""

    binding_id: str
    label: str | None = None
    label_source: int | None = None  # mutation_id that set label
    actuarial_class: str | None = None  # Assumption, Input, Calculation, or Result
    actuarial_class_reasoning: str | None = None  # Why this classification was chosen
    actuarial_class_confidence: float | None = None  # Confidence score 0-1
    is_active: bool = True
    is_composite: bool = False
    superseded_by: int | None = None  # mutation_id that superseded this binding
    composite_members: list[str] = field(default_factory=list)  # IR binding IDs if composite

    # Reconciliation classification (Story 23)
    reconciliation_required: bool | None = None
    reconciliation_rationale: str | None = None

    # Confidence scores (Story 14)
    label_confidence: float | None = None
    classification_confidence: float | None = None

    # Orphan indicator (Story 15)
    is_orphan: bool = False


@dataclass
class OverlayState:
    """In-memory state during replay."""

    bindings: dict[str, BindingOverlay] = field(default_factory=dict)
    mutations_applied: list[dict[str, Any]] = field(default_factory=list)


def validate_mutation_schema(mutation: dict[str, Any]) -> None:
    """Validate single mutation against schema.

    Args:
        mutation: Mutation dict to validate

    Raises:
        MutationValidationError: If schema invalid
    """
    required_fields = ["mutation_id", "timestamp", "action", "parameters"]
    for field_name in required_fields:
        if field_name not in mutation:
            raise MutationValidationError(
                mutation.get("mutation_id", 0), f"Missing required field: {field_name}"
            )

    # Validate types
    if not isinstance(mutation["mutation_id"], int):
        raise MutationValidationError(
            mutation["mutation_id"],
            f"mutation_id must be int, got {type(mutation['mutation_id']).__name__}",
        )

    # Sprint 1-4: Allowed actions (+ propagate_family_label)
    allowed_actions = [
        "set_label",
        "merge_bindings",
        "split_binding",
        "override_binding",
        "set_reconciliation_flag",
        "disable_binding",
        "set_orphan_status",
        "propagate_family_label",
    ]
    if mutation["action"] not in allowed_actions:
        raise MutationValidationError(
            mutation["mutation_id"],
            f"Action '{mutation['action']}' not supported. Allowed: {allowed_actions}",
        )

    # Validate parameters by action type
    params = mutation["parameters"]

    if mutation["action"] == "set_label":
        required_params = ["binding_id", "old", "new"]
        for param in required_params:
            if param not in params:
                raise MutationValidationError(
                    mutation["mutation_id"], f"set_label missing parameter: {param}"
                )

        if not isinstance(params["binding_id"], str):
            raise MutationValidationError(
                mutation["mutation_id"],
                f"binding_id must be str, got {type(params['binding_id']).__name__}",
            )

        if params["new"] is None or params["new"] == "":
            raise MutationValidationError(
                mutation["mutation_id"], "new label cannot be null or empty"
            )

        if "confidence" in params and params["confidence"] is not None:
            if not isinstance(params["confidence"], (int, float)):
                raise MutationValidationError(
                    mutation["mutation_id"],
                    f"confidence must be float, got {type(params['confidence']).__name__}",
                )

    elif mutation["action"] == "merge_bindings":
        required_params = ["source_binding_ids", "new_binding_id", "label"]
        for param in required_params:
            if param not in params:
                raise MutationValidationError(
                    mutation["mutation_id"], f"merge_bindings missing parameter: {param}"
                )

        source_ids = params["source_binding_ids"]
        if not isinstance(source_ids, list) or len(source_ids) < 2:
            raise MutationValidationError(
                mutation["mutation_id"],
                "merge_bindings requires source_binding_ids list with at least 2 IDs",
            )

    elif mutation["action"] == "split_binding":
        required_params = ["source_binding_id", "new_bindings"]
        for param in required_params:
            if param not in params:
                raise MutationValidationError(
                    mutation["mutation_id"], f"split_binding missing parameter: {param}"
                )

        if not isinstance(params["new_bindings"], list) or len(params["new_bindings"]) < 2:
            raise MutationValidationError(
                mutation["mutation_id"],
                "split_binding requires new_bindings list with at least 2 definitions",
            )

    elif mutation["action"] == "override_binding":
        required_params = ["binding_id", "old_label", "new_label"]
        for param in required_params:
            if param not in params and param != "new_label":
                raise MutationValidationError(
                    mutation["mutation_id"], f"override_binding missing parameter: {param}"
                )
        # new_label OR actuarial_class must be present
        if "new_label" not in params and "actuarial_class" not in params:
            raise MutationValidationError(
                mutation["mutation_id"],
                "override_binding must specify either 'new_label' or 'actuarial_class'",
            )

    elif mutation["action"] == "set_reconciliation_flag":
        required_params = ["binding_id", "reconciliation_required", "rationale"]
        for param in required_params:
            if param not in params:
                raise MutationValidationError(
                    mutation["mutation_id"], f"set_reconciliation_flag missing parameter: {param}"
                )

        if not isinstance(params["reconciliation_required"], bool):
            raise MutationValidationError(
                mutation["mutation_id"],
                f"reconciliation_required must be bool, got {type(params['reconciliation_required']).__name__}",
            )

    elif mutation["action"] == "disable_binding":
        required_params = ["binding_id", "reason"]
        for param in required_params:
            if param not in params:
                raise MutationValidationError(
                    mutation["mutation_id"], f"disable_binding missing parameter: {param}"
                )

    elif mutation["action"] == "set_orphan_status":
        required_params = ["binding_id", "is_orphan"]
        for param in required_params:
            if param not in params:
                raise MutationValidationError(
                    mutation["mutation_id"], f"set_orphan_status missing parameter: {param}"
                )

        if not isinstance(params["binding_id"], str):
            raise MutationValidationError(
                mutation["mutation_id"],
                f"binding_id must be str, got {type(params['binding_id']).__name__}",
            )

        if not isinstance(params["is_orphan"], bool):
            raise MutationValidationError(
                mutation["mutation_id"],
                f"is_orphan must be bool, got {type(params['is_orphan']).__name__}",
            )


def validate_mutation_sequence(mutations: list[dict[str, Any]]) -> None:
    """Validate mutation ID sequence is contiguous and starts at 1.

    Args:
        mutations: List of mutations to validate

    Raises:
        MutationSequenceError: If sequence invalid
    """
    if not mutations:
        return  # Empty is valid

    # Extract IDs
    ids = [m["mutation_id"] for m in mutations]

    # Check for duplicates
    if len(ids) != len(set(ids)):
        duplicates = [mutation_id for mutation_id in ids if ids.count(mutation_id) > 1]
        raise MutationSequenceError(f"Duplicate mutation IDs found: {duplicates}")

    # Check starts at 1
    if min(ids) != 1:
        raise MutationSequenceError(f"Mutations must start at ID 1, got {min(ids)}")

    # Check contiguous
    expected = list(range(1, len(mutations) + 1))
    if sorted(ids) != expected:
        missing = set(expected) - set(ids)
        raise MutationSequenceError(f"Mutation IDs not contiguous. Missing: {sorted(missing)}")


def validate_mutations(mutations: list[dict[str, Any]], ir_db_path: str) -> None:
    """Validate entire mutation list before replay (all-or-nothing).

    Args:
        mutations: List of mutations to validate
        ir_db_path: Path to Phase 1 IR for FK validation

    Raises:
        MutationValidationError: If any mutation invalid
        MutationSequenceError: If sequence invalid
    """
    # 1. Validate sequence
    validate_mutation_sequence(mutations)

    # 2. Validate each mutation schema
    for mutation in mutations:
        validate_mutation_schema(mutation)

    # 3. Load IR binding IDs for FK validation (fast or legacy schema)
    ir_conn = sqlite3.connect(f"file:{ir_db_path}?mode=ro", uri=True)
    cursor = ir_conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type IN ('table', 'view')")
    objects = {row[0] for row in cursor.fetchall()}
    if "agent_bindings" in objects:
        cursor.execute("SELECT binding_id FROM agent_bindings")
    else:
        cursor.execute("SELECT binding_id FROM bindings")
    binding_ids = {row[0] for row in cursor.fetchall()}
    ir_conn.close()

    # 3.5 Include bindings created by structural mutations
    created_binding_ids: set[str] = set()
    for mutation in mutations:
        if mutation.get("action") == "merge_bindings":
            new_id = mutation.get("parameters", {}).get("new_binding_id")
            if isinstance(new_id, str):
                created_binding_ids.add(new_id)
        elif mutation.get("action") == "split_binding":
            for new_binding in mutation.get("parameters", {}).get("new_bindings", []):
                new_id = new_binding.get("binding_id")
                if isinstance(new_id, str):
                    created_binding_ids.add(new_id)
    if created_binding_ids:
        binding_ids.update(created_binding_ids)

    # 4. Validate binding_id references exist
    for mutation in mutations:
        if mutation["action"] == "set_label":
            binding_id = mutation["parameters"]["binding_id"]
            if binding_id not in binding_ids:
                raise MutationValidationError(
                    mutation["mutation_id"],
                    f"binding_id '{binding_id}' does not exist in IR database. "
                    f"Available bindings: {len(binding_ids)} total.",
                )
        elif (
            mutation["action"] == "set_reconciliation_flag"
            or mutation["action"] == "disable_binding"
            or mutation["action"] == "set_orphan_status"
        ):
            binding_id = mutation["parameters"]["binding_id"]
            if binding_id not in binding_ids:
                raise MutationValidationError(
                    mutation["mutation_id"],
                    f"binding_id '{binding_id}' does not exist in IR database.",
                )


def handle_set_label(
    overlay: OverlayState, mutation: dict[str, Any], ir_db_path: str
) -> OverlayState:
    """Apply set_label mutation to overlay state.

    Args:
        overlay: Current overlay state
        mutation: Mutation to apply
        ir_db_path: Path to Phase 1 IR (for validation)

    Returns:
        Updated overlay state

    Raises:
        MutationConflictError: If old value doesn't match current
    """
    params = mutation["parameters"]
    binding_id = params["binding_id"]
    old_label = params["old"]
    new_label = params["new"]

    # Get or create binding overlay
    if binding_id not in overlay.bindings:
        overlay.bindings[binding_id] = BindingOverlay(binding_id=binding_id)

    binding = overlay.bindings[binding_id]

    # Validate old value matches current
    if binding.label != old_label:
        raise MutationConflictError(
            mutation["mutation_id"],
            f"Expected old label '{old_label}' but current is '{binding.label}'",
        )

    # Apply mutation
    binding.label = new_label
    binding.label_source = mutation["mutation_id"]

    # Set confidence if provided (Story 14)
    confidence = params.get("confidence")
    if confidence is not None:
        binding.label_confidence = float(confidence)

    # Record in history
    overlay.mutations_applied.append(mutation)

    return overlay


def handle_set_reconciliation_flag(
    overlay: OverlayState, mutation: dict[str, Any], ir_db_path: str
) -> OverlayState:
    """Apply set_reconciliation_flag mutation to overlay state.

    Args:
        overlay: Current overlay state
        mutation: Mutation to apply
        ir_db_path: Path to Phase 1 IR (for validation)

    Returns:
        Updated overlay state
    """
    params = mutation["parameters"]
    binding_id = params["binding_id"]
    required = params["reconciliation_required"]
    rationale = params["rationale"]

    # Get or create binding overlay
    if binding_id not in overlay.bindings:
        overlay.bindings[binding_id] = BindingOverlay(binding_id=binding_id)

    binding = overlay.bindings[binding_id]

    # Apply mutation
    binding.reconciliation_required = required
    binding.reconciliation_rationale = rationale

    # Record in history
    overlay.mutations_applied.append(mutation)

    return overlay


def handle_set_orphan_status(
    overlay: OverlayState, mutation: dict[str, Any], ir_db_path: str
) -> OverlayState:
    """Apply set_orphan_status mutation to overlay state.

    Args:
        overlay: Current overlay state
        mutation: Mutation to apply
        ir_db_path: Path to Phase 1 IR (for validation)

    Returns:
        Updated overlay state
    """
    params = mutation["parameters"]
    binding_id = params["binding_id"]
    is_orphan = params["is_orphan"]

    # Get or create binding overlay
    if binding_id not in overlay.bindings:
        overlay.bindings[binding_id] = BindingOverlay(binding_id=binding_id)

    binding = overlay.bindings[binding_id]

    # Apply mutation
    binding.is_orphan = is_orphan

    # Record in history
    overlay.mutations_applied.append(mutation)

    return overlay


def handle_propagate_family_label(
    overlay: OverlayState, mutation: dict[str, Any], ir_db_path: str
) -> OverlayState:
    """Apply propagate_family_label mutation to overlay state.

    Copies a label from a representative binding to a family sibling.

    Args:
        overlay: Current overlay state
        mutation: Mutation with parameters: sibling_binding_id, label
        ir_db_path: Path to Phase 1 IR

    Returns:
        Updated overlay state
    """
    params = mutation["parameters"]
    sibling_id = params["sibling_binding_id"]
    label = params["label"]

    if sibling_id not in overlay.bindings:
        overlay.bindings[sibling_id] = BindingOverlay(binding_id=sibling_id)

    sibling = overlay.bindings[sibling_id]
    sibling.label = label
    sibling.label_source = mutation["mutation_id"]

    overlay.mutations_applied.append(mutation)
    return overlay


def replay_mutations(
    ir_db_path: str, mutations_path: str, *, skip_conflicts: bool = False
) -> OverlayState:
    """Replay mutations to produce semantic overlay state.

    Args:
        ir_db_path: Path to Phase 1 ir.db (immutable)
        mutations_path: Path to mutations.json
        skip_conflicts: When True, a mutation that fails to apply
            (MutationConflictError / MutationValidationError) is logged and
            skipped rather than aborting the whole replay. Used by the LLM
            enrichment path, where a conflicting LLM-proposed mutation must not
            discard the entire enrichment. Defaults to False (strict,
            all-or-nothing) for deterministic replay.

    Returns:
        OverlayState with all applicable mutations applied

    Raises:
        MutationValidationError: If a mutation is invalid (unless skip_conflicts)
        MutationConflictError: If mutations conflict (unless skip_conflicts)
        MutationSequenceError: If sequence invalid
    """
    # 1. Load mutations
    with open(mutations_path) as f:
        mutations = json.load(f)

    # 2. Validate all mutations (all-or-nothing)
    validate_mutations(mutations, ir_db_path)

    # 3. Initialize empty overlay state
    overlay = OverlayState()

    # 4. Sort by mutation_id (ensure sequential)
    mutations.sort(key=lambda m: m["mutation_id"])

    # 5. Apply each mutation
    skipped = 0
    for mutation in mutations:
        action = mutation["action"]

        try:
            if action == "set_label":
                overlay = handle_set_label(overlay, mutation, ir_db_path)
            elif action == "merge_bindings":
                overlay = structural_mutations.handle_merge_bindings(overlay, mutation, ir_db_path)
            elif action == "split_binding":
                overlay = structural_mutations.handle_split_binding(overlay, mutation, ir_db_path)
            elif action == "override_binding":
                overlay = structural_mutations.handle_override_binding(
                    overlay, mutation, ir_db_path
                )
            elif action == "set_reconciliation_flag":
                overlay = handle_set_reconciliation_flag(overlay, mutation, ir_db_path)
            elif action == "disable_binding":
                overlay = structural_mutations.handle_disable_binding(overlay, mutation, ir_db_path)
            elif action == "set_orphan_status":
                overlay = handle_set_orphan_status(overlay, mutation, ir_db_path)
            elif action == "propagate_family_label":
                overlay = handle_propagate_family_label(overlay, mutation, ir_db_path)
            else:
                # Should never reach here (validated above)
                raise MutationValidationError(mutation["mutation_id"], f"Unknown action: {action}")
        except (MutationConflictError, MutationValidationError) as exc:
            # Strict replay (the deterministic path) is all-or-nothing. Enrichment
            # opts into skip_conflicts: one bad LLM-proposed mutation — e.g. two
            # merges that both consume the same binding, where the later one hits
            # an already-superseded source — must not discard the entire
            # enrichment. Handlers validate before mutating, so a raised mutation
            # left the overlay untouched; skipping it is safe.
            if not skip_conflicts:
                raise
            skipped += 1
            logger.warning(
                "replay_mutations: skipping mutation {} ({})",
                mutation.get("mutation_id"),
                exc,
            )

    if skipped:
        logger.info(
            "replay_mutations: skipped {} of {} mutations as conflicting "
            "(enrichment resilience); applied the rest",
            skipped,
            len(mutations),
        )

    return overlay


class MutationLogger:
    """Generate and log mutations during agent processing."""

    def __init__(self) -> None:
        """Initialize logger with empty mutation list."""
        self.mutations: list[dict[str, Any]] = []
        self.next_id = 1

    def append_mutation(
        self, action: str, parameters: dict[str, Any], metadata: dict[str, Any]
    ) -> int:
        """
        Append a mutation with full control over metadata.

        This is a low-level method for cases where you need to preserve
        all metadata fields (e.g., sprint, story, query) that the
        high-level methods don't support.

        Args:
            action: Mutation action type
            parameters: Mutation parameters
            metadata: Full metadata dict (reasoning, knowledge_source, sprint, etc.)

        Returns:
            mutation_id assigned to this mutation
        """
        mutation: dict[str, Any] = {
            "mutation_id": self.next_id,
            "timestamp": DETERMINISTIC_TIMESTAMP,
            "action": action,
            "parameters": parameters,
            "metadata": metadata,
        }

        self.mutations.append(mutation)
        mutation_id = self.next_id
        self.next_id += 1

        return mutation_id

    def set_label(
        self,
        binding_id: str,
        old_label: str | None,
        new_label: str,
        reasoning: str,
        knowledge_source: str,
        confidence: float | None = None,
    ) -> int:
        """Log a set_label mutation.

        Args:
            binding_id: Binding to label
            old_label: Previous label (None if first time)
            new_label: New label to set
            reasoning: Why this label was chosen
            knowledge_source: Where info came from (ir_candidates, rag, llm, etc.)
            confidence: Optional confidence score 0-1

        Returns:
            mutation_id assigned to this mutation
        """
        metadata: dict[str, Any] = {
            "reasoning": reasoning,
            "knowledge_source": knowledge_source,
            "sprint": 1,
        }

        if confidence is not None:
            metadata["confidence_initial"] = confidence

        mutation: dict[str, Any] = {
            "mutation_id": self.next_id,
            "timestamp": DETERMINISTIC_TIMESTAMP,
            "action": "set_label",
            "parameters": {"binding_id": binding_id, "old": old_label, "new": new_label},
            "metadata": metadata,
        }

        if confidence is not None:
            mutation["parameters"]["confidence"] = confidence

        self.mutations.append(mutation)
        mutation_id = self.next_id
        self.next_id += 1

        return mutation_id

    def propagate_family_label(
        self,
        formula_family_id: str,
        representative_binding_id: str,
        label: str,
        reasoning: str,
        family_sibling_ids: list[str],
    ) -> list[int]:
        """Log propagate_family_label mutations for all siblings.

        Args:
            formula_family_id: ID of the formula family
            representative_binding_id: Binding that was originally labelled
            label: Label to propagate
            reasoning: Why this propagation is valid
            family_sibling_ids: List of sibling binding_ids to receive the label

        Returns:
            List of mutation_ids created
        """
        mutation_ids = []
        for sibling_id in family_sibling_ids:
            mutation: dict[str, Any] = {
                "mutation_id": self.next_id,
                "timestamp": DETERMINISTIC_TIMESTAMP,
                "action": "propagate_family_label",
                "parameters": {
                    "formula_family_id": formula_family_id,
                    "representative_binding_id": representative_binding_id,
                    "sibling_binding_id": sibling_id,
                    "label": label,
                },
                "metadata": {"reasoning": reasoning, "sprint": 5},
            }
            self.mutations.append(mutation)
            mutation_ids.append(self.next_id)
            self.next_id += 1
        return mutation_ids

    def merge_bindings(
        self,
        source_binding_ids: list[str],
        new_binding_id: str,
        label: str | None,
        reasoning: str,
        confidence: float | None = None,
        actuarial_class: str | None = None,
    ) -> int:
        """Log a merge_bindings mutation."""
        metadata: dict[str, Any] = {"reasoning": reasoning, "sprint": 2}

        if confidence is not None:
            metadata["confidence_initial"] = confidence

        if actuarial_class is not None:
            metadata["actuarial_class"] = actuarial_class

        mutation: dict[str, Any] = {
            "mutation_id": self.next_id,
            "timestamp": DETERMINISTIC_TIMESTAMP,
            "action": "merge_bindings",
            "parameters": {
                "source_binding_ids": source_binding_ids,
                "new_binding_id": new_binding_id,
                "label": label,
            },
            "metadata": metadata,
        }

        self.mutations.append(mutation)
        mutation_id = self.next_id
        self.next_id += 1

        return mutation_id

    def override_binding(
        self,
        binding_id: str,
        old_label: str | None,
        new_label: str | None = None,
        actuarial_class: str | None = None,
        reasoning: str = "",
        classification_confidence: float | None = None,
    ) -> int:
        """Log an override_binding mutation.

        Args:
            binding_id: Binding ID to override
            old_label: Expected current label (for validation)
            new_label: New label to set (optional)
            actuarial_class: New actuarial classification (optional)
            reasoning: Human-readable reasoning
            classification_confidence: Confidence score for classification (0-1, optional)

        Returns:
            Mutation ID
        """
        metadata: dict[str, Any] = {"reasoning": reasoning, "sprint": 3}

        mutation: dict[str, Any] = {
            "mutation_id": self.next_id,
            "timestamp": DETERMINISTIC_TIMESTAMP,
            "action": "override_binding",
            "parameters": {"binding_id": binding_id, "old_label": old_label},
            "metadata": metadata,
        }

        if new_label is not None:
            mutation["parameters"]["new_label"] = new_label
        if actuarial_class is not None:
            mutation["parameters"]["actuarial_class"] = actuarial_class
        if classification_confidence is not None:
            mutation["parameters"]["classification_confidence"] = classification_confidence

        self.mutations.append(mutation)
        mutation_id = self.next_id
        self.next_id += 1

        return mutation_id

    def set_reconciliation_flag(
        self, binding_id: str, reconciliation_required: bool, rationale: str, source: str
    ) -> int:
        """Log a set_reconciliation_flag mutation.

        Args:
            binding_id: Binding ID
            reconciliation_required: True if reconciliation required
            rationale: Explanation
            source: 'heuristic' or 'override'

        Returns:
            mutation_id
        """
        metadata: dict[str, Any] = {"source": source, "sprint": 3}

        mutation: dict[str, Any] = {
            "mutation_id": self.next_id,
            "timestamp": DETERMINISTIC_TIMESTAMP,
            "action": "set_reconciliation_flag",
            "parameters": {
                "binding_id": binding_id,
                "reconciliation_required": reconciliation_required,
                "rationale": rationale,
            },
            "metadata": metadata,
        }

        self.mutations.append(mutation)
        mutation_id = self.next_id
        self.next_id += 1

        return mutation_id

    def split_binding(
        self,
        source_binding_id: str,
        new_bindings: list[dict[str, Any]],
        reasoning: str,
        knowledge_source: str = "deterministic",
    ) -> int:
        """Log a split_binding mutation.

        Args:
            source_binding_id: Source binding ID to split
            new_bindings: List of new binding definitions
            reasoning: Why this split was chosen
            knowledge_source: Where info came from

        Returns:
            mutation_id
        """
        metadata: dict[str, Any] = {
            "reasoning": reasoning,
            "knowledge_source": knowledge_source,
            "sprint": 2,
        }

        mutation: dict[str, Any] = {
            "mutation_id": self.next_id,
            "timestamp": DETERMINISTIC_TIMESTAMP,
            "action": "split_binding",
            "parameters": {"source_binding_id": source_binding_id, "new_bindings": new_bindings},
            "metadata": metadata,
        }

        self.mutations.append(mutation)
        mutation_id = self.next_id
        self.next_id += 1

        return mutation_id

    def disable_binding(self, binding_id: str, reason: str) -> int:
        """Log a disable_binding mutation (mark binding as garbage/inactive).

        Args:
            binding_id: Binding ID to disable
            reason: Why this binding should be disabled

        Returns:
            mutation_id
        """
        metadata: dict[str, Any] = {"reasoning": reason, "sprint": 3}

        mutation: dict[str, Any] = {
            "mutation_id": self.next_id,
            "timestamp": DETERMINISTIC_TIMESTAMP,
            "action": "disable_binding",
            "parameters": {"binding_id": binding_id, "reason": reason},
            "metadata": metadata,
        }

        self.mutations.append(mutation)
        mutation_id = self.next_id
        self.next_id += 1

        return mutation_id

    def set_orphan_status(self, binding_id: str, is_orphan: bool, reasoning: str) -> int:
        """Log a set_orphan_status mutation.

        Args:
            binding_id: Binding ID
            is_orphan: True if binding is orphan (not reachable from any root)
            reasoning: Explanation

        Returns:
            mutation_id
        """
        metadata: dict[str, Any] = {"reasoning": reasoning, "sprint": 4}

        mutation: dict[str, Any] = {
            "mutation_id": self.next_id,
            "timestamp": DETERMINISTIC_TIMESTAMP,
            "action": "set_orphan_status",
            "parameters": {"binding_id": binding_id, "is_orphan": is_orphan},
            "metadata": metadata,
        }

        self.mutations.append(mutation)
        mutation_id = self.next_id
        self.next_id += 1

        return mutation_id

    def load_from_file(self, path: str) -> None:
        """Load existing mutations from file and prepare for appending.

        Args:
            path: Path to existing mutations.json file

        Note:
            If file doesn't exist or is invalid JSON, starts fresh with next_id=1.
            If file exists, loads mutations and sets next_id to max(existing_ids) + 1.
        """
        try:
            with open(path) as f:
                existing_mutations = json.load(f)
                self.mutations = existing_mutations
                if existing_mutations:
                    max_id = max(m["mutation_id"] for m in existing_mutations)
                    self.next_id = max_id + 1
                else:
                    self.next_id = 1
        except (FileNotFoundError, json.JSONDecodeError):
            # Start fresh if file doesn't exist or is corrupted
            self.mutations = []
            self.next_id = 1

    def save(self, path: str) -> None:
        """Write mutations to JSON file.

        Args:
            path: Output file path for mutations.json
        """
        with open(path, "w") as f:
            json.dump(self.mutations, f, indent=2)
        # Ensure trailing newline
        with open(path, "a") as f:
            f.write("\n")
