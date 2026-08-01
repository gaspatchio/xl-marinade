# ABOUTME: Structural mutation handlers for merge, split, and override operations
# ABOUTME: Implements Sprint 3 advanced binding mutations for semantic overlay

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from .mutation_engine import OverlayState

from .mutation_errors import MutationConflictError


def handle_merge_bindings(
    overlay: "OverlayState", mutation: dict[str, Any], ir_db_path: str
) -> "OverlayState":
    """Apply merge_bindings mutation to create composite binding.

    Creates a new virtual (composite) binding that aggregates multiple source bindings.
    Source bindings are marked as inactive and superseded, UNLESS cell_subset is used.

    Args:
        overlay: Current overlay state
        mutation: Mutation dict with parameters:
            - source_binding_ids: list[str] - IDs to merge (must be active)
            - new_binding_id: str - ID for new composite binding (UUID)
            - label: str - Label for composite binding
            - cell_subset: list[str] (optional) - If provided, only these cells from
              the FIRST source are included. The first source stays ACTIVE (for
              subsequent partial merges). Other sources are fully consumed.
        ir_db_path: Path to Phase 1 IR (for validation)

    Returns:
        Updated overlay state

    Raises:
        MutationConflictError: If source bindings not active
        MutationValidationError: If source bindings don't exist

    Note:
        cell_subset semantic (Story 37):
        - When cell_subset is specified, the FIRST source binding stays active
          to allow subsequent partial merges of other cells
        - Caller is responsible for issuing disable_binding when all partial
          merges are complete
        - Cell integrity: cells can be in multiple bindings (not preferred but allowed),
          but cells cannot be lost or added during merge operations
    """
    from .mutation_engine import BindingOverlay

    params = mutation["parameters"]
    source_ids = params["source_binding_ids"]
    new_id = params["new_binding_id"]
    label = params["label"]
    cell_subset = params.get("cell_subset")  # Story 37: optional partial merge

    # Validate BEFORE mutating: an existing source that is already inactive must
    # abort before we auto-create any OTHER source, so a rejected merge leaves
    # overlay.bindings untouched. replay_mutations(skip_conflicts=True) relies on
    # this atomicity — a skipped merge must not leave phantom entries behind.
    for source_id in source_ids:
        existing = overlay.bindings.get(source_id)
        if existing is not None and not existing.is_active:
            raise MutationConflictError(
                mutation["mutation_id"],
                f"Cannot merge inactive binding: {source_id} "
                f"(superseded by mutation {existing.superseded_by})",
            )

    # All sources valid — auto-create any missing IR binding (not yet touched by
    # mutations); a missing source is a fresh, active binding.
    for source_id in source_ids:
        if source_id not in overlay.bindings:
            overlay.bindings[source_id] = BindingOverlay(binding_id=source_id)

    # Create new composite binding
    composite = BindingOverlay(
        binding_id=new_id,
        label=label,
        label_source=mutation["mutation_id"],
        is_active=True,
        is_composite=True,
        composite_members=list(source_ids),
    )
    confidence = params.get("confidence")
    if confidence is None:
        confidence = mutation.get("metadata", {}).get("confidence_initial")
    if confidence is None:
        source_confidences = [
            overlay.bindings[source_id].label_confidence
            for source_id in source_ids
            if overlay.bindings[source_id].label_confidence is not None
        ]
        if source_confidences:
            confidence = max(source_confidences)
    if confidence is not None:
        composite.label_confidence = float(confidence)

    # Extract classification from mutation metadata (Story 34 + Cleanup Refiner fix)
    # Check both params.metadata (EntityGrouper) and top-level metadata (CleanupRefiner)
    metadata = params.get("metadata", {})
    if not metadata:
        metadata = mutation.get("metadata", {})

    if "actuarial_class" in metadata:
        composite.actuarial_class = metadata["actuarial_class"]
        composite.actuarial_class_confidence = metadata.get("actuarial_class_confidence", 0.90)
        entity_type = metadata.get("entity_type", "composite")
        composite.actuarial_class_reasoning = f"Classified at creation: {entity_type} entity"

    overlay.bindings[new_id] = composite

    # Mark source bindings as inactive/superseded
    # Story 37: If cell_subset is specified, FIRST source stays active for partial merge
    for idx, source_id in enumerate(source_ids):
        if cell_subset and idx == 0:
            # First source with cell_subset: stays active for subsequent partial merges
            # Caller must issue disable_binding when all partial merges are complete
            continue
        binding = overlay.bindings[source_id]
        binding.is_active = False
        binding.superseded_by = mutation["mutation_id"]

    # Record in history
    overlay.mutations_applied.append(mutation)

    return overlay


def handle_split_binding(
    overlay: "OverlayState", mutation: dict[str, Any], ir_db_path: str
) -> "OverlayState":
    """Apply split_binding mutation to partition a binding.

    Creates multiple new bindings from a single source binding.
    Source binding is marked as inactive and superseded.

    Args:
        overlay: Current overlay state
        mutation: Mutation dict with parameters:
            - source_binding_id: str - ID to split (must be active)
            - new_bindings: list[dict] - Definitions for new bindings
              Each dict: {binding_id: str, label: str, range: str}
        ir_db_path: Path to Phase 1 IR (for validation)

    Returns:
        Updated overlay state

    Raises:
        MutationConflictError: If source binding not active
    """
    from .mutation_engine import BindingOverlay

    params = mutation["parameters"]
    source_id = params["source_binding_id"]
    new_bindings = params["new_bindings"]

    # Validate source binding exists and is active
    if source_id not in overlay.bindings:
        overlay.bindings[source_id] = BindingOverlay(binding_id=source_id)

    source_binding = overlay.bindings[source_id]
    if not source_binding.is_active:
        raise MutationConflictError(
            mutation["mutation_id"],
            f"Cannot split inactive binding: {source_id} "
            f"(superseded by mutation {source_binding.superseded_by})",
        )

    # Create new bindings
    for new_def in new_bindings:
        new_binding = BindingOverlay(
            binding_id=new_def["binding_id"],
            label=new_def.get("label"),
            label_source=mutation["mutation_id"] if new_def.get("label") else None,
            is_active=True,
            is_composite=False,
        )
        overlay.bindings[new_def["binding_id"]] = new_binding

    # Mark source as inactive/superseded
    source_binding.is_active = False
    source_binding.superseded_by = mutation["mutation_id"]

    # Record in history
    overlay.mutations_applied.append(mutation)

    return overlay


def handle_override_binding(
    overlay: "OverlayState", mutation: dict[str, Any], ir_db_path: str
) -> "OverlayState":
    """Apply override_binding mutation (manual correction).

    Similar to set_label but with explicit "override" semantics for manual corrections.

    Args:
        overlay: Current overlay state
        mutation: Mutation dict with parameters:
            - binding_id: str - ID to override
            - old_label: str|None - Expected current label
            - new_label: str - New label to set
        ir_db_path: Path to Phase 1 IR (for validation)

    Returns:
        Updated overlay state

    Raises:
        MutationConflictError: If old_label doesn't match current
    """
    from .mutation_engine import BindingOverlay

    params = mutation["parameters"]
    binding_id = params["binding_id"]
    old_label = params["old_label"]
    new_label = params.get("new_label")
    actuarial_class = params.get("actuarial_class")
    classification_confidence = params.get("classification_confidence")

    # Get or create binding overlay
    if binding_id not in overlay.bindings:
        overlay.bindings[binding_id] = BindingOverlay(binding_id=binding_id)

    binding = overlay.bindings[binding_id]

    # Validate old value matches current (only if old_label provided)
    if old_label is not None and binding.label != old_label:
        raise MutationConflictError(
            mutation["mutation_id"],
            f"Expected old label '{old_label}' but current is '{binding.label}'",
        )

    # Apply override
    if new_label is not None:
        binding.label = new_label
        binding.label_source = mutation["mutation_id"]

    if actuarial_class is not None:
        binding.actuarial_class = actuarial_class
        # Use provided confidence, or default to 1.0 for manual overrides
        conf = classification_confidence if classification_confidence is not None else 1.0
        binding.actuarial_class_confidence = conf
        binding.classification_confidence = conf
        binding.actuarial_class_reasoning = mutation.get("metadata", {}).get(
            "reasoning", "Manual Override"
        )

    # Record in history
    overlay.mutations_applied.append(mutation)

    return overlay


def handle_disable_binding(
    overlay: "OverlayState", mutation: dict[str, Any], ir_db_path: str
) -> "OverlayState":
    """Apply disable_binding mutation (mark binding as garbage/inactive).

    Marks a binding as inactive without superseding it (used for garbage filtering).
    Unlike merge/split which supersedes bindings, this simply filters them out.

    Args:
        overlay: Current overlay state
        mutation: Mutation dict with parameters:
            - binding_id: str - ID to disable
            - reason: str - Why this binding should be disabled
        ir_db_path: Path to Phase 1 IR (for validation)

    Returns:
        Updated overlay state
    """
    from .mutation_engine import BindingOverlay

    params = mutation["parameters"]
    binding_id = params["binding_id"]

    # Get or create binding overlay
    if binding_id not in overlay.bindings:
        overlay.bindings[binding_id] = BindingOverlay(binding_id=binding_id)

    binding = overlay.bindings[binding_id]

    # Mark as inactive (idempotent - if already inactive, just log warning)
    if not binding.is_active:
        import logging

        logger = logging.getLogger(__name__)
        logger.warning(
            f"Binding {binding_id} already inactive "
            f"(superseded by mutation {binding.superseded_by})"
        )
    else:
        binding.is_active = False
        # Don't set superseded_by for disable (this is filtering, not structural change)

    # Record in history
    overlay.mutations_applied.append(mutation)

    return overlay


def generate_composite_id(source_binding_ids: list[str]) -> str:
    """Generate DETERMINISTIC unique ID for composite binding.

    IMPORTANT: This function now generates deterministic IDs based on source bindings
    to ensure mutation replayability (Story 25 requirement).

    Args:
        source_binding_ids: List of source binding IDs being merged

    Returns:
        Deterministic composite ID with 'composite_' prefix

    Example:
        >>> generate_composite_id(["bid1", "bid2"])
        'composite_abc123...'
        >>> generate_composite_id(["bid1", "bid2"])  # Same inputs -> same output
        'composite_abc123...'
    """
    import hashlib

    # Sort IDs for determinism
    sorted_ids = sorted(source_binding_ids)
    # Hash concatenated IDs
    id_string = "|".join(sorted_ids)
    hash_digest = hashlib.sha256(id_string.encode()).hexdigest()
    # Use first 16 chars of hash
    return f"composite_{hash_digest[:16]}"
