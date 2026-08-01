# ABOUTME: VBA-specific change type classifications for version comparison.
# ABOUTME: No REFERENCE_SHIFT analog — VBA string literals don't auto-update on row insertion.

"""
VBA Change Types for version comparison.

Change classifications for VBA procedures and modules. Intentionally does NOT
include a REFERENCE_SHIFT analog — VBA string literals (`Range("A1")`) are not
rewritten by Excel when rows are inserted. If addresses look different, it's
either a logic_change (human edit) or cosmetic (if only in comments/strings).
"""

from __future__ import annotations

from enum import Enum


class VBAChangeType(str, Enum):
    """Change types for VBA diff results."""

    # Module-level changes
    MODULE_ADDED = "module_added"
    MODULE_REMOVED = "module_removed"
    MODULE_RENAMED = "module_renamed"

    # Procedure-level changes
    PROCEDURE_ADDED = "procedure_added"
    PROCEDURE_REMOVED = "procedure_removed"
    PROCEDURE_RENAMED = "procedure_renamed"
    PROCEDURE_BODY_LOGIC_CHANGED = "procedure_body_logic_changed"
    PROCEDURE_BODY_COSMETIC_ONLY = "procedure_body_cosmetic_only"
    PROCEDURE_SIGNATURE_CHANGED = "procedure_signature_changed"
    PROCEDURE_RENAMED_AND_MODIFIED = "procedure_renamed_and_modified"

    # Declaration-level changes
    DECLARATION_CHANGED = "declaration_changed"
    DECLARATION_ADDED = "declaration_added"
    DECLARATION_REMOVED = "declaration_removed"
