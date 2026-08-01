# ABOUTME: Verify invariants for VBA diff output.
# ABOUTME: Checks count conservation, match symmetry, and no rename+remove overlap.

"""
VBA Diff Verification — invariant checks on matcher output.

Verifies:
  1. Count conservation: procs_a = matched_a + removed + ambiguous_dropped_a
  2. Count conservation: procs_b = matched_b + added
  3. Match symmetry: if A matched B, B matched A (implicit in the data structure)
  4. No double-counting: no procedure appears in both renamed and removed/added
"""

from __future__ import annotations

from dataclasses import dataclass

from xl_marinade.core.ir_diff.vba_match import VBAMatchOutput


@dataclass
class VerifyResult:
    """Result of invariant checks."""

    passed: bool
    violations: list[str]


def verify_vba_match(
    output: VBAMatchOutput,
    procs_a_count: int,
    procs_b_count: int,
) -> VerifyResult:
    """
    Check invariants on a VBA match output.

    Args:
        output: The matcher output to verify
        procs_a_count: Total procedures in version A
        procs_b_count: Total procedures in version B

    Returns:
        VerifyResult with pass/fail and violation messages
    """
    violations: list[str] = []

    # Collect matched keys
    matched_a_keys = {m.key_a for m in output.matched}
    matched_b_keys = {m.key_b for m in output.matched}
    ambiguous_a = set(output.ambiguous_dropped)

    # 1. Count conservation for A
    a_accounted = len(matched_a_keys) + len(output.removed) + len(ambiguous_a)
    if a_accounted != procs_a_count:
        violations.append(
            f"A count mismatch: matched({len(matched_a_keys)}) + "
            f"removed({len(output.removed)}) + ambiguous({len(ambiguous_a)}) = "
            f"{a_accounted}, expected {procs_a_count}"
        )

    # 2. Count conservation for B
    b_accounted = len(matched_b_keys) + len(output.added)
    if b_accounted != procs_b_count:
        violations.append(
            f"B count mismatch: matched({len(matched_b_keys)}) + "
            f"added({len(output.added)}) = {b_accounted}, expected {procs_b_count}"
        )

    # 3. No duplicate matches (each key_a appears at most once)
    if len(matched_a_keys) != len(output.matched):
        violations.append(
            f"Duplicate A matches: {len(output.matched)} match results but "
            f"{len(matched_a_keys)} unique A keys"
        )

    # 4. No overlap between renamed and removed
    renamed_a_keys = {m.key_a for m in output.matched if "renamed" in m.change_type}
    overlap = renamed_a_keys & set(output.removed)
    if overlap:
        violations.append(f"Overlap between renamed and removed: {overlap}")

    # 5. No overlap between renamed and added
    renamed_b_keys = {m.key_b for m in output.matched if "renamed" in m.change_type}
    overlap_b = renamed_b_keys & set(output.added)
    if overlap_b:
        violations.append(f"Overlap between renamed and added: {overlap_b}")

    return VerifyResult(passed=len(violations) == 0, violations=violations)
