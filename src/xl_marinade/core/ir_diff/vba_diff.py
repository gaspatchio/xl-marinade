# ABOUTME: Top-level VBA diff — orchestrates extraction, matching, and change classification.
# ABOUTME: Produces a VBADiffResult with module changes, procedure changes, and summary stats.

"""
VBA Diff — compares VBA content between two workbook extractions.

Takes two ir.db paths (or two VBAExtraction objects), runs the matcher,
classifies changes, and produces a structured result for the comparison
digest and Changes tab.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass, field
from pathlib import Path

from xl_marinade.core.ir_diff.vba_change_types import VBAChangeType
from xl_marinade.core.ir_diff.vba_match import (
    _build_desc,
    match_procedures,
)
from xl_marinade.core.ir_diff.vba_verify import verify_vba_match


@dataclass
class VBAProcedureChange:
    """A single procedure-level change between two versions."""

    module_a: str | None
    module_b: str | None
    name_a: str | None
    name_b: str | None
    kind: str  # sub, function, property_get, etc.
    change_type: str  # VBAChangeType value
    confidence: float
    match_pass: int  # 0 for add/remove, 1-4 for matches
    body_a: str | None = None  # for body diff display
    body_b: str | None = None
    signature_a: str | None = None
    signature_b: str | None = None


@dataclass
class VBADiffSummary:
    """High-level summary of VBA changes between two versions."""

    modules_added: int = 0
    modules_removed: int = 0
    procedures_added: int = 0
    procedures_removed: int = 0
    procedures_logic_changed: int = 0
    procedures_cosmetic_only: int = 0
    procedures_renamed: int = 0
    procedures_renamed_modified: int = 0
    procedures_unchanged: int = 0
    total_procedures_a: int = 0
    total_procedures_b: int = 0


@dataclass
class VBADiffResult:
    """Complete VBA diff result between two extractions."""

    summary: VBADiffSummary = field(default_factory=VBADiffSummary)
    procedure_changes: list[VBAProcedureChange] = field(default_factory=list)
    verify_passed: bool = True
    verify_violations: list[str] = field(default_factory=list)


def diff_vba_from_dbs(
    ir_db_a: str | Path,
    ir_db_b: str | Path,
) -> VBADiffResult:
    """
    Compare VBA content between two ir.db files.

    Args:
        ir_db_a: Path to version A's ir.db
        ir_db_b: Path to version B's ir.db

    Returns:
        VBADiffResult with summary, procedure-level changes, and verify status
    """
    result = VBADiffResult()

    # Load procedures from both databases
    procs_a = _load_procedures_from_db(ir_db_a)
    procs_b = _load_procedures_from_db(ir_db_b)
    mods_a = _load_module_names(ir_db_a)
    mods_b = _load_module_names(ir_db_b)

    result.summary.total_procedures_a = len(procs_a)
    result.summary.total_procedures_b = len(procs_b)

    # Module-level changes
    result.summary.modules_added = len(mods_b - mods_a)
    result.summary.modules_removed = len(mods_a - mods_b)

    if not procs_a and not procs_b:
        return result  # Both VBA-free

    # Build ProcedureDesc maps for the matcher
    descs_a = {
        k: _build_desc(k, v["module"], v["name"], v["kind"], v["body"]) for k, v in procs_a.items()
    }
    descs_b = {
        k: _build_desc(k, v["module"], v["name"], v["kind"], v["body"]) for k, v in procs_b.items()
    }

    # Run matcher
    match_output = match_procedures(descs_a, descs_b)

    # Verify invariants
    verify = verify_vba_match(match_output, len(procs_a), len(procs_b))
    result.verify_passed = verify.passed
    result.verify_violations = verify.violations

    # Classify matches
    for m in match_output.matched:
        proc_a = procs_a.get(m.key_a, {})
        proc_b = procs_b.get(m.key_b, {})

        # Determine detailed change type
        if m.change_type == "unchanged":
            # Check if raw bodies are actually identical
            if proc_a.get("body") == proc_b.get("body"):
                change_type = VBAChangeType.PROCEDURE_BODY_COSMETIC_ONLY.value
                # Actually unchanged — but we classified as "unchanged" in matcher which means canonical match
                result.summary.procedures_unchanged += 1
                continue  # Don't add to changes list
            else:
                change_type = VBAChangeType.PROCEDURE_BODY_COSMETIC_ONLY.value
                result.summary.procedures_cosmetic_only += 1
        elif m.change_type == "logic_changed":
            # Check if signature also changed
            if proc_a.get("signature") != proc_b.get("signature"):
                change_type = VBAChangeType.PROCEDURE_SIGNATURE_CHANGED.value
            else:
                change_type = VBAChangeType.PROCEDURE_BODY_LOGIC_CHANGED.value
            result.summary.procedures_logic_changed += 1
        elif m.change_type == "renamed":
            change_type = VBAChangeType.PROCEDURE_RENAMED.value
            result.summary.procedures_renamed += 1
        elif m.change_type == "renamed_modified":
            change_type = VBAChangeType.PROCEDURE_RENAMED_AND_MODIFIED.value
            result.summary.procedures_renamed_modified += 1
        else:
            change_type = m.change_type
            result.summary.procedures_logic_changed += 1

        result.procedure_changes.append(
            VBAProcedureChange(
                module_a=proc_a.get("module"),
                module_b=proc_b.get("module"),
                name_a=proc_a.get("name"),
                name_b=proc_b.get("name"),
                kind=proc_a.get("kind", "unknown"),
                change_type=change_type,
                confidence=m.confidence,
                match_pass=m.match_pass,
                body_a=proc_a.get("body"),
                body_b=proc_b.get("body"),
                signature_a=proc_a.get("signature"),
                signature_b=proc_b.get("signature"),
            )
        )

    # Added procedures
    for key in match_output.added:
        proc_b = procs_b.get(key, {})
        result.procedure_changes.append(
            VBAProcedureChange(
                module_a=None,
                module_b=proc_b.get("module"),
                name_a=None,
                name_b=proc_b.get("name"),
                kind=proc_b.get("kind", "unknown"),
                change_type=VBAChangeType.PROCEDURE_ADDED.value,
                confidence=1.0,
                match_pass=0,
                body_b=proc_b.get("body"),
                signature_b=proc_b.get("signature"),
            )
        )
    result.summary.procedures_added = len(match_output.added)

    # Removed procedures
    for key in match_output.removed:
        proc_a = procs_a.get(key, {})
        result.procedure_changes.append(
            VBAProcedureChange(
                module_a=proc_a.get("module"),
                module_b=None,
                name_a=proc_a.get("name"),
                name_b=None,
                kind=proc_a.get("kind", "unknown"),
                change_type=VBAChangeType.PROCEDURE_REMOVED.value,
                confidence=1.0,
                match_pass=0,
                body_a=proc_a.get("body"),
                signature_a=proc_a.get("signature"),
            )
        )
    result.summary.procedures_removed = len(match_output.removed)

    return result


def _load_procedures_from_db(db_path: str | Path) -> dict[str, dict]:
    """Load procedures from an ir.db's vba_procedures table."""
    try:
        conn = sqlite3.connect(str(db_path))
        # Check if Phase 3 tables exist
        try:
            conn.execute("SELECT 1 FROM vba_procedures LIMIT 1")
        except sqlite3.OperationalError:
            conn.close()
            return {}

        rows = conn.execute("""
            SELECT m.name AS module_name, p.name, p.kind, p.signature, p.body,
                   p.normalized_body_hash
            FROM vba_procedures p
            JOIN vba_modules m ON p.module_id = m.module_id
        """).fetchall()
        conn.close()

        procs = {}
        for module, name, kind, sig, body, norm_hash in rows:
            key = f"{module}::{name}::{kind}"
            procs[key] = {
                "module": module,
                "name": name,
                "kind": kind,
                "signature": sig,
                "body": body,
                "normalized_body_hash": norm_hash,
            }
        return procs
    except Exception:
        return {}


def _load_module_names(db_path: str | Path) -> set[str]:
    """Load module names from an ir.db's vba_modules table."""
    try:
        conn = sqlite3.connect(str(db_path))
        try:
            names = {r[0] for r in conn.execute("SELECT name FROM vba_modules").fetchall()}
        except sqlite3.OperationalError:
            names = set()
        conn.close()
        return names
    except Exception:
        return set()
