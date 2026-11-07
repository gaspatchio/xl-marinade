# ABOUTME: Stage 5 VBA emitter — compares vba_modules/vba_procedures between two IR dbs.
# ABOUTME: Emits VBA_* Change records (module, declarations, procedure) for diff_ir.

"""
VBA diff: module-, declaration- and procedure-level changes.

``diff_vba`` reads the VBA tables of two IR databases and returns ``Change``
records for ``diff_ir`` to order, verify and serialize with every other change.
Procedures are paired by the 4-pass matcher in ``vba_match``. A procedure is
keyed by module, name, kind and ``#If`` compile branch, so two branches of one
conditionally-compiled procedure are compared branch to branch instead of
overwriting each other.

Code outside every procedure (``Option`` statements, ``Declare``s, module-level
variables, constants, comments and ``Attribute`` lines) is compared per module
as the module's *declarations*. With procedure bodies and declarations both
compared, any edit to a module's source surfaces as at least one change.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from xl_marinade.core.db_uri import connect_read_only
from xl_marinade.core.ir_diff import change_types as CT
from xl_marinade.core.ir_diff.model import Change, DiffVerificationError
from xl_marinade.core.ir_diff.vba_canonicalize import canonicalize_vba_body
from xl_marinade.core.ir_diff.vba_match import (
    MatchResult,
    ProcedureDesc,
    _build_desc,
    match_procedures,
)
from xl_marinade.core.ir_diff.vba_verify import verify_vba_match


@dataclass(slots=True, frozen=True)
class _Procedure:
    """One vba_procedures row, joined to its module's name."""

    module: str
    name: str
    kind: str
    compile_branch: str
    signature: str
    body: str
    line_start: int | None
    line_end: int | None

    @property
    def key(self) -> str:
        return f"{self.module}::{self.name}::{self.kind}::{self.compile_branch}"

    @property
    def lines(self) -> str | None:
        if self.line_start is None or self.line_end is None:
            return None
        return f"{self.line_start}-{self.line_end}"


@dataclass(slots=True)
class _Module:
    """One vba_modules row plus the line ranges its procedures occupy."""

    name: str
    kind: str
    source_text: str
    procedure_ranges: list[tuple[int, int]] = field(default_factory=list)
    # False when any procedure has no recorded line range: the code outside
    # procedures cannot then be separated from procedure bodies.
    ranges_known: bool = True


@dataclass(slots=True)
class _VBASide:
    """All VBA in one IR database. Empty for a workbook without VBA."""

    modules: dict[str, _Module] = field(default_factory=dict)
    procedures: dict[str, _Procedure] = field(default_factory=dict)


def diff_vba(db_a: str | Path, db_b: str | Path) -> list[Change]:
    """Compare the VBA of two IR databases.

    A database without VBA tables (or with empty ones) has no VBA; comparing
    two such databases returns no changes.

    Args:
        db_a: Path to Version A (earlier) IR database.
        db_b: Path to Version B (later) IR database.

    Returns:
        Unordered ``VBA_*`` changes; ``diff_ir`` orders them.

    Raises:
        DiffVerificationError: If the procedure matcher's output fails its
            invariant checks.
    """
    a = _load_vba(db_a)
    b = _load_vba(db_b)
    if not a.modules and not b.modules:
        return []

    changes = _diff_modules(a, b)
    changes += _diff_procedures(a, b)
    changes += _diff_declarations(a, b)
    return changes


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------


def _load_vba(db_path: str | Path) -> _VBASide:
    """Read vba_modules and vba_procedures. Missing tables mean no VBA.

    Any other read failure propagates: a corrupt database must not read as
    "no VBA" and hide every VBA change (``xl_marinade.diff`` wraps it as
    ``DiffError``).
    """
    side = _VBASide()
    conn = connect_read_only(db_path)
    try:
        tables = {
            r[0]
            for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table' "
                "AND name IN ('vba_modules', 'vba_procedures')"
            )
        }
        if "vba_modules" not in tables:
            return side
        for name, kind, source_text in conn.execute(
            "SELECT name, kind, source_text FROM vba_modules"
        ):
            side.modules[name] = _Module(name=name, kind=kind, source_text=source_text or "")
        if "vba_procedures" not in tables:
            return side
        rows: list[tuple] = conn.execute(
            """
            SELECT m.name, p.name, p.kind, p.compile_branch, p.signature, p.body,
                   p.line_start, p.line_end
            FROM vba_procedures p
            JOIN vba_modules m ON p.module_id = m.module_id
            """
        ).fetchall()
    finally:
        conn.close()

    for module, name, kind, branch, signature, body, line_start, line_end in rows:
        proc = _Procedure(
            module=module,
            name=name,
            kind=kind,
            compile_branch=branch or "",
            signature=signature or "",
            body=body or "",
            line_start=line_start,
            line_end=line_end,
        )
        side.procedures[proc.key] = proc
        mod = side.modules[module]
        if line_start is None or line_end is None:
            mod.ranges_known = False
        else:
            mod.procedure_ranges.append((line_start, line_end))
    return side


# ---------------------------------------------------------------------------
# Modules
# ---------------------------------------------------------------------------


def _diff_modules(a: _VBASide, b: _VBASide) -> list[Change]:
    changes: list[Change] = []
    for name in sorted(set(a.modules) | set(b.modules)):
        if name in a.modules and name in b.modules:
            continue
        if name in a.modules:
            change_type, side = CT.VBA_MODULE_REMOVED, a
        else:
            change_type, side = CT.VBA_MODULE_ADDED, b
        changes.append(
            Change(
                type=change_type,
                sort_key=(name,),
                details={
                    "module": name,
                    "module_kind": side.modules[name].kind,
                    "procedure_count": sum(1 for p in side.procedures.values() if p.module == name),
                },
            )
        )
    return changes


# ---------------------------------------------------------------------------
# Procedures
# ---------------------------------------------------------------------------


def _desc(proc: _Procedure) -> ProcedureDesc:
    # The matcher pairs on (module, name, kind) equality. Folding the compile
    # branch into the name keeps `#If` twins paired with their own branch.
    match_name = f"{proc.name}#{proc.compile_branch}" if proc.compile_branch else proc.name
    return _build_desc(proc.key, proc.module, match_name, proc.kind, proc.body)


def _diff_procedures(a: _VBASide, b: _VBASide) -> list[Change]:
    descs_a = {k: _desc(p) for k, p in a.procedures.items()}
    descs_b = {k: _desc(p) for k, p in b.procedures.items()}
    output = match_procedures(descs_a, descs_b)

    verify = verify_vba_match(output, len(descs_a), len(descs_b))
    if not verify.passed:
        raise DiffVerificationError(
            "VBA procedure matching failed verification:\n  " + "\n  ".join(verify.violations)
        )

    changes: list[Change] = []
    for m in output.matched:
        change = _matched_change(a.procedures[m.key_a], b.procedures[m.key_b], m)
        if change is not None:
            changes.append(change)
    for key in output.added:
        changes.append(_one_sided_change(CT.VBA_PROCEDURE_ADDED, b.procedures[key]))
    for key in output.removed:
        changes.append(_one_sided_change(CT.VBA_PROCEDURE_REMOVED, a.procedures[key]))
    # The matcher gives up on an A procedure with several equally plausible
    # rename targets. It is gone under its old name either way, and its
    # candidates are reported as added, so report it as removed rather than
    # dropping it from the changelist.
    for key in sorted(output.ambiguous_dropped):
        change = _one_sided_change(CT.VBA_PROCEDURE_REMOVED, a.procedures[key])
        change.details["ambiguous_rename"] = True
        changes.append(change)
    return changes


def _canonical_without_name(proc: _Procedure) -> str:
    """Canonical body with the procedure's own name (its declaration, a
    Function's return assignment, a recursive call) replaced by a placeholder."""
    canonical = canonicalize_vba_body(proc.body)
    return re.sub(rf"\b{re.escape(proc.name.lower())}\b", "\x00", canonical)


def _proc_details(proc: _Procedure) -> dict[str, object]:
    return {
        "module": proc.module,
        "procedure": proc.name,
        "kind": proc.kind,
        "compile_branch": proc.compile_branch,
    }


def _one_sided_change(change_type: str, proc: _Procedure) -> Change:
    details = _proc_details(proc)
    details["signature"] = proc.signature
    details["lines"] = proc.lines
    return Change(
        type=change_type,
        sort_key=(proc.module, proc.name, proc.kind, proc.compile_branch),
        details=details,
    )


def _matched_change(pa: _Procedure, pb: _Procedure, m: MatchResult) -> Change | None:
    """Classify a matched pair. None when the procedure is byte-identical."""
    details: dict[str, object]
    if m.change_type in ("renamed", "renamed_modified"):
        # The matcher's verdict compares hashes of bodies that include the
        # declaration line, which names the procedure -- so a pure rename
        # never hashed equal. Decide with each body's own name masked.
        change_type = (
            CT.VBA_PROCEDURE_RENAMED
            if _canonical_without_name(pa) == _canonical_without_name(pb)
            else CT.VBA_PROCEDURE_RENAMED_AND_MODIFIED
        )
        details = {
            "module": pb.module,
            "procedure_a": pa.name,
            "procedure_b": pb.name,
            "kind": pb.kind,
            "compile_branch": pb.compile_branch,
        }
    elif m.change_type == "unchanged":
        # Canonical bodies match: only comments, whitespace or case differ.
        if pa.body == pb.body:
            return None
        change_type = CT.VBA_PROCEDURE_COSMETIC_CHANGED
        details = _proc_details(pb)
    elif pa.signature != pb.signature:
        change_type = CT.VBA_PROCEDURE_SIGNATURE_CHANGED
        details = _proc_details(pb)
    else:
        change_type = CT.VBA_PROCEDURE_LOGIC_CHANGED
        details = _proc_details(pb)

    if pa.signature != pb.signature:
        details["signature_a"] = pa.signature
        details["signature_b"] = pb.signature
    details["lines_a"] = pa.lines
    details["lines_b"] = pb.lines
    details["match_pass"] = m.match_pass
    details["confidence"] = m.confidence
    return Change(
        type=change_type,
        sort_key=(pb.module, pb.name, pb.kind, pb.compile_branch),
        details=details,
    )


# ---------------------------------------------------------------------------
# Declarations (module code outside every procedure)
# ---------------------------------------------------------------------------


def _declarations(module: _Module) -> str | None:
    """The module's source lines outside every procedure; None if unknown."""
    if not module.ranges_known:
        return None
    inside: set[int] = set()
    for start, end in module.procedure_ranges:
        inside.update(range(start, end + 1))
    lines = module.source_text.splitlines()
    return "\n".join(line for i, line in enumerate(lines, start=1) if i not in inside)


def _diff_declarations(a: _VBASide, b: _VBASide) -> list[Change]:
    changes: list[Change] = []
    for name in sorted(set(a.modules) & set(b.modules)):
        decl_a = _declarations(a.modules[name])
        decl_b = _declarations(b.modules[name])
        if decl_a is None or decl_b is None or decl_a == decl_b:
            continue
        canon_a = canonicalize_vba_body(decl_a)
        canon_b = canonicalize_vba_body(decl_b)
        change_type = (
            CT.VBA_DECLARATIONS_COSMETIC_CHANGED
            if canon_a == canon_b
            else CT.VBA_DECLARATIONS_CHANGED
        )
        changes.append(
            Change(
                type=change_type,
                sort_key=(name,),
                details={"module": name, "module_kind": b.modules[name].kind},
            )
        )
    return changes
