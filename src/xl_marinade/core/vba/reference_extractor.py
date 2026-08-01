# ABOUTME: Static VBA reference extractor — extracts cell/range references from VBA
# ABOUTME: procedure bodies using pattern matching (Tier 1 literals + Tier 2 context).

"""
VBA Static Reference Extractor (Phase 2a + 2b)

Tier 1 (Phase 2a) — Literal patterns:
- Range("literal") with optional sheet qualification
- Sheets("literal").Range("literal")
- Cells(literal_row, literal_col) with optional sheet qualification
- [ShortcutNotation] including [Sheet!Range]
- .Offset(literal, literal) on literal base ranges
- Read/write classification

Tier 2 (Phase 2b) — Context-aware:
- Worksheet variable tracking: Set ws = Worksheets("literal") → ws.Range("C2") resolves
- With block context: With Sheets("X") → .Range("A1") resolves to X!A1
- Named range resolution via defined_names table
- .Offset on named range bases (resolves base, applies offset)

Does NOT handle (deferred to Phase 3/LLM):
- Selection-based operations
- Dynamic/variable cell references
- String-built addresses
"""

import json
import re
import sqlite3
from dataclasses import dataclass


@dataclass
class VBACellRef:
    """A cell/range reference extracted from VBA source code."""

    sheet: str | None  # None = unqualified (active sheet or needs context)
    target: str  # Cell address, range, or named range name
    ref_kind: str  # 'read', 'write', 'read_write'
    target_kind: str  # 'cell_range', 'named_range', 'row_col'
    precision: str  # 'exact', 'static_only'
    line_number: int | None = None


class WorksheetVariableTracker:
    """Track Dim/Set assignments of Worksheet-typed variables within a procedure.

    Resolves patterns like:
        Set wsCalc = wb.Worksheets("Calculations")
        wsCalc.Range("C2").Value = j  → Calculations!C2

    Assumes single assignment per variable (reassignment is rare in actuarial VBA;
    0 occurrences in the test set). Phase 3 LLM handles the rare reassignment case.
    """

    def __init__(self) -> None:
        self.ws_vars: dict[str, str] = {}  # lowercase var name → sheet name

    # Set ws = [wb.]Worksheets("literal") or Sheets("literal")
    _SET_WS_RE = re.compile(
        r'Set\s+(\w+)\s*=\s*\w*\.?(?:Worksheets?|Sheets)\("([^"]+)"\)',
        re.IGNORECASE,
    )
    # Dim ws As Worksheet (just registers the variable name)
    _DIM_WS_RE = re.compile(
        r"Dim\s+(\w+)\s+As\s+Worksheet",
        re.IGNORECASE,
    )

    def process_line(self, line: str) -> None:
        m = self._SET_WS_RE.search(line)
        if m:
            self.ws_vars[m.group(1).lower()] = m.group(2)
            return
        m = self._DIM_WS_RE.search(line)
        if m:
            # Register variable but don't assign sheet yet
            self.ws_vars.setdefault(m.group(1).lower(), "")

    def resolve(self, var_name: str) -> str | None:
        """Return the sheet name for a known worksheet variable, or None."""
        sheet = self.ws_vars.get(var_name.lower())
        return sheet if sheet else None


class WithBlockTracker:
    """Track With ... End With block targets for resolving dot-references.

    Resolves patterns like:
        With Sheets("X")
            .Range("A1").Value = 1   → X!A1
        End With

    Handles nested With blocks by composing parent + child targets.
    """

    def __init__(self) -> None:
        self.stack: list[str] = []

    def process_line(self, line: str) -> None:
        stripped = line.strip()
        upper = stripped.upper()
        if upper.startswith("WITH "):
            target = stripped[5:].strip()
            if self.stack and target.startswith("."):
                target = self.stack[-1] + target
            self.stack.append(target)
        elif upper == "END WITH":
            if self.stack:
                self.stack.pop()

    def resolve_dot_reference(self, ref: str) -> str | None:
        """Resolve a .Something reference using the current With context."""
        if not ref.startswith(".") or not self.stack:
            return None
        return self.stack[-1] + ref

    @property
    def current_target(self) -> str | None:
        return self.stack[-1] if self.stack else None


# Regex to detect ws_var.Range("literal") or ws_var.Cells(literal, literal)
_WS_VAR_RANGE_RE = re.compile(
    r'(\w+)\.Range\("([^"]+)"\)',
    re.IGNORECASE,
)
_WS_VAR_CELLS_RE = re.compile(
    r'(\w+)\.Cells\(\s*(\d+)\s*,\s*(?:"([A-Z]+)"|(\d+))\s*\)',
    re.IGNORECASE,
)
# Detect sheet name from With target: Sheets("name") or Worksheets("name")
_WITH_SHEET_RE = re.compile(
    r'(?:\w+\.)?(?:Sheets|Worksheets)\("([^"]+)"\)',
    re.IGNORECASE,
)


def _strip_comments(body: str) -> list[tuple[int, str]]:
    """Strip VBA line comments only, returning (line_number, clean_line) pairs.

    Does NOT strip string contents — those are needed for Range("literal") extraction.
    Comments are removed because they can contain misleading reference patterns.
    """
    result = []
    for i, line in enumerate(body.splitlines(), 1):
        # Remove line comments (after ' outside strings)
        in_string = False
        clean_end = len(line)
        for j, ch in enumerate(line):
            if ch == '"':
                in_string = not in_string
            elif ch == "'" and not in_string:
                clean_end = j
                break
        result.append((i, line[:clean_end]))
    return result


# --- Pattern matchers ---

# Sheet qualification: Sheets("name") or Worksheets("name") with optional wb prefix
_SHEET_QUAL_RE = re.compile(
    r'(?:\w+\.)?(?:Sheets|Worksheets)\("([^"]+)"\)\.',
    re.IGNORECASE,
)

# Range("literal") — captures the address/name
_RANGE_LITERAL_RE = re.compile(
    r'(?:(?:\w+\.)?(?:Sheets|Worksheets)\("([^"]+)"\)\.)?Range\("([^"]+)"\)',
    re.IGNORECASE,
)

# Cells(literal_row, literal_col) — both must be integer literals
_CELLS_LITERAL_RE = re.compile(
    r'(?:(?:\w+\.)?(?:Sheets|Worksheets)\("([^"]+)"\)\.)?Cells\(\s*(\d+)\s*,\s*(\d+)\s*\)',
    re.IGNORECASE,
)

# Shortcut notation: [A1], [Sheet1!A1:B10]
_SHORTCUT_RE = re.compile(
    r"\[([A-Za-z_][\w]*!)?([A-Z]+\d+(?::[A-Z]+\d+)?)\]",
)

# Offset(literal_row, literal_col) applied to a preceding Range/Cells reference
_OFFSET_RE = re.compile(
    r"\.Offset\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)",
    re.IGNORECASE,
)

# Write patterns (applied to the full line to classify read vs write)
_WRITE_PATTERNS = [
    re.compile(r"\.Value\s*=\s*(?!=)", re.IGNORECASE),  # .Value = expr (not ==)
    re.compile(r"\.Formula\s*=", re.IGNORECASE),
    re.compile(r"\.FormulaR1C1\s*=", re.IGNORECASE),
    re.compile(r"\.ClearContents\b", re.IGNORECASE),
    re.compile(r"\.Clear\b(?!C)", re.IGNORECASE),  # .Clear but not .ClearContents
    re.compile(r"\.Delete\b", re.IGNORECASE),
    re.compile(r"\.EntireRow\.(?:Delete|ClearContents|Clear)\b", re.IGNORECASE),
    re.compile(r"PasteSpecial\s+Paste:=xlPaste(?:Values|Formulas|All)\b", re.IGNORECASE),
]

_READ_PATTERNS = [
    re.compile(r"=\s*\S+\.Value\b", re.IGNORECASE),
    re.compile(r"If\s+.*\.Value\b", re.IGNORECASE),
    re.compile(r"\.Copy\b", re.IGNORECASE),
]

# Format-only operations (not data edges)
_FORMAT_ONLY_RE = re.compile(
    r"PasteSpecial\s+Paste:=xlPasteFormats\b|\.Interior\.|\.Font\.|\.NumberFormat",
    re.IGNORECASE,
)


def _is_named_range(target: str) -> bool:
    """Check if a target looks like a named range (vs a cell address)."""
    # Cell addresses match A1, AB123, A1:B10 etc.
    if re.match(r"^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$", target, re.IGNORECASE):
        return False
    # Everything else is likely a named range
    return True


def _classify_ref_kind(line: str) -> str:
    """Classify a line as read, write, or read_write based on context patterns."""
    # Skip format-only operations entirely
    if _FORMAT_ONLY_RE.search(line):
        return "format_only"

    is_write = any(p.search(line) for p in _WRITE_PATTERNS)
    is_read = any(p.search(line) for p in _READ_PATTERNS)

    # Special case: x = Range("A1").Value on a line with no write → read
    if not is_write and not is_read:
        # Default: if .Value appears on the right side of =, it's a read
        if re.search(r"=\s*.*(?:Range|Cells|\[)", line, re.IGNORECASE):
            is_read = True

    if is_write and is_read:
        return "read_write"
    elif is_write:
        return "write"
    elif is_read:
        return "read"
    # If we can't determine, it's likely a read (e.g., passing range to a function)
    return "read"


def _col_num_to_letter(col: int) -> str:
    """Convert 1-based column number to Excel column letter."""
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result or "A"


def extract_references(body: str, named_ranges: dict[str, str] | None = None) -> list[VBACellRef]:
    """Extract cell/range references from a VBA procedure body.

    Handles Tier 1 (literal) and Tier 2 (context-aware) patterns.

    Args:
        body: VBA procedure source code
        named_ranges: Optional dict of uppercase name → resolved cell address (e.g., "Projection!C3")
    """
    refs: list[VBACellRef] = []
    lines = _strip_comments(body)
    ws_tracker = WorksheetVariableTracker()
    _with_tracker = WithBlockTracker()

    # First pass: build worksheet variable assignments and With block context
    for _, line in lines:
        ws_tracker.process_line(line)

    # Second pass: extract references with context
    with_tracker_pass = WithBlockTracker()
    for line_num, line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Update With-block tracking
        with_tracker_pass.process_line(line)

        ref_kind = _classify_ref_kind(line)
        if ref_kind == "format_only":
            continue

        # --- Tier 2b: Worksheet variable references (wsVar.Range / wsVar.Cells) ---
        for m in _WS_VAR_RANGE_RE.finditer(line):
            var_name = m.group(1)
            target = m.group(2)
            # Skip if this is a Sheets("...").Range pattern (already handled in Tier 1)
            if var_name.lower() in ("sheets", "worksheets", "range", "application"):
                continue
            resolved_sheet = ws_tracker.resolve(var_name)
            if resolved_sheet:
                target_kind = "named_range" if _is_named_range(target) else "cell_range"
                # Check for Offset
                after_match = line[m.end() :]
                offset_match = _OFFSET_RE.match(after_match)
                if offset_match and target_kind == "cell_range" and ":" not in target:
                    try:
                        row_off = int(offset_match.group(1))
                        col_off = int(offset_match.group(2))
                        addr_m = re.match(r"([A-Z]+)(\d+)", target, re.IGNORECASE)
                        if addr_m:
                            base_col = _col_letter_to_num(addr_m.group(1))
                            base_row = int(addr_m.group(2))
                            nr, nc = base_row + row_off, base_col + col_off
                            if nr > 0 and nc > 0:
                                target = f"{_col_num_to_letter(nc)}{nr}"
                    except (ValueError, IndexError):
                        pass
                # Resolve named range if possible
                if target_kind == "named_range" and named_ranges:
                    resolved = named_ranges.get(target.upper())
                    if resolved:
                        # resolved is like "Projection!C3" — use it directly
                        if "!" in resolved:
                            parts = resolved.split("!", 1)
                            resolved_sheet = parts[0].strip("'")
                            target = parts[1].lstrip("$").replace("$", "")
                            target_kind = "cell_range"
                refs.append(
                    VBACellRef(
                        sheet=resolved_sheet,
                        target=target,
                        ref_kind=ref_kind,
                        target_kind=target_kind,
                        precision="exact",
                        line_number=line_num,
                    )
                )

        for m in _WS_VAR_CELLS_RE.finditer(line):
            var_name = m.group(1)
            if var_name.lower() in ("sheets", "worksheets", "application"):
                continue
            resolved_sheet = ws_tracker.resolve(var_name)
            if resolved_sheet:
                row = int(m.group(2))
                # Column can be letter string or number
                col_letter = m.group(3)
                col_num = m.group(4)
                if col_letter:
                    col = _col_letter_to_num(col_letter)
                elif col_num:
                    col = int(col_num)
                else:
                    continue
                if row > 0 and col > 0:
                    target = f"{_col_num_to_letter(col)}{row}"
                    refs.append(
                        VBACellRef(
                            sheet=resolved_sheet,
                            target=target,
                            ref_kind=ref_kind,
                            target_kind="cell_range",
                            precision="exact",
                            line_number=line_num,
                        )
                    )

        # --- Tier 2a: With-block context for dot-references ---
        with_target = with_tracker_pass.current_target
        if with_target:
            # Check for .Range("literal") inside With block
            for m in re.finditer(r'(?<!\w)\.Range\("([^"]+)"\)', line):
                target = m.group(1)
                # Extract sheet from With target
                sheet_m = _WITH_SHEET_RE.search(with_target)
                if sheet_m:
                    sheet = sheet_m.group(1)
                    target_kind = "named_range" if _is_named_range(target) else "cell_range"
                    refs.append(
                        VBACellRef(
                            sheet=sheet,
                            target=target,
                            ref_kind=ref_kind,
                            target_kind=target_kind,
                            precision="exact",
                            line_number=line_num,
                        )
                    )

        # --- Tier 1: Literal patterns ---

        # 1. Range("literal") with optional sheet qualification
        for m in _RANGE_LITERAL_RE.finditer(line):
            sheet = m.group(1)  # None if unqualified
            target = m.group(2)

            # Check if this Range has .Offset(literal, literal) applied
            after_range = line[m.end() :]
            offset_match = _OFFSET_RE.match(after_range)

            target_kind = "named_range" if _is_named_range(target) else "cell_range"

            # If offset on a literal cell address, compute the resulting address
            if offset_match and target_kind == "cell_range" and ":" not in target:
                try:
                    row_offset = int(offset_match.group(1))
                    col_offset = int(offset_match.group(2))
                    # Parse base address
                    addr_match = re.match(r"([A-Z]+)(\d+)", target, re.IGNORECASE)
                    if addr_match:
                        base_col_str = addr_match.group(1).upper()
                        base_row = int(addr_match.group(2))
                        # Convert col letter to number
                        base_col = 0
                        for ch in base_col_str:
                            base_col = base_col * 26 + (ord(ch) - 64)
                        new_row = base_row + row_offset
                        new_col = base_col + col_offset
                        if new_row > 0 and new_col > 0:
                            target = f"{_col_num_to_letter(new_col)}{new_row}"
                except (ValueError, IndexError):
                    pass  # Can't compute offset — keep original target

            refs.append(
                VBACellRef(
                    sheet=sheet,
                    target=target,
                    ref_kind=ref_kind,
                    target_kind=target_kind,
                    precision="exact" if sheet else "static_only",
                    line_number=line_num,
                )
            )

        # 2. Cells(literal_row, literal_col) with optional sheet
        for m in _CELLS_LITERAL_RE.finditer(line):
            sheet = m.group(1)
            row = int(m.group(2))
            col = int(m.group(3))
            if row > 0 and col > 0:
                target = f"{_col_num_to_letter(col)}{row}"
                refs.append(
                    VBACellRef(
                        sheet=sheet,
                        target=target,
                        ref_kind=ref_kind,
                        target_kind="cell_range",
                        precision="exact" if sheet else "static_only",
                        line_number=line_num,
                    )
                )

        # 3. Shortcut notation [A1] or [Sheet1!A1:B10]
        for m in _SHORTCUT_RE.finditer(line):
            sheet_part = m.group(1)
            sheet = sheet_part.rstrip("!") if sheet_part else None
            target = m.group(2)
            refs.append(
                VBACellRef(
                    sheet=sheet,
                    target=target,
                    ref_kind=ref_kind,
                    target_kind="cell_range",
                    precision="exact" if sheet else "static_only",
                    line_number=line_num,
                )
            )

    # --- Tier 2c: Named range resolution ---
    if named_ranges:
        resolved_refs: list[VBACellRef] = []
        for ref in refs:
            if ref.target_kind == "named_range":
                resolved = named_ranges.get(ref.target.upper())
                if resolved and "!" in resolved:
                    parts = resolved.split("!", 1)
                    resolved_sheet = parts[0].strip("'")
                    resolved_addr = parts[1].lstrip("$").replace("$", "")
                    resolved_refs.append(
                        VBACellRef(
                            sheet=ref.sheet or resolved_sheet,
                            target=resolved_addr,
                            ref_kind=ref.ref_kind,
                            target_kind="cell_range",
                            precision="exact",
                            line_number=ref.line_number,
                        )
                    )
                else:
                    resolved_refs.append(ref)
            else:
                resolved_refs.append(ref)
        refs = resolved_refs

    # Deduplicate: prefer qualified (sheet-resolved) refs over unqualified.
    # If we have both (Calculations, C2, write) and (None, C2, write), keep only the qualified one.
    by_target: dict[tuple[str, str], VBACellRef] = {}  # (target_upper, ref_kind) → best ref
    for ref in refs:
        key = (ref.target.upper(), ref.ref_kind)
        existing = by_target.get(key)
        if existing is None:
            by_target[key] = ref
        elif ref.sheet and not existing.sheet:
            # Prefer qualified over unqualified
            by_target[key] = ref
        elif ref.sheet and existing.sheet and ref.sheet != existing.sheet:
            # Different sheets for same target — keep both (use full key)
            full_key = (ref.sheet, ref.target.upper(), ref.ref_kind)
            by_target[full_key] = ref  # type: ignore[assignment]

    return list(by_target.values())


def extract_vba_references_to_table(conn: sqlite3.Connection) -> dict[str, int]:
    """Phase 1 of VBA reference handling: extract cell references from procedure
    bodies and store them in vba_procedure_cell_refs. Does NOT create binding
    edges (that needs bindings — see extract_and_store_references).

    Idempotent: a no-op if the table is already populated. Safe to run EARLY
    (pipeline step 5.25, before grouping) so binding-geometry passes (H5) can read
    the refs; the later edge-wiring step (6.9) reuses the rows written here.
    """
    metrics = {"refs_extracted": 0}

    tables = {
        r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    }
    if "vba_procedures" not in tables or "vba_modules" not in tables:
        return metrics
    if "vba_procedure_cell_refs" not in tables:
        return metrics

    existing = conn.execute("SELECT COUNT(*) FROM vba_procedure_cell_refs").fetchone()[0]
    if existing:
        metrics["refs_extracted"] = existing
        return metrics

    # Load named ranges for resolution (Tier 2c)
    named_ranges: dict[str, str] = {}
    if "defined_names" in tables:
        try:
            for row in conn.execute("SELECT name, destinations FROM defined_names"):
                try:
                    dests = json.loads(row[1])
                    if isinstance(dests, list) and dests:
                        dest = dests[0]
                        if isinstance(dest, str) and not dest.startswith("#"):
                            named_ranges[row[0].upper()] = dest
                except (json.JSONDecodeError, IndexError):
                    pass
        except sqlite3.OperationalError:
            pass

    procs = conn.execute("""
        SELECT p.procedure_id, p.name, p.body, m.name AS module_name,
               p.kind, p.compile_branch
        FROM vba_procedures p
        JOIN vba_modules m ON p.module_id = m.module_id
    """).fetchall()

    all_refs: list[tuple] = []
    for proc in procs:
        proc_id = proc[0]
        body = proc[2] or ""
        refs = extract_references(body, named_ranges=named_ranges)
        for ref in refs:
            full_target = f"{ref.sheet}!{ref.target}" if ref.sheet else ref.target
            all_refs.append(
                (
                    proc_id,
                    ref.ref_kind,
                    ref.target_kind,
                    full_target,
                    ref.precision,
                )
            )

    if all_refs:
        conn.executemany(
            """INSERT INTO vba_procedure_cell_refs
               (procedure_id, ref_kind, target_kind, target, precision)
               VALUES (?, ?, ?, ?, ?)""",
            all_refs,
        )
        metrics["refs_extracted"] = len(all_refs)
        conn.commit()
    return metrics


def extract_and_store_references(conn: sqlite3.Connection) -> dict[str, int]:
    """Extract VBA cell references (phase 1, idempotent) AND convert resolved
    references to binding_edges (phase 2, needs bindings).

    Returns metrics dict.
    """
    metrics = {"refs_extracted": 0, "edges_created": 0}

    # Check tables exist
    tables = {
        r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    }
    if "vba_procedures" not in tables or "vba_modules" not in tables:
        return metrics
    if "vba_procedure_cell_refs" not in tables:
        return metrics

    # Phase 1: populate vba_procedure_cell_refs (idempotent; may already have run
    # at step 5.25 so the H5 binding pass could read the refs — then a no-op here).
    metrics["refs_extracted"] = extract_vba_references_to_table(conn).get("refs_extracted", 0)

    # --- Convert cell_range refs to binding_edges ---
    # For refs with target_kind='cell_range' and a sheet, try to resolve to bindings
    if "cell_to_binding" not in tables or "cells" not in tables:
        conn.commit()
        return metrics

    # Build sheet name → sheet_id lookup
    sheet_map: dict[str, int] = {}
    try:
        for row in conn.execute("SELECT sheet_id, sheet_name FROM sheets"):
            sheet_map[row[0]] = row[1]
            sheet_map[row[1].lower()] = row[0]  # name → id
    except sqlite3.OperationalError:
        pass

    # Procedure list (id/name/module/kind/branch) for building vba:: node ids and
    # iterating each procedure's stored cell refs. (Phase-1 ref extraction now runs
    # separately, so re-query here for the edge-wiring phase.)
    procs = conn.execute("""
        SELECT p.procedure_id, p.name, p.body, m.name AS module_name,
               p.kind, p.compile_branch
        FROM vba_procedures p
        JOIN vba_modules m ON p.module_id = m.module_id
    """).fetchall()

    edge_inserts: list[tuple[str, str]] = []

    for proc in procs:
        proc_id = proc[0]
        module_name = proc[3]
        kind = proc[4]
        compile_branch = proc[5] or ""
        proc_binding_id = f"vba::{module_name}::{proc[1]}::{kind}"
        if compile_branch:
            proc_binding_id += f"::{compile_branch}"

        # Get this procedure's resolved cell_range refs
        proc_refs = conn.execute(
            """
            SELECT ref_kind, target FROM vba_procedure_cell_refs
            WHERE procedure_id = ? AND target_kind = 'cell_range' AND target LIKE '%!%'
        """,
            (proc_id,),
        ).fetchall()

        for ref_row in proc_refs:
            ref_kind = ref_row[0]
            target = ref_row[1]  # e.g., "Calculation Engine!C2"

            # Parse sheet!address
            parts = target.split("!", 1)
            if len(parts) != 2:
                continue
            sheet_name, address = parts

            # Resolve sheet name to ID
            sheet_id = sheet_map.get(sheet_name.lower())
            if sheet_id is None:
                continue

            # Parse address to find cells, look up in cell_to_binding
            # For single cells (A1) or small ranges (A1:B10), iterate
            binding_ids: set[str] = set()

            if ":" in address:
                # Range — use bounding box approach via cells table
                try:
                    addr_match = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", address, re.IGNORECASE)
                    if addr_match:
                        # Find bindings that overlap this range via cell_to_binding
                        rows = conn.execute(
                            """
                            SELECT DISTINCT ctb.binding_id
                            FROM cells c
                            JOIN cell_to_binding ctb ON c.cell_id = ctb.cell_id
                            WHERE c.sheet_id = ?
                              AND c.row BETWEEN ? AND ?
                              AND c.col BETWEEN ? AND ?
                        """,
                            (
                                sheet_id,
                                int(addr_match.group(2)),  # r1
                                int(addr_match.group(4)),  # r2
                                _col_letter_to_num(addr_match.group(1)),  # c1
                                _col_letter_to_num(addr_match.group(3)),  # c2
                            ),
                        ).fetchall()
                        for r in rows:
                            binding_ids.add(r[0])
                except (ValueError, sqlite3.OperationalError):
                    pass
            else:
                # Single cell
                try:
                    addr_match = re.match(r"([A-Z]+)(\d+)", address, re.IGNORECASE)
                    if addr_match:
                        row_num = int(addr_match.group(2))
                        col_num = _col_letter_to_num(addr_match.group(1))
                        from xl_marinade.core.new_arch.cell_identity import pack as pack_cell_id

                        cell_id = pack_cell_id(sheet_id, row_num, col_num)
                        rows = conn.execute(
                            "SELECT binding_id FROM cell_to_binding WHERE cell_id = ?", (cell_id,)
                        ).fetchall()
                        for r in rows:
                            binding_ids.add(r[0])
                except (ValueError, sqlite3.OperationalError):
                    pass

            # Create edges based on ref_kind
            for bid in binding_ids:
                if ref_kind in ("read", "read_write"):
                    # Procedure reads from binding → proc depends on binding
                    edge_inserts.append((proc_binding_id, bid))
                if ref_kind in ("write", "read_write"):
                    # Procedure writes to binding → binding depends on proc
                    edge_inserts.append((bid, proc_binding_id))

    if edge_inserts:
        conn.executemany(
            "INSERT OR IGNORE INTO binding_edges (from_binding_id, to_binding_id, edge_count) VALUES (?, ?, 1)",
            edge_inserts,
        )
        metrics["edges_created"] = len(edge_inserts)

    conn.commit()
    return metrics


def _col_letter_to_num(col_str: str) -> int:
    """Convert column letter(s) to 1-based number. A=1, B=2, AA=27."""
    result = 0
    for ch in col_str.upper():
        result = result * 26 + (ord(ch) - 64)
    return result
