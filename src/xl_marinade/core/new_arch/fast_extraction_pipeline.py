# ABOUTME: Fast extraction pipeline orchestrator - wires together ME1-ME7 components
# ABOUTME: Implements 7-step build procedure from memory_efficient_extraction_architecture.md §6.5

"""
Fast Memory-Efficient Extraction Pipeline

Orchestrates the complete extraction pipeline using streaming XML parsing,
integer cell identity, heap-first loading, and deterministic finalization.

Build procedure (normative, from design §6.5):
1. Verify SQLite version >= 3.27.0
2. Open build connection and apply PRAGMAs
3. Create schema and staging tables
4. Load raw tables in a single transaction
5. Finalize tables using normative SQL
6. Re-enable foreign_keys and create views/indexes
7. Run VACUUM INTO to produce canonical artifact

Design reference: docs/phase2_documentation_agent/design/memory_efficient_extraction_architecture.md
"""

import bisect
import contextlib
import hashlib
import json
import logging
import re
import sqlite3
import subprocess
import sys
import time
import zipfile
from collections import deque
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from xl_marinade.errors import ExtractionError

logger = logging.getLogger(__name__)

from xl_marinade.core.lazy_formulas import SheetFormulaCache
from xl_marinade.core.lazy_values import LazyValueFetcher
from xl_marinade.core.lazy_workbook import LazyWorkbook
from xl_marinade.core.names_tables import NameTableMap
from xl_marinade.core.new_arch.bulk_loader import BulkLoader
from xl_marinade.core.new_arch.canonical_json import canonicalize_and_hash
from xl_marinade.core.new_arch.cell_identity import (
    col_to_a1,
)
from xl_marinade.core.new_arch.cell_identity import (
    pack as pack_cell_id,
)
from xl_marinade.core.new_arch.cell_identity import (
    unpack as unpack_cell_id,
)
from xl_marinade.core.new_arch.fast_parser import stream_worksheet_cells
from xl_marinade.core.new_arch.formula_normalizer import (
    FormulaContext,
    normalize_formula,
    parse_a1_ref,
    shift_shared_formula,
    shift_shared_formula_a1,
)
from xl_marinade.core.new_arch.grouping_native import run_grouping_on_fast_output
from xl_marinade.core.new_arch.memory_budget import MemoryBudgetConfig
from xl_marinade.core.new_arch.reference_extractor import Edge, ReferenceExtractor
from xl_marinade.core.new_arch.styles_parser import DateFormatInfo, parse_date_format_info
from xl_marinade.core.new_arch.workbook_catalog import WorkbookCatalog
from xl_marinade.core.parser import parse_formula
from xl_marinade.core.resolution import (
    LOOKUP_FUNCTIONS,
    VOLATILE_FUNCTIONS,
    ResolutionEngine,
    ResolutionResult,
)
from xl_marinade.core.resolution_strategies import ResolutionContext, create_index_resolution_chain

try:
    import psutil

    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False


class EmptyRootExtractionError(RuntimeError):
    """Raised when a rooted extraction visits 0 cells (root cell is empty or unreachable).

    The extraction cannot produce any bindings because the specified root cell has no
    formula and no value, or its precedent graph is empty. This is almost always caused
    by the stored root pointing at a cell that changed position between workbook versions.

    Callers (worker / gateway) must surface this as a user-visible error; never mark the
    extraction as 'ready' when this is raised.
    """

    def __init__(self, root_descriptions: list[str], sheet_names: list[str] | None = None):
        self.root_descriptions = root_descriptions
        self.sheet_names = sheet_names or []
        roots_str = ", ".join(root_descriptions) if root_descriptions else "(none)"
        sheets_str = ", ".join(sheet_names) if sheet_names else ""
        msg = (
            f"Extraction produced 0 cells from root(s): {roots_str}. "
            "The root cell is empty or has no precedents in this workbook version. "
            "Pick a root cell that contains a formula or value"
        )
        if sheets_str:
            msg += f" (available sheets: {sheets_str})"
        msg += "."
        super().__init__(msg)


def _get_rss_mb() -> float | None:
    """Return current RSS in MB when available."""
    if not HAS_PSUTIL:
        return None
    try:
        return psutil.Process().memory_info().rss / (1024 * 1024)
    except Exception:
        return None


# --- VBA edge wiring (Phase 1 of VBA Graph Integration) ---

# VBA built-in keywords that should NOT be matched as procedure calls
_VBA_KEYWORDS = frozenset(
    {
        "SELECT",
        "CLEAR",
        "CALCULATE",
        "PRINT",
        "FORMAT",
        "DELETE",
        "COPY",
        "PASTE",
        "CLOSE",
        "OPEN",
        "SAVE",
        "ACTIVATE",
        "PROTECT",
        "UNPROTECT",
        "FIND",
        "REPLACE",
        "SORT",
        "REFRESH",
        "ADD",
        "REMOVE",
        "SET",
        "GET",
        "LET",
        "DIM",
        "REDIM",
        "IF",
        "THEN",
        "ELSE",
        "ELSEIF",
        "END",
        "FOR",
        "NEXT",
        "DO",
        "LOOP",
        "WHILE",
        "WEND",
        "UNTIL",
        "WITH",
        "CASE",
        "GOTO",
        "GOSUB",
        "RETURN",
        "EXIT",
        "STOP",
        "RESUME",
        "ON",
        "ERROR",
        "CALL",
        "SUB",
        "FUNCTION",
        "PROPERTY",
        "PUBLIC",
        "PRIVATE",
        "STATIC",
        "CONST",
        "TYPE",
        "ENUM",
        "DECLARE",
        "NOT",
        "AND",
        "OR",
        "XOR",
        "MOD",
        "IS",
        "LIKE",
        "NEW",
        "NOTHING",
        "TRUE",
        "FALSE",
        "BYVAL",
        "BYREF",
        "OPTIONAL",
        "PARAMARRAY",
        "AS",
        "TO",
        "STEP",
        "IN",
        "EACH",
        "RAISE",
        "IMPLEMENTS",
        "OPTION",
        "EXPLICIT",
        "COMPARE",
        "BASE",
        "PRESERVE",
        "ERASE",
        "LBOUND",
        "UBOUND",
        "ARRAY",
        "CBOOL",
        "CBYTE",
        "CCUR",
        "CDATE",
        "CDBL",
        "CDEC",
        "CINT",
        "CLNG",
        "CSNG",
        "CSTR",
        "CVAR",
        "MSGBOX",
        "INPUTBOX",
        "DEBUG",
        "TIMER",
    }
)


def _strip_vba_comments_and_strings(body: str) -> str:
    """Remove VBA comments and string literal contents for safe regex scanning.

    Prevents false-positive procedure call detection from patterns inside
    string literals or comment text.
    """
    lines = []
    for line in body.splitlines():
        # Remove line comments (after ' outside strings)
        in_string = False
        clean_end = len(line)
        for i, ch in enumerate(line):
            if ch == '"':
                in_string = not in_string
            elif ch == "'" and not in_string:
                clean_end = i
                break
        line = line[:clean_end]
        # Replace string contents (keep delimiters for structure)
        line = re.sub(r'"[^"]*"', '""', line)
        lines.append(line)
    return "\n".join(lines)


def workbook_has_vba_modules(workbook_path: Path | str) -> bool:
    """Typed-field gate (R21 Fix A): does this workbook contain VBA modules?

    A workbook-and-question-agnostic check that opens the .xlsx/.xlsm zip and
    looks for `xl/vbaProject.bin`. The decision rule for selecting the
    extraction mode (rooted vs full_workbook) keys off this typed-field
    boolean — VBA-driven workbooks can perform paste-as-values dataflow that
    isn't visible in the formula DAG, so they need full_workbook mode to
    materialise the source template rows that the paste-edge synthesiser
    (R21 Fix B) consumes.

    Returns False on any I/O / zip error (conservative: treat unreadable
    workbooks as non-VBA, fall back to rooted extraction). Does NOT consult
    workbook content; only the presence of the OOXML VBA part.
    """
    try:
        import zipfile

        with zipfile.ZipFile(str(workbook_path), "r") as zf:
            return "xl/vbaProject.bin" in zf.namelist()
    except (zipfile.BadZipFile, OSError):
        return False


def _wire_vba_edges(conn: sqlite3.Connection, enrich: bool = False) -> dict[str, int]:
    """Create binding_edges connecting VBA procedures to cell bindings and to each other.

    Three types of edges:
    1. Cell → UDF: from cell bindings that call UDFs to the VBA procedure node
    2. Proc → Proc: from caller procedures to callee procedures (call graph)
    3. VBA → Cell refs: from static analysis of Range/Cells/Sheets patterns in VBA source

    Returns metrics dict with edge counts.
    """
    metrics = {
        "cell_udf_edges": 0,
        "call_graph_edges": 0,
        "static_ref_edges": 0,
        "static_refs_extracted": 0,
    }

    # Check if VBA tables exist
    tables = {
        r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    }

    has_vba = "vba_procedures" in tables and "vba_modules" in tables
    has_udf_calls = "cell_udf_calls" in tables and "udfs" in tables
    has_bindings = "cell_to_binding" in tables

    if not has_vba:
        return metrics

    # --- 1. Cell → UDF edges (from cell_udf_calls) ---
    # Build the to_binding_id with the same format as marinade_nodes:
    # vba::<module>::<name>::function[::<compile_branch>]
    if has_udf_calls and has_bindings:
        cell_udf_edge_sql = """
            INSERT OR IGNORE INTO binding_edges (from_binding_id, to_binding_id, edge_count)
            SELECT
                ctb.binding_id AS from_binding_id,
                'vba::' || m.name || '::' || p.name || '::function' ||
                    CASE WHEN p.compile_branch != ''
                         THEN '::' || p.compile_branch ELSE '' END
                    AS to_binding_id,
                COUNT(*) AS edge_count
            FROM cell_udf_calls cuc
            JOIN cell_to_binding ctb ON cuc.cell_id = ctb.cell_id
            JOIN udfs u ON cuc.udf_id = u.udf_id
            JOIN vba_modules m ON UPPER(u.module) = UPPER(m.name)
            JOIN vba_procedures p
                ON p.module_id = m.module_id
                AND UPPER(p.name) = UPPER(u.name)
                AND p.kind = 'function'
            GROUP BY ctb.binding_id, m.name, p.name, p.compile_branch
        """
        try:
            cursor = conn.execute(cell_udf_edge_sql)
            metrics["cell_udf_edges"] = cursor.rowcount
        except sqlite3.OperationalError as e:
            print(f"  Warning: cell→UDF edge wiring failed: {e}", file=sys.stderr)

    # --- 2. Procedure → Procedure call graph ---
    # Load all procedure names + kind + compile_branch for binding_id construction
    proc_rows = conn.execute("""
        SELECT p.procedure_id, p.name, p.body, m.name AS module_name,
               p.kind, p.compile_branch
        FROM vba_procedures p
        JOIN vba_modules m ON p.module_id = m.module_id
    """).fetchall()

    if not proc_rows:
        conn.commit()
        return metrics

    def _build_proc_id(module_name: str, proc_name: str, kind: str, branch: str) -> str:
        bid = f"vba::{module_name}::{proc_name}::{kind}"
        if branch:
            bid += f"::{branch}"
        return bid

    # Build lookup: uppercase proc name → list of (module, proc_name, kind, branch)
    proc_lookup: dict[str, list[tuple[str, str, str, str]]] = {}
    for row in proc_rows:
        key = row[1].upper()  # procedure name
        if key not in _VBA_KEYWORDS:
            proc_lookup.setdefault(key, []).append((row[3], row[1], row[4], row[5] or ""))

    call_edges: list[tuple[str, str]] = []
    call_pattern = re.compile(r"\b([A-Za-z_]\w*)\s*[\(\s,]")

    for row in proc_rows:
        proc_name = row[1]
        body = row[2] or ""
        module_name = row[3]
        kind = row[4]
        compile_branch = row[5] or ""
        caller_id = _build_proc_id(module_name, proc_name, kind, compile_branch)

        clean_body = _strip_vba_comments_and_strings(body)

        for match in call_pattern.finditer(clean_body):
            called = match.group(1).upper()
            if called == proc_name.upper():
                continue  # skip self-reference
            if called in proc_lookup:
                # Prefer same module + matching compile_branch, then same module,
                # then first match. The call graph cannot perfectly disambiguate
                # without resolving #If at runtime, so this is best-effort.
                targets = proc_lookup[called]
                target = (
                    next(
                        (t for t in targets if t[0] == module_name and t[3] == compile_branch),
                        None,
                    )
                    or next(
                        (t for t in targets if t[0] == module_name),
                        None,
                    )
                    or targets[0]
                )
                callee_id = _build_proc_id(target[0], target[1], target[2], target[3])
                if (caller_id, callee_id) not in call_edges:
                    call_edges.append((caller_id, callee_id))

    if call_edges:
        conn.executemany(
            "INSERT OR IGNORE INTO binding_edges (from_binding_id, to_binding_id, edge_count) VALUES (?, ?, 1)",
            [(from_id, to_id) for from_id, to_id in call_edges],
        )
        metrics["call_graph_edges"] = len(call_edges)

    # --- 3. Static VBA reference extraction (Tier 1+2 patterns) ---
    try:
        from xl_marinade.core.vba.reference_extractor import extract_and_store_references

        ref_metrics = extract_and_store_references(conn)
        metrics["static_refs_extracted"] = ref_metrics.get("refs_extracted", 0)
        metrics["static_ref_edges"] = ref_metrics.get("edges_created", 0)
    except Exception as e:
        print(f"  Warning: VBA static reference extraction failed: {e}", file=sys.stderr)

    # --- 3b. R21 Fix B: VBA paste-edge synthesis ---
    # Emit `via_vba_paste` binding_edges for source-template ←→ output-block
    # dependencies that flow through PasteSpecial / .Value=.Value statements
    # in VBA. Runs after the formula-DAG and static-ref edges so source and
    # target bindings exist; the synthesiser only ADDS new edges, never
    # mutates existing ones.
    try:
        from xl_marinade.core.vba.paste_edges import synthesize_paste_edges

        paste_metrics = synthesize_paste_edges(conn)
        metrics["vba_paste_edges_inserted"] = paste_metrics.get("edges_inserted", 0)
        metrics["vba_paste_events_seen"] = paste_metrics.get("events_seen", 0)
        metrics["vba_paste_events_resolved"] = paste_metrics.get("events_resolved", 0)
        metrics["vba_paste_events_unresolved"] = paste_metrics.get("events_unresolved", 0)
    except Exception as e:
        print(f"  Warning: VBA paste-edge synthesis failed: {e}", file=sys.stderr)

    # --- 4. Compute VBA sheet affinity ---
    try:
        _compute_vba_sheet_affinity(conn)
    except Exception as e:
        print(f"  Warning: VBA sheet affinity computation failed: {e}", file=sys.stderr)

    # --- 5. LLM-assisted enrichment for dynamic references (Tier 3) ---
    # Opt-in only (BYOK): never fires unless the caller passes enrich=True, so the
    # default Tier-0 extraction path makes no network call even if a key is present.
    if enrich:
        try:
            from xl_marinade.llm.vba_enrichment import enrich_and_store
        except ImportError as e:
            raise ExtractionError(
                "VBA LLM enrichment (enrich=True) requires the optional add-on: "
                "pip install xl-marinade[llm]"
            ) from e
        try:
            llm_metrics = enrich_and_store(conn)
            metrics["llm_refs_inferred"] = llm_metrics.get("refs_inferred", 0)
            metrics["llm_edges_created"] = llm_metrics.get("edges_created", 0)
            metrics["llm_descriptions"] = llm_metrics.get("descriptions_generated", 0)
            metrics["llm_latency_s"] = llm_metrics.get("total_latency_s", 0.0)
        except Exception as e:
            print(f"  Warning: VBA LLM enrichment failed: {e}", file=sys.stderr)

    conn.commit()
    return metrics


def _compute_vba_sheet_affinity(conn: sqlite3.Connection) -> None:
    """Pre-compute sheet affinity for VBA procedure nodes.

    For each VBA procedure, counts edges to/from bindings on each sheet and
    determines the primary sheet (most connections). Stores in vba_sheet_affinity
    table for serve-time module assignment.
    """
    tables = {
        r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    }
    if "vba_procedures" not in tables or "binding_edges" not in tables:
        return

    # Create table if not exists
    conn.execute("""
        CREATE TABLE IF NOT EXISTS vba_sheet_affinity (
            procedure_binding_id TEXT PRIMARY KEY,
            primary_sheet TEXT,
            secondary_sheets_json TEXT
        )
    """)

    # Find all VBA binding IDs
    vba_edges = conn.execute("""
        SELECT from_binding_id, to_binding_id
        FROM binding_edges
        WHERE from_binding_id LIKE 'vba::%' OR to_binding_id LIKE 'vba::%'
    """).fetchall()

    if not vba_edges:
        conn.commit()
        return

    # Collect VBA binding IDs and their connected cell bindings
    vba_connections: dict[str, list[str]] = {}  # vba_id → [cell_binding_ids]
    for row in vba_edges:
        from_id, to_id = row[0], row[1]
        if from_id.startswith("vba::"):
            vba_connections.setdefault(from_id, []).append(to_id)
        if to_id.startswith("vba::"):
            vba_connections.setdefault(to_id, []).append(from_id)

    # For each VBA node, look up connected bindings' sheets
    if "agent_bindings" not in tables:
        conn.commit()
        return

    affinity_rows: list[tuple[str, str, str]] = []
    for vba_id, connected_ids in vba_connections.items():
        # Filter to non-VBA binding IDs
        cell_ids = [cid for cid in connected_ids if not cid.startswith("vba::")]
        if not cell_ids:
            continue

        # Count sheets
        sheet_counts: dict[str, int] = {}
        for cid in cell_ids:
            row = conn.execute(
                "SELECT sheet FROM agent_bindings WHERE binding_id = ?", (cid,)
            ).fetchone()
            if row:
                sheet = row[0]
                sheet_counts[sheet] = sheet_counts.get(sheet, 0) + 1

        if not sheet_counts:
            continue

        # Primary = most connections
        primary = max(sheet_counts, key=lambda s: sheet_counts[s])
        secondary = [s for s in sorted(sheet_counts) if s != primary]
        affinity_rows.append(
            (
                vba_id,
                primary,
                json.dumps(secondary) if secondary else "[]",
            )
        )

    if affinity_rows:
        conn.executemany(
            "INSERT OR REPLACE INTO vba_sheet_affinity (procedure_binding_id, primary_sheet, secondary_sheets_json) VALUES (?, ?, ?)",
            affinity_rows,
        )

    conn.commit()


def _extract_and_store_vba(workbook_path: Path, conn: sqlite3.Connection) -> dict:
    """
    Phase 2 VBA extraction stage: full grammar-based extraction of all procedures,
    declarations, event handlers, and security findings via xl_marinade.core.vba.

    Populates the ``udfs`` table for backward compatibility with Phase 1's
    FIND_UDF_FOR_FORMULA executor (which queries udfs by name). Phase 3 will add
    vba_modules, vba_procedures, vba_procedure_edges, and vba_chunks tables.

    Runs after loader.finalize() and before loader.drop_raw_tables().

    Args:
        workbook_path: Path to the .xlsm/.xlam/.xla file being extracted
        conn: Open SQLite connection to the build database

    Returns:
        Dict with extraction stats: {procedures, modules, declarations, events,
        security_findings, udfs_inserted, parse_errors}
    """
    stats = {
        "procedures": 0,
        "modules": 0,
        "declarations": 0,
        "events": 0,
        "security_findings": 0,
        "udfs_inserted": 0,
        "parse_errors": 0,
    }

    try:
        from xl_marinade.core.vba.extractor import extract_vba
    except ImportError:
        logger.exception("Failed to import xl_marinade.core.vba for VBA extraction")
        return stats

    try:
        extraction = extract_vba(workbook_path)
    except Exception:
        logger.exception("extract_vba raised for %s", workbook_path)
        return stats

    stats["modules"] = len(extraction.modules)
    stats["procedures"] = len(extraction.procedures)
    stats["declarations"] = len(extraction.declarations)
    stats["events"] = sum(1 for p in extraction.procedures if p.is_event_handler)
    stats["security_findings"] = len(extraction.security_findings)
    stats["parse_errors"] = len(extraction.parse_errors)

    if extraction.parse_errors:
        for err in extraction.parse_errors:
            logger.warning("VBA parse error in %s: %s", workbook_path.name, err.message)

    # Populate the udfs table with public Functions (backward compat for Phase 1).
    # Phase 3 will add a vba_procedures table covering all procedure kinds.
    public_functions = [p for p in extraction.procedures if p.kind == "function" and p.is_public]

    if public_functions:
        sorted_funcs = sorted(public_functions, key=lambda p: (p.name, p.module_name))
        try:
            with conn:
                conn.executemany(
                    """
                    INSERT OR IGNORE INTO udfs (
                        name, module, param_count, param_names_json,
                        declared_volatile, source_text, source_hash
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    [
                        (
                            p.name,
                            p.module_name,
                            p.param_count,
                            json.dumps(p.param_names, sort_keys=True),
                            "Application.Volatile" in p.body,
                            p.body,
                            p.normalized_body_hash,
                        )
                        for p in sorted_funcs
                    ],
                )
            stats["udfs_inserted"] = len(sorted_funcs)
        except sqlite3.Error:
            logger.exception("Failed to insert UDFs into udfs table for %s", workbook_path)

    # --- Phase 3: populate vba_modules, vba_procedures, vba_module_declarations ---
    try:
        _store_vba_phase3_tables(conn, extraction)
    except Exception:
        logger.exception("Failed to populate Phase 3 VBA tables for %s", workbook_path)

    return stats


def _store_vba_phase3_tables(conn: sqlite3.Connection, extraction) -> None:
    """Persist VBAExtraction data into the Phase 3 schema tables."""

    # 1. vba_modules
    module_id_map: dict[str, int] = {}  # module_name → module_id
    with conn:
        for mod in sorted(extraction.modules, key=lambda m: m.name):
            # Build security findings JSON for this module
            mod_security = [
                f
                for f in extraction.security_findings
                # Security findings from olevba are workbook-wide, not per-module.
                # Store the full set on each module for now; Phase 3+ can refine.
            ]
            cursor = conn.execute(
                """
                INSERT OR IGNORE INTO vba_modules (name, kind, source_sha256, source_text, security_findings_json)
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    mod.name,
                    mod.kind,
                    mod.source_sha256,
                    mod.source_text,
                    json.dumps(mod_security if mod == extraction.modules[0] else [], default=str),
                ),
            )
            if cursor.lastrowid:
                module_id_map[mod.name] = cursor.lastrowid
            else:
                row = conn.execute(
                    "SELECT module_id FROM vba_modules WHERE name = ?", (mod.name,)
                ).fetchone()
                if row:
                    module_id_map[mod.name] = row[0]

    # 2. vba_procedures
    proc_id_map: dict[str, int] = {}  # "module::name::kind::compile_branch" → procedure_id
    with conn:
        for proc in sorted(
            extraction.procedures,
            key=lambda p: (p.module_name, p.name, p.kind, p.compile_branch),
        ):
            mod_id = module_id_map.get(proc.module_name)
            if mod_id is None:
                continue
            trigger_json = json.dumps(proc.event_trigger_spec) if proc.event_trigger_spec else None
            params_json = json.dumps(proc.param_names, sort_keys=True)
            try:
                cursor = conn.execute(
                    """
                    INSERT OR IGNORE INTO vba_procedures (
                        module_id, name, kind, signature, parameters_json, return_type,
                        body, normalized_body_hash, is_public, is_event_handler,
                        event_trigger_spec_json, line_start, line_end, compile_branch
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        mod_id,
                        proc.name,
                        proc.kind,
                        proc.signature,
                        params_json,
                        proc.return_type,
                        proc.body,
                        proc.normalized_body_hash,
                        1 if proc.is_public else 0,
                        1 if proc.is_event_handler else 0,
                        trigger_json,
                        proc.line_start,
                        proc.line_end,
                        proc.compile_branch,
                    ),
                )
                key = f"{proc.module_name}::{proc.name}::{proc.kind}::{proc.compile_branch}"
                if cursor.lastrowid:
                    proc_id_map[key] = cursor.lastrowid
            except sqlite3.IntegrityError:
                pass  # True duplicate (rare)

    # 3. vba_module_declarations — removed (WI-5): populated but never read in production

    # 4. cell_udf_calls — scan formulas table for UDF name matches
    # This is the Phase 3 optimization of Phase 1's query-time regex scan.
    udf_rows = conn.execute("SELECT udf_id, name FROM udfs").fetchall()
    if udf_rows:
        import re

        udf_lookup = {name.upper(): udf_id for udf_id, name in udf_rows}
        udf_pattern = re.compile(
            r"\b(" + "|".join(re.escape(name) for name in udf_lookup) + r")\s*\(",
            re.IGNORECASE,
        )
        # Scan all formulas and match cells
        formula_matches = conn.execute("""
            SELECT c.cell_id, f.formula_r1c1
            FROM cells c
            JOIN formulas f ON c.formula_id = f.formula_id
            WHERE c.formula_id IS NOT NULL
        """).fetchall()

        cell_udf_pairs = []
        for cell_id, formula_r1c1 in formula_matches:
            if not formula_r1c1:
                continue
            for m in udf_pattern.finditer(formula_r1c1):
                matched_name = m.group(1).upper()
                if matched_name in udf_lookup:
                    cell_udf_pairs.append((cell_id, udf_lookup[matched_name]))

        if cell_udf_pairs:
            with conn:
                conn.executemany(
                    "INSERT OR IGNORE INTO cell_udf_calls (cell_id, udf_id) VALUES (?, ?)",
                    cell_udf_pairs,
                )

    # 5. vba_chunks — produce retrieval-ready chunks for the reasoning pipeline
    try:
        from xl_marinade.core.vba.chunker import chunk_extraction

        chunks = chunk_extraction(extraction)
        if chunks:
            with conn:
                conn.executemany(
                    """
                    INSERT OR IGNORE INTO vba_chunks (
                        procedure_id, chunk_index, chunk_text, line_start, line_end,
                        enclosing_block_kind, identifier_tokens_json, comment_tokens_json,
                        referenced_cells_json, called_procedures_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    [
                        (
                            # Resolve procedure_id from proc_id_map
                            proc_id_map.get(
                                f"{c.module_name}::{c.procedure_name}::{c.procedure_kind}::{c.procedure_compile_branch}",
                                None,
                            ),
                            c.chunk_index,
                            c.chunk_text,
                            c.line_start,
                            c.line_end,
                            c.enclosing_block_kind,
                            json.dumps(c.identifier_tokens),
                            json.dumps(c.comment_tokens),
                            json.dumps(c.referenced_cells) if c.referenced_cells else None,
                            json.dumps(c.called_procedures) if c.called_procedures else None,
                        )
                        for c in chunks
                        if proc_id_map.get(
                            f"{c.module_name}::{c.procedure_name}::{c.procedure_kind}::{c.procedure_compile_branch}"
                        )
                        is not None
                    ],
                )
    except Exception:
        logger.exception("Failed to populate vba_chunks")


# Exit codes (reserved by subsprint plan)
EXIT_SUCCESS = 0
EXIT_SLA_FAILURE = 42
EXIT_SQLITE_VERSION_FAILURE = 43

# Schema version
SCHEMA_VERSION = "3.0"
BUILD_MODE = "fast"


def get_git_sha() -> str:
    """Get current git commit SHA."""
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=5,
            check=True,
        )
        return result.stdout.strip()
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired, FileNotFoundError):
        return "unknown"


def compute_workbook_sha256(workbook_path: Path) -> str:
    """Compute SHA256 checksum of workbook file."""
    hasher = hashlib.sha256()
    with open(workbook_path, "rb") as f:
        while chunk := f.read(8192):
            hasher.update(chunk)
    return hasher.hexdigest()


def read_workbook_doc_title(workbook_path: Path) -> str | None:
    """Return the verbatim ``docProps/core.xml`` ``<dc:title>`` text if present.

    This is a deterministic provenance read of bytes already in the workbook
    (no fabrication). Returns the stripped title, or ``None`` when the part is
    absent, unparseable, or empty.
    """
    try:
        with zipfile.ZipFile(str(workbook_path), "r") as zf:
            try:
                raw = zf.read("docProps/core.xml")
            except KeyError:
                return None
    except (OSError, zipfile.BadZipFile):
        return None
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return None
    el = root.find("{http://purl.org/dc/elements/1.1/}title")
    if el is None or el.text is None:
        return None
    text = el.text.strip()
    return text or None


def _parse_root_cells(
    roots_data: dict[str, Any], sheet_name_map: dict[str, tuple[int, str]]
) -> list[tuple[int, int, int]]:
    """
    Parse root cells from roots JSON data.

    Args:
        roots_data: Roots JSON with user_root specification
        sheet_name_map: Map of sheet names (case-insensitive) to (sheet_id, canonical_name)

    Returns:
        List of (sheet_id, row, col) tuples for root cells

    Raises:
        ValueError: If roots_data is invalid or sheet not found
    """
    if "user_root" not in roots_data:
        raise ValueError("roots_data must contain 'user_root' key")

    user_root = roots_data["user_root"]
    sheet_name = user_root.get("sheet")
    range_str = user_root.get("range")

    if not sheet_name or not range_str:
        raise ValueError("user_root must contain 'sheet' and 'range' keys")

    # Resolve sheet name
    sheet_info = sheet_name_map.get(sheet_name.lower())
    if not sheet_info:
        raise ValueError(f"Sheet not found: {sheet_name}")

    sheet_id, _ = sheet_info

    # Parse range
    try:
        if ":" in range_str:
            # Range reference - expand to all cells
            parts = range_str.split(":")
            r1, c1, _, _ = parse_a1_ref(parts[0])
            r2, c2, _, _ = parse_a1_ref(parts[1])

            # Normalize order
            if r1 > r2:
                r1, r2 = r2, r1
            if c1 > c2:
                c1, c2 = c2, c1

            # Generate all cells in range
            roots = []
            for row in range(r1, r2 + 1):
                for col in range(c1, c2 + 1):
                    roots.append((sheet_id, row, col))

            return roots
        else:
            # Single cell
            row, col, _, _ = parse_a1_ref(range_str)
            return [(sheet_id, row, col)]
    except (ValueError, IndexError) as e:
        raise ValueError(f"Invalid range format '{range_str}': {e}")


class TraversalContext:
    """Shared mutable state for cell processing — used by both BFS and full-workbook modes.

    Encapsulates sheet caches, formula normalization, edge extraction, and semantic
    resolution into a single context object. The process_cell() method extracts a single
    cell and returns newly discovered cell IDs (for BFS queuing).
    """

    def __init__(
        self,
        workbook_path: str,
        sheets: list[tuple[int, str]],
        sheet_name_map: dict[str, tuple[int, str]],
        ref_extractor: ReferenceExtractor,
        resolution_engine: ResolutionEngine | None = None,
        index_chain: Any | None = None,
        resolution_context_cls: Any | None = None,
        expand_ranges_for_parity: bool = False,
        queue_all_range_cells: bool = False,
        range_traversal_max_cells: int | None = None,
        date_format_info: DateFormatInfo | None = None,
    ):
        self.workbook_path = workbook_path
        self.sheets = sheets
        self.sheet_name_map = sheet_name_map
        self.ref_extractor = ref_extractor
        self.resolution_engine = resolution_engine
        self.index_chain = index_chain
        self.resolution_context_cls = resolution_context_cls
        self.expand_ranges_for_parity = expand_ranges_for_parity
        self.queue_all_range_cells = queue_all_range_cells
        self.range_traversal_max_cells = range_traversal_max_cells

        self.sheet_id_to_name = {sheet_id: sheet_name for sheet_id, sheet_name in sheets}

        # Sheet-level caches (lazily populated)
        self.sheet_caches: dict[int, dict[tuple[int, int], tuple]] = {}
        self.sheet_shared_bases: dict[int, dict[int, tuple[str, int, int]]] = {}
        self.sheet_row_index: dict[int, dict[int, list[int]]] = {}
        self.sheet_rows_sorted: dict[int, list[int]] = {}
        self.sheet_cache_counts: dict[int, int] = {}
        self.total_cached_cells = 0
        self.load_sheet_times: dict[int, float] = {}

        # Performance counters
        self.perf_counters = {
            "normalize_formula_s": 0.0,
            "extract_edges_s": 0.0,
            "range_expand_s": 0.0,
            "range_edges_seen": 0,
            "range_edges_unique": 0,
            "range_scan_cells": 0,
            "range_expansion_skipped": 0,
            "range_expansion_skipped_full_workbook": 0,
        }
        self.seen_range_keys: set[tuple[int, int, int, int, int]] = set()
        self.expanded_range_keys: set[tuple[int, int, int, int, int]] = set()
        self.formula_cache = SheetFormulaCache(workbook_path) if expand_ranges_for_parity else None
        self.resolution_metrics: dict[str, dict[str, int]] = {}

        # Output accumulators
        self.cells_out: list[tuple] = []
        self.formulas_out: dict[str, tuple[str, str, int, int, int]] = {}
        self.value_blobs: dict[str, str] = {}
        self.edges_internal_out: list[tuple[int, int]] = []
        self.edges_range_out: list[tuple] = []
        self.edges_external_out: list[tuple[int, str]] = []
        self.total_cells = 0
        self.total_formulas = 0
        self.range_traversal_skipped = 0
        self.range_traversal_skipped_cells = 0

        self.range_edge_threshold = 10
        if range_traversal_max_cells is not None and range_traversal_max_cells <= 0:
            self.range_traversal_max_cells = None

        # Pre-compute format blobs per style index (date + number-format
        # signal flags). Blobs stay minimal: keys are only emitted when set.
        _dfi = date_format_info or DateFormatInfo()
        self._fmt_sha_by_style: dict[int, str] = {}
        self._fmt_blobs: dict[str, str] = {}
        _flags_by_style = getattr(_dfi, "format_flags_by_style", {}) or {}
        for _si in set(_dfi.date_style_indices) | set(_flags_by_style):
            _fmt_obj: dict[str, Any] = {}
            if _si in _dfi.date_style_indices:
                _fmt_obj["is_date"] = True
                if _dfi.is_1904:
                    _fmt_obj["date1904"] = True
            _fmt_obj.update(_flags_by_style.get(_si, {}))
            _fmt_json, _fmt_sha = canonicalize_and_hash(_fmt_obj)
            self._fmt_sha_by_style[_si] = _fmt_sha
            self._fmt_blobs[_fmt_sha] = _fmt_json

    def _record_resolution(self, function_name: str, result: ResolutionResult | None) -> None:
        if not result:
            return
        status = result.status or "unresolved"
        bucket = self.resolution_metrics.setdefault(function_name, {})
        bucket[status] = bucket.get(status, 0) + 1

    def load_sheet_cache(self, sheet_id: int) -> None:
        if sheet_id in self.sheet_caches:
            return
        start = time.perf_counter()
        sheet_name = self.sheet_id_to_name[sheet_id]
        print(f"  Loading sheet cache: {sheet_name}", file=sys.stderr)

        cache: dict[tuple[int, int], tuple] = {}
        shared_bases: dict[int, tuple[str, int, int]] = {}
        row_index: dict[int, list[int]] = {}
        for r, c, f, v, dt, shared_idx, shared_master, si in stream_worksheet_cells(
            self.workbook_path, sheet_name
        ):
            cache[(r, c)] = (f, None, v, dt, shared_idx, shared_master, si)
            if shared_idx is not None and shared_master and f:
                shared_bases[shared_idx] = (f, r, c)
            row_index.setdefault(r, []).append(c)

        self.sheet_caches[sheet_id] = cache
        self.sheet_shared_bases[sheet_id] = shared_bases
        for cols in row_index.values():
            cols.sort()
        self.sheet_row_index[sheet_id] = row_index
        self.sheet_rows_sorted[sheet_id] = sorted(row_index.keys())

        self.sheet_cache_counts[sheet_id] = len(cache)
        import threading

        if not hasattr(self, "_cache_lock"):
            self._cache_lock = threading.Lock()
        with self._cache_lock:
            self.total_cached_cells += len(cache)
        self.load_sheet_times[sheet_id] = time.perf_counter() - start

    def get_cell_data(
        self, sheet_id: int, row: int, col: int
    ) -> tuple[str | None, str | None, Any, str]:
        """Get cell data (formula_a1, formula_r1c1, value, data_type) for a specific cell."""
        self.load_sheet_cache(sheet_id)

        cached = self.sheet_caches[sheet_id].get(
            (row, col), (None, None, None, "blank", None, False, None)
        )
        formula_a1, formula_r1c1, value, data_type, shared_idx, shared_master = cached[:6]
        if formula_r1c1 is None:
            base_formula = formula_a1
            shared_base = None
            if not base_formula and shared_idx is not None:
                shared_base = self.sheet_shared_bases.get(sheet_id, {}).get(shared_idx)
                if shared_base:
                    base_formula, base_row, base_col = shared_base
            if base_formula:
                ctx = FormulaContext(
                    sheet_id=sheet_id, row=row, col=col, sheet_name=self.sheet_id_to_name[sheet_id]
                )
                if shared_idx is not None and not formula_a1 and shared_base:
                    formula_r1c1 = shift_shared_formula(
                        base_formula, base_row, base_col, row, col, ctx, self.sheet_name_map
                    )
                else:
                    formula_r1c1 = normalize_formula(base_formula, ctx, self.sheet_name_map)
                if not formula_a1:
                    if shared_base and (base_row != row or base_col != col):
                        formula_a1 = shift_shared_formula_a1(
                            base_formula, base_row, base_col, row, col
                        )
                    else:
                        formula_a1 = base_formula
                orig_si = cached[6] if len(cached) > 6 else None
                self.sheet_caches[sheet_id][(row, col)] = (
                    formula_a1,
                    formula_r1c1,
                    value,
                    data_type,
                    shared_idx,
                    shared_master,
                    orig_si,
                )

        return formula_a1, formula_r1c1, value, data_type

    def get_cell_style_index(self, sheet_id: int, row: int, col: int) -> int | None:
        """Return the style_index for a cached cell, or None."""
        self.load_sheet_cache(sheet_id)
        cached = self.sheet_caches[sheet_id].get((row, col))
        if cached is not None and len(cached) > 6:
            return cached[6]
        return None

    def _contains_lookup_function(self, ast: dict[str, Any]) -> bool:
        if not isinstance(ast, dict):
            return False
        if ast.get("type") == "Function":
            func_name = ast.get("name", "").upper()
            if func_name in LOOKUP_FUNCTIONS or func_name in ("CHOOSE", "ADDRESS"):
                return True
        for key in ("left", "right", "operand"):
            child = ast.get(key)
            if child and self._contains_lookup_function(child):
                return True
        return any(self._contains_lookup_function(arg) for arg in ast.get("args", []))

    def _contains_volatile_function(self, ast: dict[str, Any]) -> bool:
        if not isinstance(ast, dict):
            return False
        if ast.get("type") == "Function":
            func_name = ast.get("name", "").upper()
            if func_name in ("INDIRECT", "OFFSET"):
                return True
        for key in ("left", "right", "operand"):
            child = ast.get(key)
            if child and self._contains_volatile_function(child):
                return True
        return any(self._contains_volatile_function(arg) for arg in ast.get("args", []))

    def _collect_lookup_results(
        self, ast: dict[str, Any], cell_addr: str, out: list[ResolutionResult]
    ) -> None:
        """Resolve every lookup function in the AST (sibling-complete, subtree-opaque).

        When a resolvable lookup function is reached, its result is appended to
        ``out`` and recursion stops at that node — the resolution already consumes
        nested arguments (e.g. INDEX consumes its nested MATCH array), so recursing
        into them would double-count those drivers. Independent sibling lookups in
        other operands are still visited (no early return).
        """
        if not self.resolution_engine or not isinstance(ast, dict):
            return

        if ast.get("type") == "Function":
            func_name = ast.get("name", "").upper()
            current_sheet = cell_addr.split("!")[0] if "!" in cell_addr else ""

            if func_name == "VLOOKUP":
                result = self.resolution_engine.resolve_vlookup_semantic(
                    ast, current_sheet, cell_addr
                )
                self._record_resolution(func_name, result)
                out.append(result)
                return
            if func_name == "HLOOKUP":
                result = self.resolution_engine.resolve_hlookup_semantic(
                    ast, current_sheet, cell_addr
                )
                self._record_resolution(func_name, result)
                out.append(result)
                return
            if func_name == "INDEX":
                if self.index_chain and self.resolution_context_cls:
                    context = self.resolution_context_cls(
                        ast=ast,
                        workbook=self.resolution_engine.value_source,
                        current_sheet=current_sheet,
                        cell_address=cell_addr,
                        manual_provider=self.resolution_engine.manual_provider,
                    )
                    result = self.index_chain.resolve("INDEX", context)
                    self._record_resolution(func_name, result)
                    out.append(result)
                    return
                result = self.resolution_engine.resolve_index_semantic(
                    ast, current_sheet, cell_addr
                )
                self._record_resolution(func_name, result)
                out.append(result)
                return
            if func_name == "XLOOKUP":
                result = self.resolution_engine.resolve_xlookup_semantic(
                    ast, current_sheet, cell_addr
                )
                self._record_resolution(func_name, result)
                out.append(result)
                return
            if func_name == "MATCH":
                result = self.resolution_engine.resolve_match_semantic(
                    ast, current_sheet, cell_addr
                )
                self._record_resolution(func_name, result)
                out.append(result)
                return
            if func_name == "CHOOSE":
                result = self.resolution_engine.resolve_choose_semantic(
                    ast, current_sheet, cell_addr
                )
                self._record_resolution(func_name, result)
                out.append(result)
                return
            if func_name == "ADDRESS":
                result = self.resolution_engine.resolve_address_semantic(
                    ast, current_sheet, cell_addr
                )
                self._record_resolution(func_name, result)
                out.append(result)
                return

        # Recurse into child nodes (no early return — collect all siblings)
        for key in ("left", "right", "operand"):
            child = ast.get(key)
            if child:
                self._collect_lookup_results(child, cell_addr, out)
        for arg in ast.get("args", []):
            self._collect_lookup_results(arg, cell_addr, out)

    def _validate_byvalue_indirect(self, folded: str) -> str | None:
        """Fabrication-safety gate for a by-value-folded INDIRECT address string.

        Returns the address string only when it explicitly names a real,
        properly-quoted sheet. Rejects the failure modes a blank/uncached/typo
        selector produces, which would otherwise fabricate an edge:
        - empty/whitespace string, or no sheet qualifier (would fall back to the
          formula's own sheet);
        - an empty sheet token ("''!A:A", "!A:A") from a blank/uncached selector;
        - an unquoted sheet name containing a space (Excel #REF!);
        - a sheet name that is not a real sheet in this workbook.
        """
        folded = folded.strip()
        if "!" not in folded:
            return None
        sheet_raw, _, addr = folded.partition("!")
        sheet_raw = sheet_raw.strip()
        if not addr.strip() or not sheet_raw:
            return None
        quoted = len(sheet_raw) >= 2 and sheet_raw[0] == "'" and sheet_raw[-1] == "'"
        sheet_name = sheet_raw[1:-1].replace("''", "'") if quoted else sheet_raw
        if not sheet_name:
            return None
        if not quoted and " " in sheet_name:
            return None
        if sheet_name.lower() not in self.sheet_name_map:
            return None
        return folded

    def _resolve_indirect_node(
        self, ast: dict[str, Any], current_sheet: str
    ) -> ResolutionResult | None:
        """Resolve a single INDIRECT() AST node to a ResolutionResult.

        Handles both a literal-constant argument and the by-value (Issue #1) case
        where the argument is a computed string folded from cached cell values.
        Returns None only when INDIRECT has no arguments (caller should recurse).
        Shared by the INDIRECT branch and the OFFSET-over-INDIRECT base case.
        """
        args = ast.get("args", [])
        if len(args) < 1:
            return None
        ref_text_arg = args[0]
        if ref_text_arg.get("type") == "Const" and isinstance(ref_text_arg.get("value"), str):
            return self.resolution_engine.resolve_indirect(ref_text_arg.get("value"), current_sheet)
        # By-value resolution (Issue #1): the INDIRECT argument is a computed
        # string, e.g. "'"&$C11&"'!"&D$9. Fold it from the CACHED values of the
        # referenced cells (& / CONCATENATE / Ref are already handled by
        # _resolve_argument), then resolve the concrete address as a normal
        # INDIRECT. Fabrication-safety: _validate_byvalue_indirect requires an
        # explicit, real, properly-quoted sheet, so inactive ("-")/blank/typo
        # selectors emit no resolved edge. R1C1 mode (explicit a1=FALSE) is out of
        # scope: resolving an R1C1 string as A1 would fabricate a wrong edge (R1C1
        # "C5:C20" = whole columns E:T, not cells C5:C20), so skip it.
        a1_false = (
            len(args) >= 2
            and args[1].get("type") == "Const"
            and str(args[1].get("value")).strip().lower() in ("false", "0")
        )
        arg_result = self.resolution_engine._resolve_argument(ref_text_arg, current_sheet)
        validated = (
            self._validate_byvalue_indirect(arg_result.value)
            if not a1_false and arg_result.success and isinstance(arg_result.value, str)
            else None
        )
        if validated:
            result = self.resolution_engine.resolve_indirect(validated, current_sheet)
            if result.status == "resolved":
                result.volatile_kind = "address_resolved_from_cache"
            return result
        result = ResolutionResult()
        result.status = "context_dependent"
        result.volatile_kind = "address_computed"
        result.notes = "INDIRECT argument is not a literal constant"
        return result

    def _collect_volatile_results(
        self, ast: dict[str, Any], cell_addr: str, out: list[ResolutionResult]
    ) -> None:
        """Resolve every volatile function (INDIRECT/OFFSET) in the AST.

        Same sibling-complete, subtree-opaque contract as _collect_lookup_results:
        append a resolved volatile's result and stop recursing into its own args,
        but keep scanning independent siblings.
        """
        if not self.resolution_engine or not isinstance(ast, dict):
            return

        if ast.get("type") == "Function":
            func_name = ast.get("name", "").upper()
            current_sheet = cell_addr.split("!")[0] if "!" in cell_addr else ""

            if func_name == "INDIRECT":
                result = self._resolve_indirect_node(ast, current_sheet)
                if result is not None:
                    self._record_resolution(func_name, result)
                    out.append(result)
                    return
            if func_name == "OFFSET":
                args = ast.get("args", [])
                if len(args) >= 3:
                    base_ref = None
                    base_from_cache = False
                    base_arg = args[0]
                    if base_arg.get("type") == "Ref":
                        base_ref = base_arg.get("ref")
                    elif (
                        base_arg.get("type") == "Function"
                        and base_arg.get("name", "").upper() == "INDIRECT"
                    ):
                        # OFFSET over a by-value INDIRECT anchor (Issue #1), e.g.
                        # OFFSET(INDIRECT("'"&$C31&"'!A3"),0,0,79,27): resolve the
                        # INDIRECT to its concrete cell, then OFFSET expands from it.
                        # The inner INDIRECT also resolves on its own via recursion,
                        # so K15 ends up with both the anchor cell and the OFFSET range.
                        base_result = self._resolve_indirect_node(base_arg, current_sheet)
                        if base_result and base_result.status == "resolved":
                            base_ref = base_result.resolved_volatile_ref
                            # Snapshot-specific if the anchor came from a cached selector.
                            base_from_cache = (
                                base_result.volatile_kind == "address_resolved_from_cache"
                            )

                    row_offset = None
                    col_offset = None
                    row_result = self.resolution_engine._resolve_argument(args[1], current_sheet)
                    col_result = self.resolution_engine._resolve_argument(args[2], current_sheet)
                    if row_result.success:
                        with contextlib.suppress(ValueError, TypeError):
                            row_offset = int(row_result.value)
                    if col_result.success:
                        with contextlib.suppress(ValueError, TypeError):
                            col_offset = int(col_result.value)

                    height = None
                    width = None
                    if len(args) >= 4:
                        height_result = self.resolution_engine._resolve_argument(
                            args[3], current_sheet
                        )
                        if height_result.success:
                            with contextlib.suppress(ValueError, TypeError):
                                height = int(height_result.value)
                    if len(args) >= 5:
                        width_result = self.resolution_engine._resolve_argument(
                            args[4], current_sheet
                        )
                        if width_result.success:
                            with contextlib.suppress(ValueError, TypeError):
                                width = int(width_result.value)

                    if base_ref:
                        result = self.resolution_engine.resolve_offset(
                            base_ref, row_offset, col_offset, height, width, current_sheet
                        )
                        if base_from_cache and result.status == "resolved":
                            result.volatile_kind = "address_resolved_from_cache"
                        self._record_resolution(func_name, result)
                        out.append(result)
                        return

        for key in ("left", "right", "operand"):
            child = ast.get(key)
            if child:
                self._collect_volatile_results(child, cell_addr, out)
        for arg in ast.get("args", []):
            self._collect_volatile_results(arg, cell_addr, out)

    def _semantic_edges(
        self, formula_a1: str | None, cell_addr: str, ctx: FormulaContext
    ) -> list[Edge]:
        """Extract semantic edges from lookup/volatile functions."""
        if not self.resolution_engine or not formula_a1:
            return []

        upper_formula = formula_a1.upper()
        if not any(
            func in upper_formula
            for func in LOOKUP_FUNCTIONS.union(VOLATILE_FUNCTIONS).union({"CHOOSE", "ADDRESS"})
        ):
            return []

        try:
            ast = parse_formula(formula_a1)
        except Exception:
            return []

        semantic_refs: list[str] = []
        # Refs that come from an Issue #1 by-value cache resolution are
        # snapshot-specific; tag their edges so the binding graph can mark them
        # range_dynamic and consumers can caveat them (Issue #2 provenance).
        dynamic_refs: set[str] = set()

        try:
            if self._contains_lookup_function(ast):
                lookup_results: list[ResolutionResult] = []
                self._collect_lookup_results(ast, cell_addr, lookup_results)
                for result in lookup_results:
                    if result and result.status in (
                        "resolved",
                        "partial_resolved",
                        "conservative_fallback",
                    ):
                        if result.resolved_lookup_ref:
                            semantic_refs.append(result.resolved_lookup_ref)
                        semantic_refs.extend(result.lookup_drivers)

            if self._contains_volatile_function(ast):
                volatile_results: list[ResolutionResult] = []
                self._collect_volatile_results(ast, cell_addr, volatile_results)
                for result in volatile_results:
                    if result and result.status == "resolved":
                        is_dynamic = result.volatile_kind == "address_resolved_from_cache"
                        if result.resolved_volatile_ref:
                            semantic_refs.append(result.resolved_volatile_ref)
                            if is_dynamic:
                                dynamic_refs.add(result.resolved_volatile_ref)
                        for drv in result.volatile_drivers:
                            semantic_refs.append(drv)
                            if is_dynamic:
                                dynamic_refs.add(drv)
        except Exception:
            # Resolution engine can fail on unexpected formula patterns;
            # skip semantic edges rather than aborting extraction.
            pass

        edges: list[Edge] = []
        for ref in semantic_refs:
            if not ref:
                continue
            if "!" in ref:
                sheet_part, ref_part = ref.split("!", 1)
                sheet_part = sheet_part.strip("'").replace("''", "'")
                sheet_info = self.sheet_name_map.get(sheet_part.lower())
                if not sheet_info:
                    edges.append(Edge(type="external", external_ref=f"UNRESOLVED:{ref}"))
                    continue
                target_sheet_id = sheet_info[0]
            else:
                ref_part = ref
                target_sheet_id = ctx.sheet_id

            edge = self.ref_extractor._extract_range_edge(ref_part, target_sheet_id, ctx)
            if edge:
                if ref in dynamic_refs:
                    edge.provenance = "resolved_from_cache"
                edges.append(edge)

        return edges

    def process_cell(
        self,
        cell_id: int,
        sheet_id: int,
        row: int,
        col: int,
        *,
        visited: set[int] | None = None,
        queued: set[int] | None = None,
    ) -> list[int]:
        """Process a single cell: normalize formula, extract edges, record outputs.

        Returns a list of newly discovered cell_ids (for BFS mode to enqueue;
        full-workbook mode discards them).

        Args:
            cell_id: Packed cell identity
            sheet_id, row, col: Unpacked cell coordinates
            visited: Set of already-visited cell_ids (for BFS dedup of edge targets)
            queued: Set of already-queued cell_ids (for BFS dedup of edge targets)
        """
        sheet_name = self.sheet_id_to_name[sheet_id]
        a1 = col_to_a1(col) + str(row)

        # Get cell data
        formula, cached_r1c1, value, data_type = self.get_cell_data(sheet_id, row, col)

        # Normalize formula if present
        formula_r1c1 = cached_r1c1
        if not formula_r1c1 and formula:
            ctx = FormulaContext(sheet_id=sheet_id, row=row, col=col, sheet_name=sheet_name)
            t0 = time.perf_counter()
            formula_r1c1 = normalize_formula(formula, ctx, self.sheet_name_map)
            self.perf_counters["normalize_formula_s"] += time.perf_counter() - t0

        has_value = value is not None and value != ""
        if not formula_r1c1 and not has_value:
            return []

        if formula_r1c1:
            self.total_formulas += 1
            candidate = (formula_r1c1, formula or "", sheet_id, row, col)
            existing = self.formulas_out.get(formula_r1c1)
            if existing is None:
                self.formulas_out[formula_r1c1] = candidate
            else:
                existing_key = (existing[3], existing[4], existing[2], existing[1])
                candidate_key = (candidate[3], candidate[4], candidate[2], candidate[1])
                if candidate_key < existing_key:
                    self.formulas_out[formula_r1c1] = candidate

        # Compute value SHA256 if value present
        value_sha256 = None
        if has_value:
            value_json, value_sha256 = canonicalize_and_hash(value)
            if value_sha256 not in self.value_blobs:
                self.value_blobs[value_sha256] = value_json

        # Determine format SHA for cells with format metadata (date / flags)
        format_sha256 = None
        if self._fmt_sha_by_style:
            si = self.get_cell_style_index(sheet_id, row, col)
            if si is not None and si in self._fmt_sha_by_style:
                format_sha256 = self._fmt_sha_by_style[si]
                if format_sha256 not in self.value_blobs:
                    self.value_blobs[format_sha256] = self._fmt_blobs[format_sha256]

        # Add cell to output
        self.cells_out.append(
            (
                cell_id,
                sheet_id,
                row,
                col,
                a1,
                formula_r1c1,
                formula or "",
                value_sha256,
                format_sha256,
                data_type,
                0,
                0,
                None,
            )
        )
        self.total_cells += 1

        # Extract edges if formula present
        new_cell_ids: list[int] = []
        if formula_r1c1:
            ctx = FormulaContext(sheet_id=sheet_id, row=row, col=col, sheet_name=sheet_name)

            formula_for_edges = formula_r1c1
            if self.expand_ranges_for_parity and self.formula_cache:
                formula_for_edges = (
                    self.formula_cache.get_formula(f"{sheet_name}!{a1}") or formula or ""
                )
                if formula_for_edges.startswith("="):
                    formula_for_edges = formula_for_edges[1:]
                t0 = time.perf_counter()
                edges = self.ref_extractor.extract_edges(formula_for_edges, ctx)
                self.perf_counters["extract_edges_s"] += time.perf_counter() - t0
            else:
                t0 = time.perf_counter()
                edges = self.ref_extractor.extract_edges(formula_r1c1, ctx)
                self.perf_counters["extract_edges_s"] += time.perf_counter() - t0
            semantic_edges = self._semantic_edges(
                formula_for_edges if self.expand_ranges_for_parity else formula,
                f"{sheet_name}!{a1}",
                ctx,
            )
            if semantic_edges:
                edges.extend(semantic_edges)

            # Local dedup: track edges emitted from this cell
            emitted_internal: set[int] = set()
            emitted_range: set[tuple] = set()
            emitted_external: set[str] = set()

            for edge in edges:
                if edge.type == "internal":
                    queue_target = True
                    if self.expand_ranges_for_parity:
                        target_sheet_id, target_row, target_col = unpack_cell_id(edge.to_cell_id)
                        target_formula, _, _, _ = self.get_cell_data(
                            target_sheet_id, target_row, target_col
                        )
                        queue_target = bool(target_formula)

                    if edge.to_cell_id not in emitted_internal:
                        self.edges_internal_out.append((cell_id, edge.to_cell_id))
                        emitted_internal.add(edge.to_cell_id)

                    if queue_target:
                        if visited is not None and queued is not None:
                            if edge.to_cell_id not in visited and edge.to_cell_id not in queued:
                                new_cell_ids.append(edge.to_cell_id)
                        else:
                            new_cell_ids.append(edge.to_cell_id)

                elif edge.type == "range":
                    self.perf_counters["range_edges_seen"] += 1
                    if self.expand_ranges_for_parity and edge.is_implicit_full_range:
                        continue
                    range_key = (edge.to_sheet_id, edge.to_r1, edge.to_c1, edge.to_r2, edge.to_c2)

                    range_cell_count = edge.cell_count
                    if range_cell_count is None:
                        range_cell_count = (edge.to_r2 - edge.to_r1 + 1) * (
                            edge.to_c2 - edge.to_c1 + 1
                        )
                    store_range_edge = True
                    if (
                        self.expand_ranges_for_parity
                        and range_cell_count < self.range_edge_threshold
                    ):
                        store_range_edge = False

                    if range_key not in self.seen_range_keys:
                        self.seen_range_keys.add(range_key)
                        self.perf_counters["range_edges_unique"] += 1

                    if store_range_edge and range_key not in emitted_range:
                        self.edges_range_out.append(
                            (
                                cell_id,
                                edge.to_sheet_id,
                                edge.to_r1,
                                edge.to_c1,
                                edge.to_r2,
                                edge.to_c2,
                                edge.to_range_a1,
                                range_cell_count,
                                edge.provenance,
                            )
                        )
                        emitted_range.add(range_key)

                    emit_expanded_edges = (
                        self.expand_ranges_for_parity
                        and range_cell_count < self.range_edge_threshold
                    )
                    should_traverse_range = True
                    if (
                        self.range_traversal_max_cells is not None
                        and not self.expand_ranges_for_parity
                        and not self.queue_all_range_cells
                        and range_cell_count >= self.range_traversal_max_cells
                    ):
                        should_traverse_range = False
                        self.range_traversal_skipped += 1
                        self.range_traversal_skipped_cells += range_cell_count

                    if not should_traverse_range:
                        continue
                    if not self.expand_ranges_for_parity and range_key in self.expanded_range_keys:
                        self.perf_counters["range_expansion_skipped"] += 1
                        continue

                    # Full-workbook mode: every populated cell is visited via the outer
                    # sheet loop, so sparse expansion here is redundant work. Skipping
                    # also avoids pre-loading other sheets' caches (breaking the per-sheet
                    # eviction invariant that keeps RSS bounded). BFS mode still needs
                    # this to discover reachable cells.
                    if visited is None:
                        self.perf_counters["range_expansion_skipped_full_workbook"] += 1
                        self.expanded_range_keys.add(range_key)
                        continue

                    # SPARSE EXPANSION: Queue cells from range that exist in sheet cache
                    self.load_sheet_cache(edge.to_sheet_id)

                    cache = self.sheet_caches[edge.to_sheet_id]
                    row_index = self.sheet_row_index.get(edge.to_sheet_id, {})
                    rows_sorted = self.sheet_rows_sorted.get(edge.to_sheet_id, [])
                    start_idx = bisect.bisect_left(rows_sorted, edge.to_r1)
                    t0 = time.perf_counter()
                    for r in rows_sorted[start_idx:]:
                        if r > edge.to_r2:
                            break
                        cols = row_index.get(r)
                        if not cols:
                            continue
                        col_start_idx = bisect.bisect_left(cols, edge.to_c1)
                        for c in cols[col_start_idx:]:
                            if c > edge.to_c2:
                                break
                            self.perf_counters["range_scan_cells"] += 1

                            cell_data = cache[(r, c)]
                            has_formula = bool(cell_data[0] or cell_data[4] is not None)
                            is_cross_sheet = edge.to_sheet_id != sheet_id
                            should_queue = (
                                has_formula or self.queue_all_range_cells or is_cross_sheet
                            )
                            if not should_queue and not emit_expanded_edges:
                                continue

                            target_cell_id = None
                            if emit_expanded_edges and has_formula:
                                target_cell_id = pack_cell_id(edge.to_sheet_id, r, c)
                                if target_cell_id not in emitted_internal:
                                    self.edges_internal_out.append((cell_id, target_cell_id))
                                    emitted_internal.add(target_cell_id)

                            if should_queue:
                                if target_cell_id is None:
                                    target_cell_id = pack_cell_id(edge.to_sheet_id, r, c)
                                if visited is not None and queued is not None:
                                    if (
                                        target_cell_id not in visited
                                        and target_cell_id not in queued
                                    ):
                                        new_cell_ids.append(target_cell_id)
                                else:
                                    new_cell_ids.append(target_cell_id)
                    self.perf_counters["range_expand_s"] += time.perf_counter() - t0
                    self.expanded_range_keys.add(range_key)

                elif edge.type == "external":
                    if edge.external_ref not in emitted_external:
                        self.edges_external_out.append((cell_id, edge.external_ref))
                        emitted_external.add(edge.external_ref)

        return new_cell_ids

    def to_output_dict(
        self,
        traversal_time_s: float = 0.0,
        peak_rss_mb: float = 0.0,
        telemetry_samples: list[dict] | None = None,
    ) -> dict[str, Any]:
        """Build the standard output dict matching _execute_root_traversal's return format."""
        if self.formula_cache:
            self.formula_cache.close()

        cache_stats = {
            "sheets_loaded": len(self.sheet_cache_counts),
            "total_cached_cells": self.total_cached_cells,
            "cached_cells_per_sheet": {
                self.sheet_id_to_name[sheet_id]: count
                for sheet_id, count in sorted(self.sheet_cache_counts.items())
            },
            "sheet_load_seconds": {
                self.sheet_id_to_name[sheet_id]: round(self.load_sheet_times.get(sheet_id, 0.0), 3)
                for sheet_id in sorted(self.sheet_cache_counts)
            },
        }

        return {
            "cells": self.cells_out,
            "formulas": [self.formulas_out[k] for k in sorted(self.formulas_out)],
            "json_blobs": list(self.value_blobs.items()),
            "edges_internal": self.edges_internal_out,
            "edges_range": self.edges_range_out,
            "edges_external": self.edges_external_out,
            "total_cells": self.total_cells,
            "total_formulas": self.total_formulas,
            "resolution_metrics": self.resolution_metrics,
            "telemetry": {
                "traversal_time_s": round(traversal_time_s, 3),
                "peak_rss_mb": round(peak_rss_mb, 2) if peak_rss_mb else None,
                "cache_stats": cache_stats,
                "range_traversal": {
                    "max_cells": self.range_traversal_max_cells,
                    "skipped_ranges": self.range_traversal_skipped,
                    "skipped_cells": self.range_traversal_skipped_cells,
                },
                "perf_counters": {
                    "normalize_formula_s": round(self.perf_counters["normalize_formula_s"], 3),
                    "extract_edges_s": round(self.perf_counters["extract_edges_s"], 3),
                    "range_expand_s": round(self.perf_counters["range_expand_s"], 3),
                    "range_edges_seen": self.perf_counters["range_edges_seen"],
                    "range_edges_unique": self.perf_counters["range_edges_unique"],
                    "range_scan_cells": self.perf_counters["range_scan_cells"],
                    "range_expansion_skipped": self.perf_counters["range_expansion_skipped"],
                    "range_expansion_skipped_full_workbook": self.perf_counters.get(
                        "range_expansion_skipped_full_workbook", 0
                    ),
                },
                "samples": telemetry_samples or [],
                "per_sheet": getattr(self, "_per_sheet_telemetry", []),
            },
        }


def _execute_full_workbook_extraction(
    workbook_path: str,
    sheets: list[tuple[int, str]],
    sheet_name_map: dict[str, tuple[int, str]],
    ref_extractor: ReferenceExtractor,
    resolution_engine: ResolutionEngine | None = None,
    index_chain: Any | None = None,
    resolution_context_cls: Any | None = None,
    date_format_info: DateFormatInfo | None = None,
) -> dict[str, Any]:
    """
    Process all cells on all sheets without BFS traversal.

    Unlike _execute_root_traversal which walks backwards from roots through
    precedents, this function visits every populated cell on every sheet.

    Returns the same output format as _execute_root_traversal.
    """
    ctx = TraversalContext(
        workbook_path=workbook_path,
        sheets=sheets,
        sheet_name_map=sheet_name_map,
        ref_extractor=ref_extractor,
        resolution_engine=resolution_engine,
        index_chain=index_chain,
        resolution_context_cls=resolution_context_cls,
        date_format_info=date_format_info,
    )

    traversal_start = time.perf_counter()
    process = psutil.Process() if HAS_PSUTIL else None
    peak_rss_mb = 0.0
    telemetry_samples: list[dict[str, Any]] = []
    iteration = 0
    last_report_time = traversal_start

    print(f"  Starting full-workbook extraction across {len(sheets)} sheets...", file=sys.stderr)

    # In full-workbook mode, load one sheet at a time and evict after processing.
    # This caps memory to ~1 sheet cache instead of all sheets simultaneously,
    # preventing swap thrashing on memory-constrained machines (e.g., 16GB for
    # a large workbook's 2.3M cells across 18 sheets).

    per_sheet_telemetry: list[dict[str, Any]] = []
    import gc as _gc

    for sheet_idx, (sheet_id, sheet_name) in enumerate(sheets):
        sheet_start = time.perf_counter()
        rss_before_sheet = None
        if process:
            try:
                rss_before_sheet = process.memory_info().rss / (1024 * 1024)
            except Exception:
                pass

        load_start = time.perf_counter()
        ctx.load_sheet_cache(sheet_id)
        cache = ctx.sheet_caches[sheet_id]
        load_elapsed = time.perf_counter() - load_start

        rss_after_load = None
        if process:
            try:
                rss_after_load = process.memory_info().rss / (1024 * 1024)
            except Exception:
                pass

        sheet_cell_count = len(cache)
        sheet_iter_start = iteration
        process_start = time.perf_counter()

        print(
            f"  [sheet {sheet_idx + 1}/{len(sheets)}] {sheet_name}: "
            f"{sheet_cell_count:,} cells loaded in {load_elapsed:.2f}s "
            f"(rss {rss_before_sheet:.0f}->{rss_after_load:.0f} MB)"
            if rss_before_sheet is not None
            else f"  [sheet {sheet_idx + 1}/{len(sheets)}] {sheet_name}: "
            f"{sheet_cell_count:,} cells loaded in {load_elapsed:.2f}s",
            file=sys.stderr,
            flush=True,
        )

        for row, col in sorted(cache.keys()):
            cell_id = pack_cell_id(sheet_id, row, col)
            iteration += 1

            # Process cell (discard returned new_cell_ids — no BFS queue)
            ctx.process_cell(cell_id, sheet_id, row, col)

            # Telemetry sampling every 1000 iterations
            if iteration % 1000 == 0:
                now = time.perf_counter()
                elapsed_total = now - traversal_start
                elapsed_interval = now - last_report_time
                rate_overall = iteration / elapsed_total if elapsed_total > 0 else 0.0
                rate_interval = 1000 / elapsed_interval if elapsed_interval > 0 else 0.0

                rss_mb = None
                if process:
                    try:
                        rss_mb = process.memory_info().rss / (1024 * 1024)
                        if rss_mb > peak_rss_mb:
                            peak_rss_mb = rss_mb
                    except Exception:
                        rss_mb = None

                print(
                    f"    Iteration {iteration}: {ctx.total_cells} cells, "
                    f"rate={rate_overall:.2f}/s, rss={rss_mb:.1f}MB"
                    if rss_mb
                    else f"    Iteration {iteration}: {ctx.total_cells} cells, "
                    f"rate={rate_overall:.2f}/s",
                    file=sys.stderr,
                )

                telemetry_samples.append(
                    {
                        "iteration": iteration,
                        "cells_out": ctx.total_cells,
                        "elapsed_s": round(elapsed_total, 3),
                        "rate_overall": round(rate_overall, 3),
                        "rate_interval": round(rate_interval, 3),
                        "rss_mb": round(rss_mb, 2) if rss_mb is not None else None,
                        "peak_rss_mb": round(peak_rss_mb, 2) if peak_rss_mb else None,
                    }
                )
                last_report_time = now

        process_elapsed = time.perf_counter() - process_start
        sheet_cells_processed = iteration - sheet_iter_start
        sheet_rate = sheet_cells_processed / process_elapsed if process_elapsed > 0 else 0.0

        # Evict sheet cache after processing to bound memory usage.
        # In full-workbook mode, each sheet is processed once and never revisited.
        del ctx.sheet_caches[sheet_id]
        if sheet_id in ctx.sheet_shared_bases:
            del ctx.sheet_shared_bases[sheet_id]
        if sheet_id in ctx.sheet_row_index:
            del ctx.sheet_row_index[sheet_id]
        if sheet_id in ctx.sheet_rows_sorted:
            del ctx.sheet_rows_sorted[sheet_id]
        _gc.collect()

        rss_after_evict = None
        if process:
            try:
                rss_after_evict = process.memory_info().rss / (1024 * 1024)
                if rss_after_evict > peak_rss_mb:
                    peak_rss_mb = rss_after_evict
            except Exception:
                pass

        sheet_total_elapsed = time.perf_counter() - sheet_start
        print(
            f"  [sheet {sheet_idx + 1}/{len(sheets)}] {sheet_name}: "
            f"processed {sheet_cells_processed:,} cells in {process_elapsed:.2f}s "
            f"({sheet_rate:,.0f} cells/s), total {sheet_total_elapsed:.2f}s, "
            f"rss after evict={rss_after_evict:.0f} MB, cells_out={ctx.total_cells:,}, "
            f"edges_internal={len(ctx.edges_internal_out):,}, "
            f"edges_range={len(ctx.edges_range_out):,}, "
            f"value_blobs={len(ctx.value_blobs):,}"
            if rss_after_evict is not None
            else f"  [sheet {sheet_idx + 1}/{len(sheets)}] {sheet_name}: "
            f"processed {sheet_cells_processed:,} cells in {process_elapsed:.2f}s "
            f"({sheet_rate:,.0f} cells/s)",
            file=sys.stderr,
            flush=True,
        )

        per_sheet_telemetry.append(
            {
                "sheet_idx": sheet_idx,
                "sheet_name": sheet_name,
                "sheet_id": sheet_id,
                "cells_in_sheet": sheet_cell_count,
                "cells_processed": sheet_cells_processed,
                "load_time_s": round(load_elapsed, 3),
                "process_time_s": round(process_elapsed, 3),
                "total_time_s": round(sheet_total_elapsed, 3),
                "rate_cells_per_s": round(sheet_rate, 1),
                "rss_before_mb": round(rss_before_sheet, 1) if rss_before_sheet else None,
                "rss_after_load_mb": round(rss_after_load, 1) if rss_after_load else None,
                "rss_after_evict_mb": round(rss_after_evict, 1) if rss_after_evict else None,
                "cumulative_cells_out": ctx.total_cells,
                "cumulative_edges_internal": len(ctx.edges_internal_out),
                "cumulative_edges_range": len(ctx.edges_range_out),
                "cumulative_value_blobs": len(ctx.value_blobs),
            }
        )

    traversal_time_s = time.perf_counter() - traversal_start
    print(
        f"  Full-workbook extraction complete: {ctx.total_cells} cells, "
        f"{ctx.total_formulas} formulas ({traversal_time_s:.2f}s, "
        f"peak_rss={peak_rss_mb:.1f}MB)",
        file=sys.stderr,
    )

    ctx._per_sheet_telemetry = per_sheet_telemetry
    return ctx.to_output_dict(traversal_time_s, peak_rss_mb, telemetry_samples)


def _execute_root_traversal(
    workbook_path: str,
    root_cells: list[tuple[int, int, int]],
    sheets: list[tuple[int, str]],
    sheet_name_map: dict[str, tuple[int, str]],
    ref_extractor: ReferenceExtractor,
    resolution_engine: ResolutionEngine | None = None,
    index_chain: Any | None = None,
    resolution_context_cls: Any | None = None,
    expand_ranges_for_parity: bool = False,
    queue_all_range_cells: bool = False,
    range_traversal_max_cells: int | None = None,
    date_format_info: DateFormatInfo | None = None,
) -> dict[str, Any]:
    """
    Execute BFS traversal from root cells.

    Walks precedents from roots, extracting formulas and edges on-demand.
    Only reachable cells are included in the output.

    Delegates per-cell processing to TraversalContext.process_cell().
    """
    ctx = TraversalContext(
        workbook_path=workbook_path,
        sheets=sheets,
        sheet_name_map=sheet_name_map,
        ref_extractor=ref_extractor,
        resolution_engine=resolution_engine,
        index_chain=index_chain,
        resolution_context_cls=resolution_context_cls,
        expand_ranges_for_parity=expand_ranges_for_parity,
        queue_all_range_cells=queue_all_range_cells,
        range_traversal_max_cells=range_traversal_max_cells,
        date_format_info=date_format_info,
    )

    # BFS traversal state
    visited: set[int] = set()
    queued: set[int] = set()
    queue: deque[int] = deque()

    # Seed queue with roots
    for sheet_id, row, col in root_cells:
        cell_id = pack_cell_id(sheet_id, row, col)
        queue.append(cell_id)
        queued.add(cell_id)

    print(f"  Starting BFS from {len(root_cells)} root cells...", file=sys.stderr)

    iteration = 0
    traversal_start = time.perf_counter()
    last_report_time = traversal_start
    process = psutil.Process() if HAS_PSUTIL else None
    peak_rss_mb = 0.0
    telemetry_samples: list[dict[str, Any]] = []

    while queue:
        iteration += 1

        if iteration % 1000 == 0:
            now = time.perf_counter()
            elapsed_total = now - traversal_start
            elapsed_interval = now - last_report_time
            rate_overall = len(visited) / elapsed_total if elapsed_total > 0 else 0.0
            rate_interval = 1000 / elapsed_interval if elapsed_interval > 0 else 0.0
            queue_ratio = len(queue) / len(visited) if len(visited) > 0 else 0.0

            rss_mb = None
            if process:
                try:
                    rss_mb = process.memory_info().rss / (1024 * 1024)
                    if rss_mb > peak_rss_mb:
                        peak_rss_mb = rss_mb
                except Exception:
                    rss_mb = None

            print(
                "    Iteration {iteration}: {visited} visited, {queued} queued, "
                "rate={rate:.2f}/s (interval={interval_rate:.2f}/s), "
                "queue_ratio={queue_ratio:.2f}x, rss={rss}MB".format(
                    iteration=iteration,
                    visited=len(visited),
                    queued=len(queue),
                    rate=rate_overall,
                    interval_rate=rate_interval,
                    queue_ratio=queue_ratio,
                    rss=f"{rss_mb:.1f}" if rss_mb is not None else "n/a",
                ),
                file=sys.stderr,
            )

            telemetry_samples.append(
                {
                    "iteration": iteration,
                    "visited": len(visited),
                    "queued": len(queue),
                    "elapsed_s": round(elapsed_total, 3),
                    "rate_overall": round(rate_overall, 3),
                    "rate_interval": round(rate_interval, 3),
                    "queue_ratio": round(queue_ratio, 3),
                    "rss_mb": round(rss_mb, 2) if rss_mb is not None else None,
                    "peak_rss_mb": round(peak_rss_mb, 2) if peak_rss_mb else None,
                }
            )

            last_report_time = now

        cell_id = queue.popleft()
        queued.discard(cell_id)

        if cell_id in visited:
            continue

        visited.add(cell_id)

        # Unpack and process cell via TraversalContext
        sheet_id, row, col = unpack_cell_id(cell_id)
        new_cell_ids = ctx.process_cell(
            cell_id,
            sheet_id,
            row,
            col,
            visited=visited,
            queued=queued,
        )

        # Enqueue newly discovered cells
        for new_id in new_cell_ids:
            queue.append(new_id)
            queued.add(new_id)

    traversal_time_s = time.perf_counter() - traversal_start
    print(
        f"  Traversal complete: {ctx.total_cells} cells, {ctx.total_formulas} formulas "
        f"({traversal_time_s:.2f}s, peak_rss={peak_rss_mb:.1f}MB)",
        file=sys.stderr,
    )

    # Fail loudly if a rooted traversal produced nothing. The alternative — silently
    # writing an empty DB — marks the extraction as 'ready' and breaks every
    # downstream concept resolution. We must reject this outcome.
    if root_cells and ctx.total_cells == 0:
        sheet_id_to_name = {sheet_id: sheet_name for sheet_id, sheet_name in sheets}
        root_descs: list[str] = []
        for sheet_id, row, col in root_cells:
            sheet_name = sheet_id_to_name.get(sheet_id, f"sheet_id={sheet_id}")
            root_descs.append(f"{sheet_name}!{col_to_a1(col)}{row}")
        available_sheets = [sheet_name for _, sheet_name in sheets]
        raise EmptyRootExtractionError(root_descs, available_sheets)

    return ctx.to_output_dict(traversal_time_s, peak_rss_mb, telemetry_samples)


def run_fast_extraction(
    workbook_path: Path,
    roots_data: dict[str, Any],
    output_db: Path,
    max_memory_mb: int = 1800,
    batch_size: int = 10_000,
    allow_non_canonical: bool = False,
    expand_ranges_for_parity: bool = False,
    queue_all_range_cells: bool = False,
    range_traversal_max_cells: int | None = None,
    enrich: bool = False,
) -> dict[str, Any]:
    """
    Run fast memory-efficient extraction pipeline with root-based traversal.

    Implements precedents-based traversal from roots per IR creation spec §1-§6.
    Only reachable cells (from roots via precedents) are extracted.

    Args:
        workbook_path: Path to Excel workbook (.xlsx/.xlsm)
        roots_data: Roots JSON data with user_root specification
        output_db: Path to output SQLite database
        max_memory_mb: Maximum memory budget in MB (default: 1800)
        batch_size: Batch size for executemany (default: 10,000)
        allow_non_canonical: Allow non-canonical builds (for testing only)
        expand_ranges_for_parity: If True, expand range edges into cell edges for parity runs
        queue_all_range_cells: If True, queue all populated cells in ranges (not just formula cells)

    Returns:
        Telemetry dictionary with performance metrics

    Raises:
        RuntimeError: If SQLite version < 3.27.0 (VACUUM INTO required)
    """
    # Step 1: Verify SQLite version
    version = sqlite3.sqlite_version_info
    if version < (3, 27, 0):
        print(
            f"Error: SQLite 3.27.0+ required for fast extraction, found {sqlite3.sqlite_version}",
            file=sys.stderr,
        )
        raise ExtractionError(f"SQLite 3.27.0+ required, found {sqlite3.sqlite_version}")

    run_start = time.perf_counter()
    print("Fast extraction starting...", file=sys.stderr)
    print(f"  Workbook: {workbook_path.name}", file=sys.stderr)
    print(f"  Output: {output_db}", file=sys.stderr)
    print(f"  Memory budget: {max_memory_mb} MB", file=sys.stderr)

    # Get git SHA and workbook checksum
    git_sha = get_git_sha()
    workbook_sha256 = compute_workbook_sha256(workbook_path)

    # Create build DB path (temporary)
    build_db_path = output_db.parent / f"{output_db.stem}_build.db"

    # Load schema
    schema_path = Path(__file__).parent / "schema.sql"
    with open(schema_path, encoding="utf-8") as f:
        schema_sql = f.read()

    try:
        # Step 2-3: Initialize bulk loader (opens connection, applies PRAGMAs, creates schema)
        memory_config = MemoryBudgetConfig(
            max_memory_mb=max_memory_mb, warn_threshold_pct=0.8, check_interval_rows=10_000
        )

        with BulkLoader(
            build_db_path=str(build_db_path),
            batch_size=batch_size,
            allow_non_canonical=allow_non_canonical,
            memory_budget_config=memory_config,
        ) as loader:
            loader.create_schema(schema_sql)
            print("  Schema created", file=sys.stderr)

            # Step 4: Load workbook data
            print("\nExtracting workbook metadata...", file=sys.stderr)

            # Extract sheet catalog
            catalog = WorkbookCatalog(str(workbook_path))
            # sheets is list of (sheet_id, sheet_name, rel_id, target) tuples
            sheets_full = catalog.sheets
            # Extract just (sheet_id, sheet_name) for loading
            sheets = [(sheet_id, sheet_name) for sheet_id, sheet_name, _, _ in sheets_full]

            print(f"  Sheets found: {len(sheets)}", file=sys.stderr)

            # Load sheets into DB
            loader.load_sheets(sheets)

            # Build sheet name map for reference extractor
            sheet_name_map = {
                sheet_name.lower(): (sheet_id, sheet_name) for sheet_id, sheet_name in sheets
            }

            # WI-11: Open workbook once, share across pipeline stages
            # (NameTableMap, resolution engine, grouping evidence extraction)
            shared_wb = None
            name_table_map = None
            try:
                shared_wb = LazyWorkbook(workbook_path, data_only=False, keep_vba=False)
                shared_wb.__enter__()
                name_table_map = NameTableMap(shared_wb)
            except Exception as e:
                print(f"Warning: Failed to build name/table map: {e}", file=sys.stderr)

            resolution_engine = ResolutionEngine(
                LazyValueFetcher(workbook_path, max_cached_sheets=None)
            )
            index_chain = None
            resolution_context_cls = None
            try:
                index_chain = create_index_resolution_chain(
                    resolution_engine, resolution_engine.manual_provider
                )
                resolution_context_cls = ResolutionContext
            except Exception as e:
                print(f"Warning: Failed to build INDEX resolution chain: {e}", file=sys.stderr)

            ref_extractor = ReferenceExtractor(
                sheet_name_map=sheet_name_map,
                defined_names={},
                table_refs={},
                name_table_map=name_table_map,
                allow_row_col_ranges=not expand_ranges_for_parity,
            )

            # Parse root cells from roots_data
            print("\nParsing root cells...", file=sys.stderr)
            root_cells = _parse_root_cells(roots_data, sheet_name_map)
            print(f"  Root cells: {len(root_cells)}", file=sys.stderr)
            for root in root_cells:
                print(f"    {root}", file=sys.stderr)
            root_spec = roots_data.get("user_root", {})

            # Parse date format metadata from styles.xml
            dfi = parse_date_format_info(workbook_path)
            if dfi.date_style_indices:
                print(f"  Date-formatted styles: {len(dfi.date_style_indices)}", file=sys.stderr)

            # Execute root-based traversal
            print("\nExecuting root-based traversal...", file=sys.stderr)
            traversal_result = _execute_root_traversal(
                workbook_path=str(workbook_path),
                root_cells=root_cells,
                sheets=sheets,
                sheet_name_map=sheet_name_map,
                ref_extractor=ref_extractor,
                resolution_engine=resolution_engine,
                index_chain=index_chain,
                resolution_context_cls=resolution_context_cls,
                expand_ranges_for_parity=expand_ranges_for_parity,
                queue_all_range_cells=queue_all_range_cells,
                range_traversal_max_cells=range_traversal_max_cells,
                date_format_info=dfi,
            )

            total_cells = traversal_result["total_cells"]
            total_formulas = traversal_result["total_formulas"]
            resolution_metrics = traversal_result.get("resolution_metrics", {})

            print(f"  Total cells visited: {total_cells}", file=sys.stderr)
            print(f"  Total formulas: {total_formulas}", file=sys.stderr)

            # Load cells, formulas, and edges from traversal result
            loader.load_raw_cells(iter(traversal_result["cells"]))
            loader.load_raw_formulas(iter(traversal_result["formulas"]))
            loader.load_raw_json_blobs(iter(traversal_result["json_blobs"]))
            loader.load_raw_edges_internal(iter(traversal_result["edges_internal"]))
            loader.load_raw_edges_range(iter(traversal_result["edges_range"]))
            loader.load_raw_edges_external(iter(traversal_result["edges_external"]))

            grouping_metrics = {
                "bindings_total": 0,
                "bindings_formula": 0,
                "bindings_constant": 0,
                "binding_edges_total": 0,
            }

            post_processing_telemetry = {
                "stages": {},
                "status": "ok",
                "peak_rss_mb": None,
            }
            post_processing_start = time.perf_counter()
            post_processing_peak = 0.0

            def record_stage(stage_name: str, stage_fn: Any) -> None:
                nonlocal post_processing_peak
                stage_start = time.perf_counter()
                rss_before = _get_rss_mb()
                status = "ok"
                error_message = None
                try:
                    stage_fn()
                except Exception as exc:
                    status = "fail"
                    error_message = str(exc)
                    post_processing_telemetry["status"] = "fail"
                    post_processing_telemetry["error"] = {
                        "stage": stage_name,
                        "message": error_message,
                    }
                    raise
                finally:
                    duration = time.perf_counter() - stage_start
                    rss_after = _get_rss_mb()
                    stage_entry = {
                        "duration_s": round(duration, 3),
                        "rss_before_mb": round(rss_before, 2) if rss_before is not None else None,
                        "rss_after_mb": round(rss_after, 2) if rss_after is not None else None,
                        "status": status,
                    }
                    if error_message:
                        stage_entry["error"] = error_message
                    post_processing_telemetry["stages"][stage_name] = stage_entry
                    for val in (rss_before, rss_after):
                        if val is not None and val > post_processing_peak:
                            post_processing_peak = val

            # Step 5: Finalize tables (dedupe + deterministic ordering)
            print("\nFinalizing tables...", file=sys.stderr)
            record_stage("finalize_tables", loader.finalize)

            # Step 5.25: Extract VBA (all procedures, declarations, events, security).
            # Uses the Phase 2 grammar-based extractor (xl_marinade.core.vba).
            # Populates the udfs table for Phase 1 backward compat.
            # Failures are logged but do not fail extraction.
            vba_stats = [{}]

            def _run_vba_extraction() -> None:
                vba_stats[0] = _extract_and_store_vba(workbook_path, loader.conn)
                s = vba_stats[0]
                print(
                    f"  VBA: {s.get('procedures', 0)} procs, "
                    f"{s.get('events', 0)} events, "
                    f"{s.get('declarations', 0)} decls, "
                    f"{s.get('udfs_inserted', 0)} UDFs stored, "
                    f"{s.get('security_findings', 0)} security, "
                    f"{s.get('parse_errors', 0)} errors",
                    file=sys.stderr,
                )

            record_stage("extract_vba", _run_vba_extraction)

            # Step 5.5: Drop raw staging tables to reduce artifact size
            print("Dropping raw staging tables...", file=sys.stderr)
            record_stage("drop_raw_tables", loader.drop_raw_tables)

            # Step 6: Create views and indexes
            print("Creating views and indexes...", file=sys.stderr)
            record_stage("create_views_indexes", loader.create_views)

            def _write_user_roots() -> None:
                sheet_name = root_spec.get("sheet", "")
                range_a1 = root_spec.get("range", "")
                label_hint = root_spec.get("label_hint")
                if not sheet_name or not range_a1:
                    return
                sheet_info = sheet_name_map.get(sheet_name.lower())
                canonical_sheet = sheet_info[1] if sheet_info else sheet_name
                loader.conn.execute("DELETE FROM user_roots")
                loader.conn.execute(
                    """
                    INSERT INTO user_roots (root_id, sheet, range_a1, label_hint)
                    VALUES (?, ?, ?, ?)
                    """,
                    (1, canonical_sheet, range_a1, label_hint),
                )
                loader.conn.commit()

            record_stage("write_user_roots", _write_user_roots)

            def _write_resolution_metrics() -> None:
                if not resolution_metrics:
                    return
                rows = [
                    (function_name, status, count)
                    for function_name, status_map in resolution_metrics.items()
                    for status, count in status_map.items()
                ]
                if not rows:
                    return
                loader.conn.execute("DELETE FROM resolution_metrics")
                loader.conn.executemany(
                    """
                    INSERT INTO resolution_metrics (function_name, status, count)
                    VALUES (?, ?, ?)
                    """,
                    rows,
                )
                loader.conn.commit()

            record_stage("write_resolution_metrics", _write_resolution_metrics)

            def _write_defined_names() -> None:
                if name_table_map is None:
                    return
                rows = []
                for info in name_table_map._names.values():
                    rows.append(
                        (
                            info.name,
                            info.scope,
                            json.dumps(info.ranges, ensure_ascii=False),
                            1 if info.is_external else 0,
                        )
                    )
                if not rows:
                    return
                loader.conn.execute("DELETE FROM defined_names")
                loader.conn.executemany(
                    """
                    INSERT INTO defined_names (name, scope, destinations, is_external)
                    VALUES (?, ?, ?, ?)
                    """,
                    rows,
                )
                loader.conn.commit()

            record_stage("write_defined_names", _write_defined_names)

            # Data-validation rules + cell comments (additive channel; no
            # traversal / bindings / cells impact).
            def _extract_validations_comments() -> None:
                from xl_marinade.core.new_arch.validations_comments import (
                    extract_and_store_validations_comments,
                )

                counts = extract_and_store_validations_comments(
                    workbook_path, sheets_full, loader.conn
                )
                if counts["validations"] or counts["comments"]:
                    print(
                        f"  Data validations: {counts['validations']}, "
                        f"cell comments: {counts['comments']}",
                        file=sys.stderr,
                    )

            record_stage("extract_validations_comments", _extract_validations_comments)

            # Step 6.5: Run grouping/refinement (Story 03)
            def _run_grouping() -> None:
                nonlocal grouping_metrics
                print("\nRunning grouping and refinement...", file=sys.stderr)
                grouping_metrics = run_grouping_on_fast_output(
                    conn=loader.conn,
                    workbook_sha256=workbook_sha256,
                    workbook_path=str(workbook_path),
                    ir_db_path=str(build_db_path),
                    workbook=shared_wb,
                )
                print(f"  Bindings created: {grouping_metrics['bindings_total']}", file=sys.stderr)
                print(f"    Formula: {grouping_metrics['bindings_formula']}", file=sys.stderr)
                print(f"    Constant: {grouping_metrics['bindings_constant']}", file=sys.stderr)
                print(
                    f"  Binding edges: {grouping_metrics['binding_edges_total']}", file=sys.stderr
                )

            record_stage("grouping_refinement", _run_grouping)

            # Step 6.75: Time-axis inference + time-dependence annotation (Sprint 6)
            def _infer_time_axis() -> None:
                # These are deterministic post-processing steps on the extracted IR.
                from xl_marinade.core import schema as ir_schema
                from xl_marinade.core.time_axis_inference import (
                    infer_time_index_candidates_from_conn,
                )
                from xl_marinade.core.time_dependence import infer_time_dependence_from_conn

                candidates = infer_time_index_candidates_from_conn(loader.conn)
                if candidates:
                    ir_schema.insert_time_index_candidates(loader.conn, candidates)
                    annotations = infer_time_dependence_from_conn(loader.conn)
                    if annotations:
                        ir_schema.insert_binding_time_annotations(loader.conn, annotations)
                    loader.conn.commit()

                # Optional visibility (stderr) for debugging/ops
                proj = (
                    [c for c in candidates if c.get("sheet") == "Projection"] if candidates else []
                )
                if proj:
                    best = proj[0]
                    print(
                        f"  Time-axis candidates: Projection rank1 binding_id={best.get('binding_id')} conf={best.get('confidence')}",
                        file=sys.stderr,
                    )

            record_stage("time_axis_inference", _infer_time_axis)

            # Step 6.8: Table candidate extraction (Sprint 10 Story 7)
            def _extract_table_candidates() -> None:
                from xl_marinade.core.new_arch.table_candidates import extract_table_candidates

                extract_table_candidates(conn=loader.conn)

                # Optional visibility for debugging
                count = loader.conn.execute("SELECT COUNT(*) FROM table_candidates").fetchone()[0]
                if count > 0:
                    print(f"  Table candidates: {count}", file=sys.stderr)

            record_stage("table_candidate_extraction", _extract_table_candidates)

            # Step 6.85: Formula family extraction
            def _extract_formula_families() -> None:
                from xl_marinade.core.new_arch.formula_families import extract_formula_families

                extract_formula_families(conn=loader.conn)
                count = loader.conn.execute("SELECT COUNT(*) FROM formula_families").fetchone()[0]
                if count > 0:
                    print(f"  Formula families: {count}", file=sys.stderr)

            record_stage("formula_family_extraction", _extract_formula_families)

            # Step 6.9: Wire VBA edges into binding graph
            vba_edge_metrics = [{}]

            def _wire_vba() -> None:
                vba_edge_metrics[0] = _wire_vba_edges(loader.conn, enrich=enrich)
                m = vba_edge_metrics[0]
                if any(
                    m.get(k, 0) > 0
                    for k in (
                        "cell_udf_edges",
                        "call_graph_edges",
                        "static_ref_edges",
                        "llm_edges_created",
                    )
                ):
                    parts = [
                        f"{m.get('cell_udf_edges', 0)} cell→UDF",
                        f"{m.get('call_graph_edges', 0)} proc→proc",
                        f"{m.get('static_ref_edges', 0)} static refs ({m.get('static_refs_extracted', 0)} extracted)",
                    ]
                    if m.get("llm_refs_inferred", 0) > 0:
                        parts.append(
                            f"{m.get('llm_edges_created', 0)} LLM refs "
                            f"({m.get('llm_refs_inferred', 0)} inferred, "
                            f"{m.get('llm_latency_s', 0):.1f}s)"
                        )
                    print(f"  VBA edges: {', '.join(parts)}", file=sys.stderr)

            record_stage("wire_vba_edges", _wire_vba)

            # Step 6.95: R21 Fix D — OFFSET-formula edge synthesis
            # Static analysis of OFFSET(...) calls in cell formulas emits
            # `via_offset_static` / `via_offset_volatile` binding_edges so the
            # dependency walker can reach through OFFSET-mediated bridges
            # (R20: Risk Drivers→Calculation Engine, 164+ refs). Independent
            # of VBA — runs for VBA and non-VBA workbooks alike.
            offset_edge_metrics = [{}]

            def _synthesise_offset_edges() -> None:
                try:
                    from xl_marinade.core.new_arch.offset_edges import synthesize_offset_edges

                    offset_edge_metrics[0] = synthesize_offset_edges(loader.conn)
                except Exception as exc:  # noqa: BLE001 — defensive: never fail the build
                    print(f"  Warning: OFFSET edge synthesis failed: {exc}", file=sys.stderr)
                    return
                m = offset_edge_metrics[0]
                if m.get("calls_seen", 0) > 0:
                    print(
                        f"  OFFSET edges: {m.get('edges_inserted', 0)} inserted "
                        f"({m.get('calls_static', 0)} static, "
                        f"{m.get('calls_volatile', 0)} volatile, "
                        f"{m.get('calls_skipped', 0)} skipped)",
                        file=sys.stderr,
                    )

            record_stage("synthesise_offset_edges", _synthesise_offset_edges)

            # Write metadata
            def _write_metadata() -> None:
                loader.conn.execute("""
                    CREATE TABLE IF NOT EXISTS ir_metadata (
                        key TEXT PRIMARY KEY,
                        value TEXT NOT NULL
                    )
                """)

                metadata = {
                    "schema_version": SCHEMA_VERSION,
                    "build_mode": BUILD_MODE,
                    "extraction_mode": "rooted",
                    "sqlite_version": sqlite3.sqlite_version,
                    "extractor_git_sha": git_sha,
                    "workbook_sha256": workbook_sha256,
                    "original_filename": Path(workbook_path).name,
                }
                _doc_title = read_workbook_doc_title(Path(workbook_path))
                if _doc_title:
                    metadata["doc_title"] = _doc_title

                loader.conn.executemany(
                    "INSERT OR REPLACE INTO ir_metadata (key, value) VALUES (?, ?)",
                    metadata.items(),
                )
                loader.conn.commit()

            record_stage("write_metadata", _write_metadata)

            # Step 7: VACUUM INTO to produce canonical artifact
            print("Creating canonical artifact...", file=sys.stderr)
            record_stage("vacuum_into", lambda: loader.vacuum_into(str(output_db)))

            post_processing_telemetry["total_time_s"] = round(
                time.perf_counter() - post_processing_start, 3
            )
            post_processing_telemetry["peak_rss_mb"] = (
                round(post_processing_peak, 2) if post_processing_peak else None
            )

            memory_telemetry = loader.get_memory_telemetry()

        # Close shared workbook (WI-11)
        if shared_wb:
            try:
                shared_wb.__exit__(None, None, None)
            except Exception:
                pass

        # Build DB is automatically closed by context manager
        # Remove build DB
        if build_db_path.exists():
            build_db_path.unlink()

        # Collect final metrics
        output_size_mb = output_db.stat().st_size / (1024 * 1024)
        run_time_s = time.perf_counter() - run_start
        traversal_telemetry = traversal_result.get("telemetry", {})

        if memory_telemetry:
            print(
                f"  Memory telemetry: peak_rss={memory_telemetry.get('peak_rss_mb')}MB, "
                f"checks={memory_telemetry.get('check_count')}",
                file=sys.stderr,
            )

        print("\n✅ Fast extraction complete!", file=sys.stderr)
        print(f"  Total cells: {total_cells}", file=sys.stderr)
        print(f"  Total formulas: {total_formulas}", file=sys.stderr)
        print(f"  Total bindings: {grouping_metrics['bindings_total']}", file=sys.stderr)
        print(f"  Output size: {output_size_mb:.2f} MB", file=sys.stderr)
        print(f"  Schema version: {SCHEMA_VERSION}", file=sys.stderr)
        print(f"  Build mode: {BUILD_MODE}", file=sys.stderr)
        print(f"  Total time: {run_time_s:.2f}s", file=sys.stderr)

        telemetry_payload = {
            "schema_version": SCHEMA_VERSION,
            "build_mode": BUILD_MODE,
            "sqlite_version": sqlite3.sqlite_version,
            "extractor_git_sha": git_sha,
            "workbook_sha256": workbook_sha256,
            "ir_db_path": str(output_db),
            "total_cells": total_cells,
            "total_formulas": total_formulas,
            "bindings_total": grouping_metrics["bindings_total"],
            "bindings_formula": grouping_metrics["bindings_formula"],
            "bindings_constant": grouping_metrics["bindings_constant"],
            "binding_edges_total": grouping_metrics["binding_edges_total"],
            "telemetry": {
                "run_time_s": round(run_time_s, 3),
                "traversal": traversal_telemetry,
                "memory_budget": memory_telemetry,
                "post_processing": post_processing_telemetry,
            },
        }

        telemetry_path = output_db.parent / "telemetry.json"
        telemetry_fast_path = output_db.parent / "telemetry_fast.json"
        try:
            telemetry_path.parent.mkdir(parents=True, exist_ok=True)
            with open(telemetry_path, "w", encoding="utf-8") as handle:
                json.dump(telemetry_payload, handle, indent=2)
            with open(telemetry_fast_path, "w", encoding="utf-8") as handle:
                json.dump(telemetry_payload, handle, indent=2)
        except OSError as exc:
            print(f"Warning: failed to write telemetry.json: {exc}", file=sys.stderr)

        return telemetry_payload

    except Exception as e:
        telemetry_path = output_db.parent / "telemetry.json"
        telemetry_fast_path = output_db.parent / "telemetry_fast.json"
        failure_payload = {
            "schema_version": SCHEMA_VERSION,
            "build_mode": BUILD_MODE,
            "sqlite_version": sqlite3.sqlite_version,
            "extractor_git_sha": git_sha,
            "workbook_sha256": workbook_sha256,
            "ir_db_path": str(output_db),
            "telemetry": {
                "error": str(e),
            },
        }
        try:
            telemetry_path.parent.mkdir(parents=True, exist_ok=True)
            with open(telemetry_path, "w", encoding="utf-8") as handle:
                json.dump(failure_payload, handle, indent=2)
            with open(telemetry_fast_path, "w", encoding="utf-8") as handle:
                json.dump(failure_payload, handle, indent=2)
        except OSError:
            pass
        # Clean up build DB on error
        if build_db_path.exists():
            build_db_path.unlink()
        raise


def run_full_workbook_extraction(
    workbook_path: Path,
    output_db: Path,
    max_memory_mb: int = 1800,
    batch_size: int = 10_000,
    allow_non_canonical: bool = False,
    normalize_bindings: bool = True,
    enrich: bool = False,
) -> dict[str, Any]:
    """
    Run full-workbook extraction — process all cells on all sheets without a root.

    Shares the same 7-step build procedure as run_fast_extraction() but calls
    _execute_full_workbook_extraction() instead of _execute_root_traversal().
    No root cell is needed; no user_roots table is populated.

    Writes extraction_mode='full_workbook' to ir_metadata.
    """
    # Step 1: Verify SQLite version
    version = sqlite3.sqlite_version_info
    if version < (3, 27, 0):
        print(
            f"Error: SQLite 3.27.0+ required for fast extraction, found {sqlite3.sqlite_version}",
            file=sys.stderr,
        )
        raise ExtractionError(f"SQLite 3.27.0+ required, found {sqlite3.sqlite_version}")

    run_start = time.perf_counter()
    print("Full-workbook extraction starting...", file=sys.stderr)
    print(f"  Workbook: {workbook_path.name}", file=sys.stderr)
    print(f"  Output: {output_db}", file=sys.stderr)
    print(f"  Memory budget: {max_memory_mb} MB", file=sys.stderr)

    git_sha = get_git_sha()
    workbook_sha256 = compute_workbook_sha256(workbook_path)

    build_db_path = output_db.parent / f"{output_db.stem}_build.db"

    schema_path = Path(__file__).parent / "schema.sql"
    with open(schema_path, encoding="utf-8") as f:
        schema_sql = f.read()

    try:
        memory_config = MemoryBudgetConfig(
            max_memory_mb=max_memory_mb, warn_threshold_pct=0.8, check_interval_rows=10_000
        )

        with BulkLoader(
            build_db_path=str(build_db_path),
            batch_size=batch_size,
            allow_non_canonical=allow_non_canonical,
            memory_budget_config=memory_config,
        ) as loader:
            loader.create_schema(schema_sql)
            print("  Schema created", file=sys.stderr)

            # Extract sheet catalog
            print("\nExtracting workbook metadata...", file=sys.stderr)
            catalog = WorkbookCatalog(str(workbook_path))
            sheets_full = catalog.sheets
            sheets = [(sheet_id, sheet_name) for sheet_id, sheet_name, _, _ in sheets_full]
            print(f"  Sheets found: {len(sheets)}", file=sys.stderr)

            loader.load_sheets(sheets)

            sheet_name_map = {
                sheet_name.lower(): (sheet_id, sheet_name) for sheet_id, sheet_name in sheets
            }

            # WI-11: Share workbook across pipeline stages
            shared_wb = None
            name_table_map = None
            try:
                shared_wb = LazyWorkbook(workbook_path, data_only=False, keep_vba=False)
                shared_wb.__enter__()
                name_table_map = NameTableMap(shared_wb)
            except Exception as e:
                print(f"Warning: Failed to build name/table map: {e}", file=sys.stderr)

            resolution_engine = ResolutionEngine(
                LazyValueFetcher(workbook_path, max_cached_sheets=None)
            )
            index_chain = None
            resolution_context_cls = None
            try:
                index_chain = create_index_resolution_chain(
                    resolution_engine, resolution_engine.manual_provider
                )
                resolution_context_cls = ResolutionContext
            except Exception as e:
                print(f"Warning: Failed to build INDEX resolution chain: {e}", file=sys.stderr)

            ref_extractor = ReferenceExtractor(
                sheet_name_map=sheet_name_map,
                defined_names={},
                table_refs={},
                name_table_map=name_table_map,
                allow_row_col_ranges=True,
            )

            dfi = parse_date_format_info(workbook_path)
            if dfi.date_style_indices:
                print(f"  Date-formatted styles: {len(dfi.date_style_indices)}", file=sys.stderr)

            # Execute full-workbook traversal (no roots, all sheets)
            print("\nExecuting full-workbook extraction...", file=sys.stderr)
            traversal_result = _execute_full_workbook_extraction(
                workbook_path=str(workbook_path),
                sheets=sheets,
                sheet_name_map=sheet_name_map,
                ref_extractor=ref_extractor,
                resolution_engine=resolution_engine,
                index_chain=index_chain,
                resolution_context_cls=resolution_context_cls,
                date_format_info=dfi,
            )

            total_cells = traversal_result["total_cells"]
            total_formulas = traversal_result["total_formulas"]
            resolution_metrics = traversal_result.get("resolution_metrics", {})

            print(f"  Total cells: {total_cells}", file=sys.stderr)
            print(f"  Total formulas: {total_formulas}", file=sys.stderr)

            # Load into bulk loader
            loader.load_raw_cells(iter(traversal_result["cells"]))
            loader.load_raw_formulas(iter(traversal_result["formulas"]))
            loader.load_raw_json_blobs(iter(traversal_result["json_blobs"]))
            loader.load_raw_edges_internal(iter(traversal_result["edges_internal"]))
            loader.load_raw_edges_range(iter(traversal_result["edges_range"]))
            loader.load_raw_edges_external(iter(traversal_result["edges_external"]))

            grouping_metrics = {
                "bindings_total": 0,
                "bindings_formula": 0,
                "bindings_constant": 0,
                "binding_edges_total": 0,
            }

            post_processing_telemetry: dict[str, Any] = {
                "stages": {},
                "status": "ok",
                "peak_rss_mb": None,
            }
            post_processing_start = time.perf_counter()
            post_processing_peak = 0.0

            def record_stage(stage_name: str, stage_fn: Any) -> None:
                nonlocal post_processing_peak
                stage_start = time.perf_counter()
                rss_before = _get_rss_mb()
                status = "ok"
                error_message = None
                try:
                    stage_fn()
                except Exception as exc:
                    status = "fail"
                    error_message = str(exc)
                    post_processing_telemetry["status"] = "fail"
                    post_processing_telemetry["error"] = {
                        "stage": stage_name,
                        "message": error_message,
                    }
                    raise
                finally:
                    duration = time.perf_counter() - stage_start
                    rss_after = _get_rss_mb()
                    stage_entry = {
                        "duration_s": round(duration, 3),
                        "rss_before_mb": round(rss_before, 2) if rss_before is not None else None,
                        "rss_after_mb": round(rss_after, 2) if rss_after is not None else None,
                        "status": status,
                    }
                    if error_message:
                        stage_entry["error"] = error_message
                    post_processing_telemetry["stages"][stage_name] = stage_entry
                    for val in (rss_before, rss_after):
                        if val is not None and val > post_processing_peak:
                            post_processing_peak = val

            # Steps 5-7: same as run_fast_extraction
            print("\nFinalizing tables...", file=sys.stderr)
            record_stage("finalize_tables", loader.finalize)

            vba_stats = [{}]

            def _run_vba_extraction() -> None:
                vba_stats[0] = _extract_and_store_vba(workbook_path, loader.conn)
                s = vba_stats[0]
                print(
                    f"  VBA: {s.get('procedures', 0)} procs, "
                    f"{s.get('events', 0)} events, "
                    f"{s.get('declarations', 0)} decls, "
                    f"{s.get('udfs_inserted', 0)} UDFs stored, "
                    f"{s.get('security_findings', 0)} security, "
                    f"{s.get('parse_errors', 0)} errors",
                    file=sys.stderr,
                )

            record_stage("extract_vba", _run_vba_extraction)

            print("Dropping raw staging tables...", file=sys.stderr)
            record_stage("drop_raw_tables", loader.drop_raw_tables)

            print("Creating views and indexes...", file=sys.stderr)
            record_stage("create_views_indexes", loader.create_views)

            # No user_roots for full-workbook extraction
            def _write_no_user_roots() -> None:
                loader.conn.execute("DELETE FROM user_roots")
                loader.conn.commit()

            record_stage("write_user_roots", _write_no_user_roots)

            def _write_resolution_metrics() -> None:
                if not resolution_metrics:
                    return
                rows = [
                    (function_name, status, count)
                    for function_name, status_map in resolution_metrics.items()
                    for status, count in status_map.items()
                ]
                if not rows:
                    return
                loader.conn.execute("DELETE FROM resolution_metrics")
                loader.conn.executemany(
                    "INSERT INTO resolution_metrics (function_name, status, count) VALUES (?, ?, ?)",
                    rows,
                )
                loader.conn.commit()

            record_stage("write_resolution_metrics", _write_resolution_metrics)

            def _write_defined_names() -> None:
                if name_table_map is None:
                    return
                rows = []
                for info in name_table_map._names.values():
                    rows.append(
                        (
                            info.name,
                            info.scope,
                            json.dumps(info.ranges, ensure_ascii=False),
                            1 if info.is_external else 0,
                        )
                    )
                if not rows:
                    return
                loader.conn.execute("DELETE FROM defined_names")
                loader.conn.executemany(
                    "INSERT INTO defined_names (name, scope, destinations, is_external) VALUES (?, ?, ?, ?)",
                    rows,
                )
                loader.conn.commit()

            record_stage("write_defined_names", _write_defined_names)

            # Data-validation rules + cell comments (additive channel; no
            # traversal / bindings / cells impact).
            def _extract_validations_comments() -> None:
                from xl_marinade.core.new_arch.validations_comments import (
                    extract_and_store_validations_comments,
                )

                counts = extract_and_store_validations_comments(
                    workbook_path, sheets_full, loader.conn
                )
                if counts["validations"] or counts["comments"]:
                    print(
                        f"  Data validations: {counts['validations']}, "
                        f"cell comments: {counts['comments']}",
                        file=sys.stderr,
                    )

            record_stage("extract_validations_comments", _extract_validations_comments)

            def _run_grouping() -> None:
                nonlocal grouping_metrics
                print("\nRunning grouping and refinement...", file=sys.stderr)
                grouping_metrics = run_grouping_on_fast_output(
                    conn=loader.conn,
                    workbook_sha256=workbook_sha256,
                    workbook_path=str(workbook_path),
                    ir_db_path=str(build_db_path),
                    workbook=shared_wb,
                )
                print(f"  Bindings created: {grouping_metrics['bindings_total']}", file=sys.stderr)

            record_stage("grouping_refinement", _run_grouping)

            # Step 6.6: variable normalization (H2+H9) — split over-merged
            # bindings into per-variable columns/rows while keeping genuine 2D
            # matrices whole. Runs after grouping (bindings/edges exist) and
            # before labels/time-axis/table/family extraction (which key on the
            # final binding set). Default ON; --no-normalize-bindings disables it.
            normalize_metrics = [{}]

            def _normalize_bindings_fn() -> None:
                if not normalize_bindings:
                    return
                from xl_marinade.core.new_arch.variable_normalizer import (
                    normalize_bindings_fn,
                )

                normalize_metrics[0] = normalize_bindings_fn(
                    loader.conn,
                    workbook_path=str(workbook_path),
                    workbook_sha256=workbook_sha256,
                    name_table_map=name_table_map,
                )
                m = normalize_metrics[0]
                print(
                    f"  Variable normalization: {m.get('bindings_split', 0)} split, "
                    f"{m.get('bindings_kept_matrix', 0)} matrices kept, "
                    f"{m.get('bindings_header_stripped', 0)} headers stripped, "
                    f"{m.get('bindings_vba_split', 0)} vba panels split, "
                    f"{m.get('bindings_listobject_collapsed', 0)} listobjects collapsed, "
                    f"+{m.get('bindings_added', 0)}/-{m.get('bindings_removed', 0)} bindings",
                    file=sys.stderr,
                )

            record_stage("variable_normalize", _normalize_bindings_fn)

            def _backfill_binding_labels_fn() -> None:
                # Phase-1.5 Lever C ran as a standalone script; every ad-hoc
                # re-extraction that skipped it shipped all-NULL bindings.label
                # (Cycle 17: several real-model regressions). Inline so a fresh
                # ir.db is always label-complete.
                from xl_marinade.core.labelling.simple_labeller import backfill_binding_labels

                real, fb = backfill_binding_labels(loader.conn)
                print(f"  Binding labels: {real} from candidates, {fb} fallback", file=sys.stderr)

            record_stage("backfill_binding_labels", _backfill_binding_labels_fn)

            def _infer_time_axis() -> None:
                from xl_marinade.core import schema as ir_schema
                from xl_marinade.core.time_axis_inference import (
                    infer_time_index_candidates_from_conn,
                )
                from xl_marinade.core.time_dependence import infer_time_dependence_from_conn

                candidates = infer_time_index_candidates_from_conn(loader.conn)
                if candidates:
                    ir_schema.insert_time_index_candidates(loader.conn, candidates)
                    annotations = infer_time_dependence_from_conn(loader.conn)
                    if annotations:
                        ir_schema.insert_binding_time_annotations(loader.conn, annotations)
                    loader.conn.commit()

            record_stage("time_axis_inference", _infer_time_axis)

            def _extract_table_candidates_fn() -> None:
                from xl_marinade.core.new_arch.table_candidates import extract_table_candidates

                extract_table_candidates(conn=loader.conn)

            record_stage("table_candidate_extraction", _extract_table_candidates_fn)

            def _extract_formula_families_fn() -> None:
                from xl_marinade.core.new_arch.formula_families import extract_formula_families

                extract_formula_families(conn=loader.conn)

            record_stage("formula_family_extraction", _extract_formula_families_fn)

            # Wire VBA edges into binding graph
            vba_edge_metrics = [{}]

            def _wire_vba() -> None:
                vba_edge_metrics[0] = _wire_vba_edges(loader.conn, enrich=enrich)
                m = vba_edge_metrics[0]
                if any(
                    m.get(k, 0) > 0
                    for k in (
                        "cell_udf_edges",
                        "call_graph_edges",
                        "static_ref_edges",
                        "llm_edges_created",
                    )
                ):
                    parts = [
                        f"{m.get('cell_udf_edges', 0)} cell→UDF",
                        f"{m.get('call_graph_edges', 0)} proc→proc",
                        f"{m.get('static_ref_edges', 0)} static refs ({m.get('static_refs_extracted', 0)} extracted)",
                    ]
                    if m.get("llm_refs_inferred", 0) > 0:
                        parts.append(
                            f"{m.get('llm_edges_created', 0)} LLM refs "
                            f"({m.get('llm_refs_inferred', 0)} inferred, "
                            f"{m.get('llm_latency_s', 0):.1f}s)"
                        )
                    print(f"  VBA edges: {', '.join(parts)}", file=sys.stderr)

            record_stage("wire_vba_edges", _wire_vba)

            # R21 Fix D — OFFSET-formula edge synthesis (see rooted-mode
            # block above for rationale). Same hook for full_workbook so
            # OFFSET-mediated bridges are wired regardless of extraction mode.
            offset_edge_metrics_fw = [{}]

            def _synthesise_offset_edges_fw() -> None:
                try:
                    from xl_marinade.core.new_arch.offset_edges import synthesize_offset_edges

                    offset_edge_metrics_fw[0] = synthesize_offset_edges(loader.conn)
                except Exception as exc:  # noqa: BLE001 — defensive
                    print(f"  Warning: OFFSET edge synthesis failed: {exc}", file=sys.stderr)
                    return
                m = offset_edge_metrics_fw[0]
                if m.get("calls_seen", 0) > 0:
                    print(
                        f"  OFFSET edges: {m.get('edges_inserted', 0)} inserted "
                        f"({m.get('calls_static', 0)} static, "
                        f"{m.get('calls_volatile', 0)} volatile, "
                        f"{m.get('calls_skipped', 0)} skipped)",
                        file=sys.stderr,
                    )

            record_stage("synthesise_offset_edges", _synthesise_offset_edges_fw)

            # Write metadata with extraction_mode='full_workbook'
            def _write_metadata() -> None:
                loader.conn.execute("""
                    CREATE TABLE IF NOT EXISTS ir_metadata (
                        key TEXT PRIMARY KEY,
                        value TEXT NOT NULL
                    )
                """)
                metadata = {
                    "schema_version": SCHEMA_VERSION,
                    "build_mode": BUILD_MODE,
                    "extraction_mode": "full_workbook",
                    "sqlite_version": sqlite3.sqlite_version,
                    "extractor_git_sha": git_sha,
                    "workbook_sha256": workbook_sha256,
                    "original_filename": Path(workbook_path).name,
                }
                _doc_title = read_workbook_doc_title(Path(workbook_path))
                if _doc_title:
                    metadata["doc_title"] = _doc_title
                loader.conn.executemany(
                    "INSERT OR REPLACE INTO ir_metadata (key, value) VALUES (?, ?)",
                    metadata.items(),
                )
                loader.conn.commit()

            record_stage("write_metadata", _write_metadata)

            print("Creating canonical artifact...", file=sys.stderr)
            record_stage("vacuum_into", lambda: loader.vacuum_into(str(output_db)))

            post_processing_telemetry["total_time_s"] = round(
                time.perf_counter() - post_processing_start, 3
            )
            post_processing_telemetry["peak_rss_mb"] = (
                round(post_processing_peak, 2) if post_processing_peak else None
            )

            memory_telemetry = loader.get_memory_telemetry()

        # Close shared workbook (WI-11)
        if shared_wb:
            try:
                shared_wb.__exit__(None, None, None)
            except Exception:
                pass

        if build_db_path.exists():
            build_db_path.unlink()

        output_size_mb = output_db.stat().st_size / (1024 * 1024)
        run_time_s = time.perf_counter() - run_start
        traversal_telemetry = traversal_result.get("telemetry", {})

        print("\n✅ Full-workbook extraction complete!", file=sys.stderr)
        print(f"  Total cells: {total_cells}", file=sys.stderr)
        print(f"  Total formulas: {total_formulas}", file=sys.stderr)
        print(f"  Total bindings: {grouping_metrics['bindings_total']}", file=sys.stderr)
        print(f"  Output size: {output_size_mb:.2f} MB", file=sys.stderr)
        print(f"  Total time: {run_time_s:.2f}s", file=sys.stderr)

        telemetry_payload = {
            "schema_version": SCHEMA_VERSION,
            "build_mode": BUILD_MODE,
            "extraction_mode": "full_workbook",
            "sqlite_version": sqlite3.sqlite_version,
            "extractor_git_sha": git_sha,
            "workbook_sha256": workbook_sha256,
            "ir_db_path": str(output_db),
            "total_cells": total_cells,
            "total_formulas": total_formulas,
            "bindings_total": grouping_metrics["bindings_total"],
            "bindings_formula": grouping_metrics["bindings_formula"],
            "bindings_constant": grouping_metrics["bindings_constant"],
            "binding_edges_total": grouping_metrics["binding_edges_total"],
            "telemetry": {
                "run_time_s": round(run_time_s, 3),
                "traversal": traversal_telemetry,
                "memory_budget": memory_telemetry,
                "post_processing": post_processing_telemetry,
            },
        }

        telemetry_path = output_db.parent / "telemetry.json"
        try:
            telemetry_path.parent.mkdir(parents=True, exist_ok=True)
            with open(telemetry_path, "w", encoding="utf-8") as handle:
                json.dump(telemetry_payload, handle, indent=2)
        except OSError as exc:
            print(f"Warning: failed to write telemetry.json: {exc}", file=sys.stderr)

        return telemetry_payload

    except Exception:
        if build_db_path.exists():
            build_db_path.unlink()
        raise
