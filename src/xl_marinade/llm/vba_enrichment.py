# ABOUTME: LLM-assisted enrichment for VBA references that static analysis can't resolve.
# ABOUTME: Handles Selection-based, .End(), string-built, and dynamic patterns (Tier 3).

"""
VBA LLM Reference Enrichment (Phase 3)

For procedures with unresolved dynamic references (Selection-based operations,
.End() range extensions, string-built addresses), submits the procedure source
plus workbook structure to the LLM and parses inferred cell references.

Also generates an actuarial concept description for each procedure.

Controlled by config: FULL_WORKBOOK_EXTRACTION_ENABLED must be true,
and the enrichment runs as part of the VBA edge wiring step.
"""

import json
import logging
import os
import re
import sqlite3
import time
from dataclasses import dataclass

logger = logging.getLogger(__name__)

# Dynamic patterns that indicate unresolved references needing LLM
_DYNAMIC_INDICATORS = [
    re.compile(r"Selection\.", re.IGNORECASE),
    re.compile(r"\.End\(xl", re.IGNORECASE),
    re.compile(r'Range\("[A-Z]+" & ', re.IGNORECASE),
    re.compile(r"Range\(Selection", re.IGNORECASE),
    re.compile(r"ActiveSheet\b", re.IGNORECASE),
    re.compile(r"ActiveCell\b", re.IGNORECASE),
    re.compile(r"Cells\(\s*\w+\s*,\s*\w+\s*\)", re.IGNORECASE),  # Cells(var, var)
]

SYSTEM_PROMPT = """You are an expert VBA code analyst specialising in actuarial Excel models.

Given a VBA procedure and the workbook's structure, identify:
1. Cell ranges the procedure READS from (input dependencies)
2. Cell ranges the procedure WRITES to (outputs)
3. A one-sentence description of what this procedure does in actuarial terms

For each reference, include:
- sheet: The Excel sheet name
- range: The cell address or range (e.g., "A2:ALI13", "C2", "A4:A1000")
- kind: "read" or "write"
- reasoning: Brief explanation of how you determined this

If a range is dynamic (e.g., extends to the last row of data), describe the likely range
based on the code structure (e.g., "A2:A{lastRow}" becomes "A2:A1000 (estimated)").

If you cannot determine the exact range, provide your best estimate with "(estimated)" suffix.

Respond with valid JSON only:
{
  "reads": [{"sheet": "...", "range": "...", "reasoning": "..."}],
  "writes": [{"sheet": "...", "range": "...", "reasoning": "..."}],
  "description": "One sentence describing the procedure's actuarial role"
}"""


@dataclass
class LLMInferredRef:
    sheet: str
    target: str
    ref_kind: str  # 'read' or 'write'
    reasoning: str


@dataclass
class LLMEnrichmentResult:
    procedure_name: str
    refs: list[LLMInferredRef]
    description: str
    model_used: str
    latency_s: float


def _needs_enrichment(body: str) -> bool:
    """Check if a procedure body contains dynamic patterns that static analysis can't resolve."""
    for pattern in _DYNAMIC_INDICATORS:
        if pattern.search(body):
            return True
    return False


def _build_workbook_context(conn: sqlite3.Connection) -> str:
    """Build workbook structure context for the LLM prompt."""
    parts = ["Workbook structure:"]

    # Sheet names
    try:
        sheets = conn.execute("SELECT sheet_name FROM sheets ORDER BY sheet_id").fetchall()
        parts.append(f"Sheets: {', '.join(r[0] for r in sheets)}")
    except sqlite3.OperationalError:
        pass

    # Named ranges (show first 30 most relevant)
    try:
        names = conn.execute(
            "SELECT name, destinations FROM defined_names ORDER BY name LIMIT 30"
        ).fetchall()
        if names:
            parts.append("\nNamed ranges:")
            for r in names:
                try:
                    dests = json.loads(r[1])
                    if isinstance(dests, list) and dests and not str(dests[0]).startswith("#"):
                        parts.append(f"  {r[0]} → {dests[0]}")
                except (json.JSONDecodeError, IndexError):
                    pass
    except sqlite3.OperationalError:
        pass

    return "\n".join(parts)


def _build_static_context(conn: sqlite3.Connection, procedure_id: int) -> str:
    """Build context from existing static analysis results."""
    try:
        refs = conn.execute(
            """
            SELECT ref_kind, target_kind, target, precision
            FROM vba_procedure_cell_refs
            WHERE procedure_id = ?
        """,
            (procedure_id,),
        ).fetchall()
        if not refs:
            return "No static references found."
        lines = ["Already identified by static analysis:"]
        for r in refs:
            lines.append(f"  {r[0]:10s} {r[2]} ({r[3]})")
        return "\n".join(lines)
    except sqlite3.OperationalError:
        return ""


def enrich_procedure(
    procedure_name: str,
    body: str,
    workbook_context: str,
    static_context: str,
    api_key: str | None = None,
    model: str = "gpt-4.1-nano",
) -> LLMEnrichmentResult | None:
    """Submit a single VBA procedure to the LLM for reference enrichment.

    Returns inferred references and a description, or None on failure.
    """
    try:
        from openai import OpenAI
    except ImportError:
        logger.warning("openai package not installed — skipping LLM enrichment")
        return None

    key = api_key or os.getenv("OPENAI_API_KEY")
    if not key:
        logger.warning("No OPENAI_API_KEY — skipping LLM enrichment")
        return None

    client = OpenAI(api_key=key)

    user_content = f"""Procedure: {procedure_name}

```vba
{body}
```

{workbook_context}

{static_context}

Identify cell ranges this procedure reads from and writes to that are NOT already listed
in the static analysis results above. Focus on Selection-based operations, .End() dynamic
ranges, string-built addresses, and loop-based cell access patterns."""

    start = time.monotonic()
    try:
        response = client.chat.completions.create(
            model=model,
            temperature=0.0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_content},
            ],
        )
        content = response.choices[0].message.content or "{}"
    except Exception as e:
        logger.warning("LLM call failed for %s: %s", procedure_name, e)
        return None

    latency = time.monotonic() - start

    # Parse response
    try:
        data = json.loads(content)
    except json.JSONDecodeError:
        # Try extracting from markdown fences
        m = re.search(r"```(?:json)?\s*(.*?)```", content, re.DOTALL)
        if m:
            try:
                data = json.loads(m.group(1))
            except json.JSONDecodeError:
                logger.warning("Failed to parse LLM response for %s", procedure_name)
                return None
        else:
            logger.warning("Failed to parse LLM response for %s", procedure_name)
            return None

    refs: list[LLMInferredRef] = []

    for entry in data.get("reads", []):
        sheet = entry.get("sheet", "")
        target = entry.get("range", "")
        reasoning = entry.get("reasoning", "")
        if sheet and target:
            # Strip $ signs and (estimated) suffix for the stored address
            clean_target = target.replace("$", "").split(" (")[0]
            refs.append(
                LLMInferredRef(
                    sheet=sheet,
                    target=clean_target,
                    ref_kind="read",
                    reasoning=reasoning,
                )
            )

    for entry in data.get("writes", []):
        sheet = entry.get("sheet", "")
        target = entry.get("range", "")
        reasoning = entry.get("reasoning", "")
        if sheet and target:
            clean_target = target.replace("$", "").split(" (")[0]
            refs.append(
                LLMInferredRef(
                    sheet=sheet,
                    target=clean_target,
                    ref_kind="write",
                    reasoning=reasoning,
                )
            )

    description = data.get("description", "")

    return LLMEnrichmentResult(
        procedure_name=procedure_name,
        refs=refs,
        description=description,
        model_used=model,
        latency_s=latency,
    )


def enrich_and_store(conn: sqlite3.Connection) -> dict:
    """Run LLM enrichment for all procedures with unresolved dynamic references.

    Stores results in vba_procedure_cell_refs with precision='inferred'.

    Returns metrics dict.
    """
    import sys

    metrics = {
        "procedures_submitted": 0,
        "refs_inferred": 0,
        "edges_created": 0,
        "total_latency_s": 0.0,
        "descriptions_generated": 0,
    }

    # Check tables exist
    tables = {
        r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    }
    if "vba_procedures" not in tables or "vba_modules" not in tables:
        return metrics

    # Load procedures (including kind + compile_branch for binding_id construction)
    procs = conn.execute("""
        SELECT p.procedure_id, p.name, p.body, m.name AS module_name,
               p.kind, p.compile_branch
        FROM vba_procedures p
        JOIN vba_modules m ON p.module_id = m.module_id
    """).fetchall()

    # Filter to procedures that need enrichment
    candidates = [
        (p[0], p[1], p[2] or "", p[3], p[4], p[5] or "")
        for p in procs
        if _needs_enrichment(p[2] or "")
    ]
    if not candidates:
        return metrics

    workbook_context = _build_workbook_context(conn)

    print(
        f"  LLM enrichment: {len(candidates)} procedures with dynamic references", file=sys.stderr
    )

    all_refs: list[tuple] = []
    all_edges: list[tuple[str, str]] = []

    # Build sheet name → sheet_id lookup for edge creation
    sheet_map: dict[str, int] = {}
    try:
        for row in conn.execute("SELECT sheet_id, sheet_name FROM sheets"):
            sheet_map[row[1].lower()] = row[0]
    except sqlite3.OperationalError:
        pass

    for proc_id, proc_name, body, module_name, kind, compile_branch in candidates:
        static_context = _build_static_context(conn, proc_id)
        full_name = f"{module_name}.{proc_name}"

        result = enrich_procedure(
            procedure_name=full_name,
            body=body,
            workbook_context=workbook_context,
            static_context=static_context,
        )

        if not result:
            continue

        metrics["procedures_submitted"] += 1
        metrics["total_latency_s"] += result.latency_s

        proc_binding_id = f"vba::{module_name}::{proc_name}::{kind}"
        if compile_branch:
            proc_binding_id += f"::{compile_branch}"

        for ref in result.refs:
            full_target = f"{ref.sheet}!{ref.target}"
            all_refs.append(
                (
                    proc_id,
                    ref.ref_kind,
                    "cell_range",
                    full_target,
                    "dynamic",  # Schema CHECK allows: exact, static_only, dynamic
                )
            )

            # Try to resolve to binding edge
            sheet_id = sheet_map.get(ref.sheet.lower())
            if sheet_id is not None and "cell_to_binding" in tables:
                # For range references, find overlapping bindings
                addr_match = re.match(
                    r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", ref.target, re.IGNORECASE
                )
                if addr_match:
                    from xl_marinade.core.vba.reference_extractor import _col_letter_to_num

                    r1 = int(addr_match.group(2))
                    c1 = _col_letter_to_num(addr_match.group(1))
                    r2 = int(addr_match.group(4)) if addr_match.group(4) else r1
                    c2 = _col_letter_to_num(addr_match.group(3)) if addr_match.group(3) else c1

                    try:
                        bindings = conn.execute(
                            """
                            SELECT DISTINCT ctb.binding_id
                            FROM cells c
                            JOIN cell_to_binding ctb ON c.cell_id = ctb.cell_id
                            WHERE c.sheet_id = ?
                              AND c.row BETWEEN ? AND ?
                              AND c.col BETWEEN ? AND ?
                        """,
                            (sheet_id, r1, r2, c1, c2),
                        ).fetchall()

                        for b in bindings:
                            bid = b[0]
                            if ref.ref_kind == "read":
                                all_edges.append((proc_binding_id, bid))
                            elif ref.ref_kind == "write":
                                all_edges.append((bid, proc_binding_id))
                    except sqlite3.OperationalError:
                        pass

        if result.description:
            metrics["descriptions_generated"] += 1

        print(
            f"    {full_name}: {len(result.refs)} refs inferred, "
            f"{result.latency_s:.1f}s, desc={'yes' if result.description else 'no'}",
            file=sys.stderr,
        )

    # Store refs
    if all_refs:
        if "vba_procedure_cell_refs" in tables:
            conn.executemany(
                """INSERT INTO vba_procedure_cell_refs
                   (procedure_id, ref_kind, target_kind, target, precision)
                   VALUES (?, ?, ?, ?, ?)""",
                all_refs,
            )
        metrics["refs_inferred"] = len(all_refs)

    # Store edges
    if all_edges:
        conn.executemany(
            "INSERT OR IGNORE INTO binding_edges (from_binding_id, to_binding_id, edge_count) VALUES (?, ?, 1)",
            all_edges,
        )
        metrics["edges_created"] = len(all_edges)

    conn.commit()
    return metrics
