# ABOUTME: Simple one-pass labelling algorithm for Sprint 1
# ABOUTME: Selects best label candidate from Phase 1 IR using heuristic scoring

import json
import logging
import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path

from xl_marinade.core.db_uri import connect_read_only

from .header_context import build_header_context
from .mutation_engine import MutationLogger, replay_mutations
from .overlay_database import write_overlay_to_db
from .utils.validation import is_valid_label_candidate

logger = logging.getLogger(__name__)


@dataclass
class Binding:
    """Simplified binding representation from Phase 1 IR."""

    binding_id: str
    sheet: str
    address_a1: str
    debug_label: str | None
    label_candidates_json: str


def is_numeric(text: str) -> bool:
    """Check if text is numeric."""
    if not text:
        return False
    # Remove common numeric decorators
    cleaned = text.strip().replace(",", "").replace("$", "").replace("%", "")
    try:
        float(cleaned)
        return True
    except ValueError:
        return False


def extract_text_from_formula(formula: str) -> str | None:
    """
    Extract meaningful text from a formula string.

    Some Excel cells contain formulas that generate labels, like:
    - ="Total Cash In ="&Z14&"+"&AA14  → "Total Cash In"
    - ="Policies inforce"              → "Policies inforce"
    - ="Net Cash Flow: "&TEXT(X12,...)  → "Net Cash Flow"

    This is useful when all cells in a scan range contain formulas
    but some formulas have embedded text that would be good labels.

    Args:
        formula: Formula string (starting with '=')

    Returns:
        Extracted text if found and meaningful, None otherwise
    """
    if not formula or not formula.startswith("="):
        return None

    # Pattern: ="text" at the start (full string formula or concatenation start)
    # Matches: ="Total Cash In =" or ="Policies inforce"
    match = re.match(r'^="([^"]+)"', formula)
    if match:
        text = match.group(1).strip()
        # Filter out pure formatting patterns (parentheses, operators, etc.)
        if text and not re.match(r"^[\(\)\s\=\+\-\*\/\&\:\#\,\.\d]+$", text):
            # Remove trailing punctuation like " =" or ": " or " - "
            text = re.sub(r"[\s\:\=\-]+$", "", text)
            if len(text) >= 2:  # Minimum meaningful length
                return text

    return None


def clean_label(label: str) -> str:
    """
    Clean and normalize selected label.

    Rules:
    - Strip leading/trailing whitespace
    - Collapse multiple spaces to single space
    - Remove control characters
    - Limit length to 200 characters

    Args:
        label: Raw label string

    Returns:
        Cleaned label string
    """
    # Strip whitespace
    label = label.strip()

    # Collapse multiple spaces
    label = re.sub(r"\s+", " ", label)

    # Remove control characters (keep only printable)
    label = "".join(char for char in label if char.isprintable())

    # Limit length
    if len(label) > 200:
        label = label[:197] + "..."

    # Fallback if empty
    if not label:
        label = "Unlabelled"

    return label


# Cycle 17 #434: short letters+digits codes ("INV001", "POL-123") are row
# IDs, not descriptive labels. Used by select_best_literal to skip past an
# ID cell sitting between the data and the real header.
_ID_CODE_RE = re.compile(r"^[A-Za-z]{1,6}[\s_-]?\d{1,6}[A-Za-z]?$")


def _looks_like_id_code(text: str) -> bool:
    return bool(_ID_CODE_RE.match(text.strip()))


def select_best_literal(literals: list[str], candidate_type: str | None = None) -> str:
    """
    Select best literal from candidate's literals list.

    Heuristics:
    1. Prefer non-numeric text
    2. Filter out formulas (starting with =), BUT extract embedded text from formulas
    3. For scan types, prioritize the closest (first) valid text literal
    4. Otherwise, prefer longer meaningful text or combine

    Args:
        literals: List of literal values from candidate cells
        candidate_type: Type of candidate (e.g., 'scan_above', 'scan_left')

    Returns:
        Selected literal string
    """
    # Keep original literals for formula text extraction fallback
    original_literals = literals

    # Remove blanks, None, and FORMULAS
    # We filter formulas here to prevent "Formula Pollution" in labels
    literals = [
        lit for lit in literals if lit and str(lit).strip() and not str(lit).strip().startswith("=")
    ]

    if not literals:
        # No regular text literals found - try extracting text from formulas
        # This handles cases where header cells contain formulas like ="Total Cash In ="&...
        for lit in original_literals:
            if lit and str(lit).strip().startswith("="):
                extracted = extract_text_from_formula(str(lit))
                if extracted and is_valid_label_candidate(extracted):
                    return extracted
        return "Unlabelled"

    # If all numeric, use first
    if all(is_numeric(str(lit)) for lit in literals):
        return str(literals[0])

    # Filter to text only
    text_literals = [str(lit) for lit in literals if not is_numeric(str(lit))]

    if not text_literals:
        # Fallback to first numeric if no text found
        return str(literals[0])

    # If single text literal, use it
    if len(text_literals) == 1:
        return text_literals[0]

    # NEW LOGIC: For scan candidates, prioritize the closest header
    # In evidence.py, scan results are ordered closest-to-furthest (literals[0] is closest)
    # unless they were sorted?
    # Checking evidence.py:
    #   scan_left: range(start_col - 1, end_col - 1, -1) -> appended closest first
    #   scan_above: range(start_row - 1, end_row - 1, -1) -> appended closest first
    #   BUT then: literals=literals (passed as is)
    #   AND cells=sorted(cells_data) (sorted by position)
    #   The 'literals' field usually preserves scan order.

    if candidate_type in (
        "scan_above",
        "scan_left",
        "scan_right",
        "scan_below",
        "table_header_row",
        "row_segment",
        "col_segment",
    ):
        # Return the first (closest) valid text literal
        # This avoids "Greedy Concatenation" like "month smoking status gender..."
        # Cycle 17 #434: except when that literal is a row-ID code (e.g. an
        # asset id "INV001" in the column between the data and the header) —
        # prefer the nearest non-ID literal, falling back to the ID code
        # when it is the only text available.
        # H6: skip row-ID codes (#434) AND junk tokens (footnotes, bare
        # addresses, GUIDs, numerics) so the walk lands on the nearest
        # *descriptive* header (e.g. past "(2)" onto "Net Amount @ Risk").
        for lit in text_literals:
            if not _looks_like_id_code(lit) and is_valid_label_candidate(lit):
                return lit
        # No descriptive header found — fall back to the prior behaviour
        # (nearest non-ID literal, else the closest literal).
        for lit in text_literals:
            if not _looks_like_id_code(lit):
                return lit
        return text_literals[0]

    # Multiple text literals (non-scan): combine with space
    combined = " ".join(text_literals)

    # If combined is reasonable length, use it
    if len(combined) <= 100:
        return combined

    # Otherwise, use longest single literal
    return max(text_literals, key=len)


def _parse_a1_address(address: str) -> tuple[int, int] | None:
    """
    Parse A1 notation to (col_index, row_index).

    Args:
        address: A1 address like "B5" or "Sheet1!B5"

    Returns:
        Tuple of (col, row) as 0-based indices, or None if unparseable
    """
    import re

    # Remove sheet prefix if present
    if "!" in address:
        address = address.split("!")[-1]

    # Match column letters and row numbers
    match = re.match(r"^([A-Z]+)(\d+)$", address.upper())
    if not match:
        return None

    col_letters, row_str = match.groups()

    # Convert column letters to number (A=0, B=1, ... Z=25, AA=26, etc.)
    col_num = 0
    for char in col_letters:
        col_num = col_num * 26 + (ord(char) - ord("A") + 1)
    col_num -= 1  # Make 0-based

    row_num = int(row_str) - 1  # Make 0-based

    return (col_num, row_num)


def calculate_distance(addr1: str, addr2: str) -> int:
    """
    Calculate Manhattan distance between two A1 addresses.

    Args:
        addr1: First address (e.g., "Sheet1!B5" or "B5")
        addr2: Second address (e.g., "Sheet1!B10" or "B10")

    Returns:
        Manhattan distance in cells (sum of row and column differences)
    """
    try:
        pos1 = _parse_a1_address(addr1)
        pos2 = _parse_a1_address(addr2)

        if pos1 is None or pos2 is None:
            return 999  # Large distance for unparseable addresses

        col1, row1 = pos1
        col2, row2 = pos2

        # Manhattan distance: |col1 - col2| + |row1 - row2|
        return abs(col1 - col2) + abs(row1 - row2)
    except Exception:
        return 999  # Large distance on any error


def score_candidate(candidate: dict, binding: Binding) -> float:
    """
    Score a label candidate (higher = better).

    Updated in Sprint 2 Story 1 to provide calibrated confidence scores.

    Heuristics:
    - Candidate type priority (granular scores)
    - Text quality (length, non-numeric)
    - Multi-cell context
    - Proximity for scan types

    Args:
        candidate: Label candidate dict from IR
        binding: Binding being labelled

    Returns:
        Score 0.0 to 1.0
    """
    score = 0.0

    # Rule 1: Candidate type priority (Sprint 2: calibrated, granular scores)
    type_scores = {
        "named_exact": 0.95,  # Exact name match - highest confidence
        "named_superset": 0.90,  # Name fully contains binding - high confidence
        "named_subset": 0.95,  # Named range subset - highest confidence
        "table_column_headers": 0.85,  # Table column header - high confidence
        "table_header_row": 0.80,  # Table header row - high confidence
        "sheet_name": 0.45,  # H6: below scan_left (0.55) so a valid adjacent header wins (was 0.60)
        "merged_header": 0.70,  # Merged cell header - medium-high
        "merged_left": 0.70,  # Merged header (left) - medium-high
        "merged_above": 0.70,  # Merged header (above) - medium-high
        "scan_above": 0.65,  # Adjacent text above - medium-high (raised from 0.60)
        "scan_left": 0.55,  # Adjacent text left - medium (raised from 0.50)
        "scan_below": 0.60,  # Name rank 4: header in the row BELOW a 1x1 scalar
        # (emitted by evidence.py only when above/left have no
        # text header); above sheet_name, below scan_above.
        "row_segment": 0.60,  # Axis row segment - medium
        "col_segment": 0.60,  # Axis col segment - medium
    }

    candidate_type = candidate.get("type")
    base_score = type_scores.get(candidate_type, 0.1)  # Default 0.1 for unknown types

    # Log warning for unrecognized types
    if candidate_type and candidate_type not in type_scores:
        logger.warning(f"Unrecognized candidate type: {candidate_type}. Using default score 0.1.")

    score += base_score

    # Rule 2: Text quality
    literals = candidate.get("literals", [])
    best_literal = None
    if literals:
        # Pass candidate type to helper
        best_literal = select_best_literal(literals, candidate_type)

        # Prefer text over numbers
        if not is_numeric(best_literal):
            score += 0.2

        # Prefer reasonable length (5-50 chars)
        if 5 <= len(best_literal) <= 50:
            score += 0.1

        # Penalize very short (< 3 chars) — EXCEPT a valid, non-numeric short literal
        # used as an ADJACENT header (scan_left/scan_above). Domain symbols like K, S0,
        # T, q, rf ARE the variable name; the bare <3 penalty sinks them below the 0.75
        # sheet-name fallback, so the labeller picks the sheet name over the real header.
        is_short_adjacent_header = (
            candidate_type in ("scan_left", "scan_above")
            and not is_numeric(best_literal)
            and is_valid_label_candidate(best_literal)
        )
        if len(best_literal) < 3 and not is_short_adjacent_header:
            score -= 0.2

        # Penalize very long (> 100 chars)
        if len(best_literal) > 100:
            score -= 0.1

    # Rule 3: Multi-cell candidates (more context)
    cells = candidate.get("cells", [])
    num_cells = len(cells)
    if num_cells > 1:
        score += 0.1

    # Rule 4: Proximity for scan types. A real header sits ADJACENT to the binding; a
    # label grabbed from FAR away (across blank cells or a populated data row) is almost
    # never the header. So distance is a graduated signal — bonus when adjacent, PENALTY
    # when far — measured at the cell of the SELECTED literal, relative to the binding's
    # TOP-LEFT cell. Two bugs made this misfire before:
    #   (a) the distance was taken to the NEAREST cell in the scan window (often a blank
    #       between the binding and a far stray token like 'x'), not the selected literal;
    #   (b) calculate_distance was passed the binding's whole RANGE address ('H9:H993'),
    #       which it cannot parse -> 999 -> proximity silently never fired for vectors.
    # Fixing both lets a true adjacent header (a model's 'Date' header at H8, d=1) beat a distant
    # scan_left that walked across a data row ('Issue age' 6 cols away), and sinks a stray
    # 'x' six rows up below the real scan_below header ('Investment Margin').
    if candidate.get("type") in ("scan_above", "scan_left") and cells:
        anchor = binding.address_a1.split("!")[-1].split(":")[0]  # top-left cell
        has_values = any("value" in c for c in cells)
        sel_addr = None
        if has_values and best_literal:
            for cell in cells:
                cv = cell.get("value")
                if cv is not None and str(cv).strip() == str(best_literal).strip():
                    sel_addr = cell.get("address", "")
                    break
        if sel_addr:
            dist = calculate_distance(sel_addr, anchor)
        else:
            # no value info (synthetic candidates) -> nearest-cell distance
            dist = min(
                (
                    calculate_distance(c.get("address", ""), anchor)
                    for c in cells
                    if c.get("address")
                ),
                default=999,
            )
        if dist <= 2:
            score += 0.05
        elif dist >= 5:
            score -= 0.15

    # A numeric literal is a VALUE, not a header — cap it below the sheet-name fallback
    # so a scan that merely walked up/left into the data and grabbed a number ('100',
    # '0.4') never wins as a label. Bare numbers stay VALID per H6 (so a lone numeric
    # candidate is still selectable when nothing else exists), but they must not be
    # PREFERRED over a real header or the sheet-name fallback.
    if best_literal is not None and is_numeric(best_literal):
        return min(score, 0.30)

    return min(score, 1.0)  # Cap at 1.0


def generate_fallback_label(binding: Binding) -> str:
    """
    Generate fallback label when no candidates available.

    Pattern: Use debug_label if meaningful, else "sheet!address"

    Args:
        binding: Binding without label candidates

    Returns:
        Fallback label string
    """
    # Use debug_label if available and meaningful
    # Reject binding ID patterns (contain :: anywhere in string)
    if binding.debug_label and "::" not in binding.debug_label:
        return binding.debug_label

    # Otherwise use sheet + address
    # address_a1 already contains sheet prefix, so strip it to avoid duplication
    address = binding.address_a1
    if "!" in address:
        # Strip sheet prefix from address_a1 (e.g., "Data!A1:A10" -> "A1:A10")
        address = address.split("!", 1)[1]

    return f"{binding.sheet}!{address}"


def extract_label_text(candidate: dict) -> str:
    """
    Extract label text from candidate.

    Args:
        candidate: Label candidate dict from IR

    Returns:
        Label text string
    """
    literals = candidate.get("literals", [])
    candidate_type = candidate.get("type")

    if not literals:
        return "Unlabelled"

    return select_best_literal(literals, candidate_type)


def simple_label_selection(binding: Binding) -> tuple[str, float, str]:
    """
    Select best label from IR candidates using simple heuristics.

    Args:
        binding: Binding to label (from Phase 1 IR)

    Returns:
        Tuple of (Selected label string, Score 0-1, Candidate type)
    """
    # Parse label_candidates_json
    try:
        candidates_data = json.loads(binding.label_candidates_json)
        # Handle case where JSON is "null" (becomes None)
        if not isinstance(candidates_data, dict):
            return generate_fallback_label(binding), 0.0, "fallback"
        label_candidates = candidates_data.get("label_candidates", [])
    except (json.JSONDecodeError, KeyError, TypeError):
        # Malformed candidates - use fallback
        return generate_fallback_label(binding), 0.0, "fallback"

    if not label_candidates:
        return generate_fallback_label(binding), 0.0, "fallback"

    # Score each candidate
    scored_candidates = []
    for cand in label_candidates:
        # Validation check: reject invalid candidates (e.g. formula strings)
        literals = cand.get("literals", [])
        candidate_type = cand.get("type")

        if not literals:
            continue

        # Check if the extracted text would be valid
        # Pass candidate_type to ensure consistency with scoring logic
        text = select_best_literal(literals, candidate_type)

        # Reject "Unlabelled" - it's a fallback marker, not a real label
        if text == "Unlabelled":
            continue

        if not is_valid_label_candidate(text):
            # Skip this candidate
            continue

        scored_candidates.append((score_candidate(cand, binding), cand))

    if not scored_candidates:
        return generate_fallback_label(binding), 0.0, "fallback"

    # Sort by score (highest first)
    scored_candidates.sort(key=lambda x: x[0], reverse=True)

    # Select best candidate
    best_score, best_candidate = scored_candidates[0]

    # Name rank 2: a code-style defined-name slug (named_*, e.g. 'mtx_UW_Mortality_Table')
    # outscores adjacent text headers (named_superset 0.90 > scan_above 0.65) but never
    # matches a human-authored variable name — whereas the genuine adjacent header on the
    # sheet ('UW Mortality Table', 'Age', 'Premium') usually does. When the winner is a
    # named_* slug AND a valid non-numeric scan_above/scan_left header is available, prefer
    # that header. Gated on the header actually existing (non-numeric + valid) so a binding
    # whose only real label IS its defined name keeps it.
    if best_candidate.get("type") in ("named_exact", "named_superset", "named_subset"):
        for s, cand in scored_candidates:
            ct = cand.get("type")
            if ct in ("scan_above", "scan_left"):
                lit = select_best_literal(cand.get("literals", []), ct)
                if lit and not is_numeric(lit) and is_valid_label_candidate(lit):
                    best_score, best_candidate = s, cand
                    break

    # Extract label text
    label = extract_label_text(best_candidate)

    # Clean and normalize
    label = clean_label(label)

    return label, best_score, best_candidate.get("type", "unknown")


def backfill_binding_labels(conn: sqlite3.Connection, *, dry_run: bool = False) -> tuple[int, int]:
    """Populate ``bindings.label`` for every binding via the deterministic labeller.

    The extractor leaves ``bindings.label`` NULL; this fills it from each
    binding's ``spatial_candidates`` blob (Phase-1.5 Lever C). Runs as a
    stage inside ``run_full_workbook_extraction`` so a fresh ir.db is
    label-complete — every Cycle-17 ad-hoc re-extraction that skipped the
    standalone backfill script silently shipped all-NULL labels across
    several real-model regressions, degrading compare digests into one-sided
    added/removed noise.

    Returns ``(real_labels, fallback_labels)``. Idempotent: re-running
    overwrites with the same deterministic selection.
    """
    # header_context_json is added by newer schemas; ALTER-guard so backfill also works on
    # DBs built before the column existed.
    cols = {r[1] for r in conn.execute("PRAGMA table_info(bindings)")}
    if "header_context_json" not in cols:
        conn.execute("ALTER TABLE bindings ADD COLUMN header_context_json TEXT")

    rows = conn.execute(
        """
        SELECT b.binding_id, s.sheet_name,
               s.sheet_name || '!' || b.address_a1 AS addr,
               jsc.json AS spatial_candidates
        FROM bindings b
        JOIN sheets s ON b.sheet_id = s.sheet_id
        LEFT JOIN json_blobs jsc ON b.spatial_candidates_blob_id = jsc.blob_id
        """
    ).fetchall()

    real = 0
    fallback = 0
    updates: list[tuple[str, str, str]] = []
    for binding_id, sheet_name, addr, spatial in rows:
        binding = Binding(
            binding_id=binding_id,
            sheet=sheet_name,
            address_a1=addr,
            debug_label=None,
            label_candidates_json=spatial if spatial else "{}",
        )
        label, _score, source = simple_label_selection(binding)
        if source == "fallback":
            fallback += 1
        else:
            real += 1
        # Derive the header hierarchy from the SAME candidates, pinning primary == label.
        try:
            cands = json.loads(spatial).get("label_candidates", []) if spatial else []
        except (json.JSONDecodeError, TypeError, AttributeError):
            cands = []
        layers = build_header_context(addr, cands, label, source)
        updates.append((label, json.dumps(layers, ensure_ascii=False), binding_id))

    if not dry_run and updates:
        conn.executemany(
            "UPDATE bindings SET label = ?, header_context_json = ? WHERE binding_id = ?",
            updates,
        )
        conn.commit()

    return real, fallback


def label_all_bindings(
    ir_db_path: str, output_mutations_path: str, output_overlay_path: str
) -> int:
    """
    Label all bindings in IR and generate mutations + overlay database.

    Args:
        ir_db_path: Path to Phase 1 IR database
        output_mutations_path: Where to write mutations.json
        output_overlay_path: Where to write semantic_overlay.db

    Returns:
        Number of bindings labelled
    """
    # Load IR database
    # Validate path exists before opening
    if not Path(ir_db_path).exists():
        raise FileNotFoundError(f"IR database not found: {ir_db_path}")

    ir_conn = connect_read_only(ir_db_path)

    # Get all bindings from agent view
    rows = ir_conn.execute("""
        SELECT binding_id, sheet, address, spatial_candidates
        FROM agent_bindings
    """).fetchall()

    bindings = [
        Binding(
            binding_id=row[0],
            sheet=row[1],
            address_a1=row[2],
            debug_label=None,  # Not available in agent view
            label_candidates_json=row[3] if row[3] else "{}",
        )
        for row in rows
    ]

    ir_conn.close()

    # Initialize mutation logger
    mutation_logger = MutationLogger()

    # Label each binding
    labelled_count = 0
    for binding in bindings:
        try:
            # Select label
            label, score, best_type = simple_label_selection(binding)

            # Determine candidate type and knowledge source for reasoning
            if best_type == "fallback":
                knowledge_source = "fallback"
                reasoning = "No label candidates available or parsing error, using fallback"
            else:
                knowledge_source = "ir_candidates"
                reasoning = f"Selected from IR candidates (type: {best_type})"

            # Log mutation
            mutation_logger.set_label(
                binding_id=binding.binding_id,
                old_label=None,
                new_label=label,
                reasoning=reasoning,
                knowledge_source=knowledge_source,
                confidence=score,
            )

            labelled_count += 1

        except Exception as e:
            logger.warning(f"Failed to label {binding.binding_id}: {e}")
            # Use fallback on error
            mutation_logger.set_label(
                binding_id=binding.binding_id,
                old_label=None,
                new_label=generate_fallback_label(binding),
                reasoning=f"Fallback after error: {e}",
                knowledge_source="fallback",
                confidence=0.0,
            )
            labelled_count += 1

    # Save mutations
    mutation_logger.save(output_mutations_path)

    # Replay mutations to generate overlay
    overlay = replay_mutations(ir_db_path, output_mutations_path)

    # Write overlay database
    write_overlay_to_db(overlay, output_mutations_path, ir_db_path, output_overlay_path)

    return labelled_count
