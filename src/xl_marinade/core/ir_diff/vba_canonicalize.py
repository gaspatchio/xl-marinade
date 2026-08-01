# ABOUTME: VBA body canonicalizer for diff matching.
# ABOUTME: Produces a normalized form for change detection — the displayed source stays verbatim.

"""
VBA Canonicalization for diff matching.

Normalization rules (applied ONLY to the canonical form used by
normalized_body_hash and the matcher — displayed source in
CodeEvidence.source_chunks and vba_procedures.body is preserved verbatim):

1. Case-fold identifiers and keywords (VBA is case-insensitive)
   - String literals preserve case
   - Comments preserve case (but are stripped — see rule 2)
2. Strip comments (' and Rem)
3. Normalize line continuations (_ at EOL → join lines)
4. Collapse whitespace (runs of whitespace → single space)
5. Remove blank lines
6. Split statement separator : into separate lines
7. Rem → ' (unify before stripping)
8. Numeric literals: normalize 1000# → 1000, &H3E8 → 920 (canonical form only)
"""

from __future__ import annotations

import hashlib
import re


def canonicalize_vba_body(source: str) -> str:
    """
    Produce a canonical form of a VBA procedure body for change detection.

    The canonical form is used by the diff matcher to determine whether two
    procedure bodies are semantically identical (modulo formatting/comments).
    It is NOT used for display.

    Returns:
        Lowercased, comment-stripped, whitespace-collapsed, continuation-joined text.
    """
    # Step 1: Normalize line continuations (_ at EOL → join lines)
    # Must happen before line-by-line processing
    text = re.sub(r"\s+_\s*\n\s*", " ", source)

    lines = text.split("\n")
    output_lines: list[str] = []

    for line in lines:
        # Step 7: Unify Rem comments → ' style (so stripping catches both)
        stripped = line.lstrip()
        if stripped.lower().startswith("rem "):
            continue  # Skip Rem-only lines

        # Step 2: Strip Attribute lines (module metadata, not executable)
        if stripped.startswith("Attribute "):
            continue

        # Step 6: Split statement separator : into separate lines
        # But NOT inside strings. Simple heuristic: split on : not inside quotes.
        if ":" in line and '"' not in line:
            parts = [p.strip() for p in line.split(":") if p.strip()]
            for part in parts:
                processed = _process_line(part)
                if processed:
                    output_lines.append(processed)
            continue

        processed = _process_line(line)
        if processed:
            output_lines.append(processed)

    # Step 5: Remove blank lines (already filtered by _process_line)
    return "\n".join(output_lines)


def _process_line(line: str) -> str | None:
    """Process a single line: strip trailing comment, case-fold, collapse whitespace."""
    # Strip trailing comments (after ')
    # Be careful of strings: simple heuristic — find ' that's not inside quotes
    result = _strip_trailing_comment(line)

    # Skip if line is now empty or comment-only
    stripped = result.strip()
    if not stripped:
        return None
    if stripped.startswith("'"):
        return None

    # Step 1: Case-fold identifiers/keywords but preserve string literals
    result = _case_fold_preserving_strings(result)

    # Step 4: Collapse whitespace
    result = re.sub(r"\s+", " ", result).strip()

    # Step 8: Normalize numeric literals (canonical form only)
    result = _normalize_numerics(result)

    return result if result else None


def _strip_trailing_comment(line: str) -> str:
    """Strip trailing ' comment, respecting string literals."""
    in_string = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_string = not in_string
        elif ch == "'" and not in_string:
            return line[:i]
    return line


def _case_fold_preserving_strings(text: str) -> str:
    """Lowercase everything except content inside double-quoted strings."""
    parts = text.split('"')
    result = []
    for i, part in enumerate(parts):
        if i % 2 == 0:
            # Outside string — case-fold
            result.append(part.lower())
        else:
            # Inside string — preserve
            result.append(part)
    return '"'.join(result)


def _normalize_numerics(text: str) -> str:
    """Normalize VBA numeric literal forms to canonical representation."""

    # &H hex literals → decimal
    def _hex_to_dec(m: re.Match) -> str:
        try:
            return str(int(m.group(1), 16))
        except ValueError:
            return m.group(0)

    text = re.sub(r"&[Hh]([0-9A-Fa-f]+)&?", _hex_to_dec, text)

    # Type suffix: 1000# → 1000, 1000& → 1000, 1000! → 1000
    text = re.sub(r"(\d+)[#&!@%]", r"\1", text)

    return text


def canonical_body_hash(source: str) -> str:
    """Compute SHA-256 hash of the canonical form of a VBA body."""
    canonical = canonicalize_vba_body(source)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()
