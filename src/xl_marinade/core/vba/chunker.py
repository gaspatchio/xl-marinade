# ABOUTME: VBA procedure chunker — produces retrieval-ready chunks from VBAExtraction.
# ABOUTME: Phase 3 v1: one chunk per procedure. Future: split at block boundaries for long procs.

"""
VBA Chunker for retrieval-style LLM prompting.

Produces one chunk per procedure (v1 strategy). Each chunk carries metadata
for keyword-based retrieval: identifier tokens, comment tokens, referenced
cells, and called procedures.

Future refinement: procedures exceeding a configurable line cap would be split
at outermost block boundaries (If/For/With/Do/Select Case), never mid-statement.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from xl_marinade.core.vba.extractor import VBAExtraction


@dataclass
class VBAChunk:
    """A single retrieval-ready chunk of VBA source code."""

    procedure_name: str
    module_name: str
    procedure_kind: str
    procedure_compile_branch: str
    chunk_index: int
    chunk_text: str
    line_start: int
    line_end: int
    enclosing_block_kind: str | None  # None for whole-procedure chunks
    identifier_tokens: list[str]  # lowercased identifiers for keyword search
    comment_tokens: list[str]  # lowercased comment text tokens
    referenced_cells: list[str]  # cell addresses found in the source
    called_procedures: list[str]  # procedure names called from this chunk


# VBA keywords to exclude from identifier token lists (too common to be useful for search)
_VBA_NOISE_TOKENS = frozenset(
    {
        "dim",
        "as",
        "set",
        "let",
        "if",
        "then",
        "else",
        "elseif",
        "end",
        "sub",
        "function",
        "property",
        "get",
        "exit",
        "for",
        "next",
        "do",
        "loop",
        "while",
        "wend",
        "with",
        "select",
        "case",
        "to",
        "step",
        "each",
        "in",
        "and",
        "or",
        "not",
        "is",
        "nothing",
        "true",
        "false",
        "byval",
        "byref",
        "optional",
        "public",
        "private",
        "friend",
        "static",
        "const",
        "type",
        "enum",
        "new",
        "on",
        "error",
        "goto",
        "resume",
        "call",
        "me",
        "integer",
        "long",
        "double",
        "single",
        "string",
        "boolean",
        "variant",
        "object",
        "byte",
        "currency",
        "date",
        "decimal",
        "longlong",
        "longptr",
    }
)

# Regex patterns for extracting metadata from VBA source
_IDENTIFIER_RE = re.compile(r"\b([A-Za-z_]\w{2,})\b")
_COMMENT_RE = re.compile(r"'(.+)$", re.MULTILINE)
_CELL_REF_RE = re.compile(
    r"""(?:Range\s*\(\s*"([^"]+)"\s*\)|"""  # Range("A1:B10")
    r"""\[([A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?)\])""",  # [A1] or [A1:B10]
    re.IGNORECASE,
)
_PROC_CALL_RE = re.compile(r"\b([A-Z_]\w+)\s*(?:\(|$)", re.MULTILINE)


def _extract_identifier_tokens(source: str) -> list[str]:
    """Extract meaningful identifier tokens from VBA source."""
    raw = set()
    for m in _IDENTIFIER_RE.finditer(source):
        token = m.group(1).lower()
        if token not in _VBA_NOISE_TOKENS and len(token) > 2:
            raw.add(token)
    return sorted(raw)


def _extract_comment_tokens(source: str) -> list[str]:
    """Extract tokens from VBA comments."""
    tokens = set()
    for m in _COMMENT_RE.finditer(source):
        comment = m.group(1).strip()
        for word in re.split(r"\W+", comment):
            w = word.lower()
            if len(w) > 2:
                tokens.add(w)
    return sorted(tokens)


def _extract_cell_references(source: str) -> list[str]:
    """Extract cell address references from VBA source."""
    refs = set()
    for m in _CELL_REF_RE.finditer(source):
        ref = m.group(1) or m.group(2)
        if ref:
            refs.add(ref)
    return sorted(refs)


def _extract_called_procedures(source: str, known_names: set[str]) -> list[str]:
    """Extract names of called procedures (from the known procedure set)."""
    called = set()
    for m in _PROC_CALL_RE.finditer(source):
        name = m.group(1)
        if name.lower() in known_names:
            called.add(name)
    return sorted(called)


def chunk_extraction(extraction: VBAExtraction) -> list[VBAChunk]:
    """
    Produce retrieval-ready chunks from a VBAExtraction.

    Phase 3 v1 strategy: one chunk per procedure, no block-boundary splitting.
    Each chunk carries keyword metadata for BM25-style retrieval.

    Args:
        extraction: Complete VBA extraction result

    Returns:
        List of VBAChunk objects ready for storage in vba_chunks table
    """
    # Build set of known procedure names (lowercased) for call detection
    known_names = {p.name.lower() for p in extraction.procedures}

    chunks: list[VBAChunk] = []
    for proc in extraction.procedures:
        body = proc.body
        if not body.strip():
            continue

        chunk = VBAChunk(
            procedure_name=proc.name,
            module_name=proc.module_name,
            procedure_kind=proc.kind,
            procedure_compile_branch=proc.compile_branch,
            chunk_index=0,  # v1: always 0 (one chunk per procedure)
            chunk_text=body,
            line_start=proc.line_start,
            line_end=proc.line_end,
            enclosing_block_kind=None,  # whole-procedure chunk
            identifier_tokens=_extract_identifier_tokens(body),
            comment_tokens=_extract_comment_tokens(body),
            referenced_cells=_extract_cell_references(body),
            called_procedures=_extract_called_procedures(body, known_names),
        )
        chunks.append(chunk)

    return chunks
