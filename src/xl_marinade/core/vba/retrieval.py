# ABOUTME: VBA chunk retrieval — keyword-based search over vba_chunks table.
# ABOUTME: BM25-style LIKE matching on identifier and comment tokens, with same-procedure reranking.

"""
VBA Chunk Retrieval for the reasoning pipeline.

Queries the vba_chunks table using keyword matching (consistent with the
existing semantic index's BM25-style approach). Returns chunks ranked by
relevance to a concept query, with a soft preference for grouping chunks
from the same procedure.

Used by the executor's FIND_PROCEDURE action to provide relevant code
excerpts when answering questions about VBA procedures.
"""

from __future__ import annotations

import json
import logging
import re
import sqlite3
from dataclasses import dataclass

logger = logging.getLogger(__name__)


@dataclass
class RetrievedChunk:
    """A chunk retrieved from the vba_chunks table with a relevance score."""

    chunk_id: int
    procedure_id: int
    procedure_name: str
    module_name: str
    chunk_index: int
    chunk_text: str
    line_start: int
    line_end: int
    enclosing_block_kind: str | None
    score: float  # 0-1 relevance score


def retrieve_chunks(
    conn: sqlite3.Connection,
    query: str,
    *,
    max_chunks: int = 5,
    max_tokens: int = 4000,
) -> list[RetrievedChunk]:
    """
    Retrieve VBA chunks matching a query string.

    Uses keyword matching against identifier_tokens_json and comment_tokens_json
    stored in vba_chunks. Scores chunks by the number of matching query tokens.
    Applies a same-procedure reranking boost to prefer grouped chunks.

    Args:
        conn: SQLite connection to the ir.db
        query: Search query (procedure name, concept, or question text)
        max_chunks: Maximum chunks to return (default 5)
        max_tokens: Approximate token budget for combined chunk text (~4 chars/token)

    Returns:
        List of RetrievedChunk sorted by descending score
    """
    # Tokenize the query
    query_tokens = set()
    for word in re.split(r"[\s._\-]+", query.lower()):
        if len(word) > 2:
            query_tokens.add(word)

    if not query_tokens:
        return []

    # Check if vba_chunks table exists
    try:
        conn.execute("SELECT 1 FROM vba_chunks LIMIT 1")
    except sqlite3.OperationalError:
        logger.debug("vba_chunks table not found")
        return []

    # Fetch all chunks with their metadata
    try:
        rows = conn.execute("""
            SELECT
                c.chunk_id,
                c.procedure_id,
                p.name AS procedure_name,
                m.name AS module_name,
                c.chunk_index,
                c.chunk_text,
                c.line_start,
                c.line_end,
                c.enclosing_block_kind,
                c.identifier_tokens_json,
                c.comment_tokens_json
            FROM vba_chunks c
            JOIN vba_procedures p ON c.procedure_id = p.procedure_id
            JOIN vba_modules m ON p.module_id = m.module_id
        """).fetchall()
    except sqlite3.Error as e:
        logger.warning("Failed to query vba_chunks: %s", e)
        return []

    if not rows:
        return []

    # Score each chunk by token overlap
    scored: list[tuple[float, RetrievedChunk]] = []
    for row in rows:
        (
            chunk_id,
            proc_id,
            proc_name,
            mod_name,
            chunk_idx,
            chunk_text,
            line_start,
            line_end,
            block_kind,
            id_json,
            comment_json,
        ) = row

        try:
            id_tokens = set(json.loads(id_json)) if id_json else set()
        except (json.JSONDecodeError, TypeError):
            id_tokens = set()
        try:
            comment_tokens = set(json.loads(comment_json)) if comment_json else set()
        except (json.JSONDecodeError, TypeError):
            comment_tokens = set()

        # Score: identifier match (weight 2) + comment match (weight 1) + name match (weight 3)
        id_matches = len(query_tokens & id_tokens)
        comment_matches = len(query_tokens & comment_tokens)
        name_tokens = set(re.split(r"[\s._]+", proc_name.lower()))
        name_matches = len(query_tokens & name_tokens)

        raw_score = name_matches * 3.0 + id_matches * 2.0 + comment_matches * 1.0
        if raw_score <= 0:
            continue

        # Normalize to 0-1
        max_possible = len(query_tokens) * 3.0  # if all tokens matched name
        score = min(raw_score / max(max_possible, 1.0), 1.0)

        scored.append(
            (
                score,
                RetrievedChunk(
                    chunk_id=chunk_id,
                    procedure_id=proc_id,
                    procedure_name=proc_name,
                    module_name=mod_name,
                    chunk_index=chunk_idx,
                    chunk_text=chunk_text,
                    line_start=line_start,
                    line_end=line_end,
                    enclosing_block_kind=block_kind,
                    score=round(score, 3),
                ),
            )
        )

    # Sort by score descending
    scored.sort(key=lambda x: -x[0])

    # Apply same-procedure reranking: boost chunks from procedures that
    # already have a top-ranked chunk (prefer grouping over scattering)
    if len(scored) > max_chunks:
        top_proc_ids = {scored[i][1].procedure_id for i in range(min(3, len(scored)))}
        for i in range(3, len(scored)):
            if scored[i][1].procedure_id in top_proc_ids:
                scored[i] = (scored[i][0] + 0.1, scored[i][1])
        scored.sort(key=lambda x: -x[0])

    # Apply token budget
    result: list[RetrievedChunk] = []
    total_chars = 0
    char_budget = max_tokens * 4  # rough: 4 chars per token

    for score, chunk in scored[: max_chunks * 2]:  # fetch extra for budget filtering
        if len(result) >= max_chunks:
            break
        chunk_chars = len(chunk.chunk_text)
        if total_chars + chunk_chars > char_budget and result:
            break  # don't exceed budget (but always include at least 1)
        result.append(chunk)
        total_chars += chunk_chars

    return result
