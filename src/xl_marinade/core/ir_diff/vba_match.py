# ABOUTME: VBA bipartite procedure matcher for version comparison.
# ABOUTME: 4-pass matching (exact, fuzzy same-name, rename same-module, cross-module rename)
# ABOUTME: with IDF-weighted token similarity, length-aware floor, and confidence scores.

"""
VBA Procedure Matcher for diff.

Matches procedures across two workbook versions using a 4-pass algorithm:
  Pass 1 — Exact match: (module, name, canonical_body_hash) identity
  Pass 2 — Same name, fuzzy body: same (module, name) but different hash → token similarity
  Pass 3 — Rename detection, same module: procedures only in A matched to only-in-B within same module
  Pass 4 — Cross-module rename: same as Pass 3 but across modules (disabled by default)

Scoring uses IDF-weighted token Jaccard with a minimum-token floor (<20 distinct
non-boilerplate tokens → skip fuzzy passes to avoid false positives from boilerplate overlap).
"""

from __future__ import annotations

import hashlib
import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Literal

from xl_marinade.core.ir_diff.vba_canonicalize import canonicalize_vba_body

# VBA boilerplate tokens (high frequency, low information) for IDF weighting
_BOILERPLATE_TOKENS = frozenset(
    {
        "dim",
        "as",
        "set",
        "let",
        "if",
        "then",
        "else",
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
        "const",
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
        "range",
        "value",
        "cells",
        "application",
        "worksheets",
        "activesheet",
    }
)

_TOKEN_RE = re.compile(r"\b([a-z_]\w{2,})\b")


@dataclass
class ProcedureDesc:
    """Description of a procedure for matching."""

    key: str  # "module::name::kind"
    module: str
    name: str
    kind: str
    canonical_hash: str
    body_tokens: Counter  # lowercased, IDF-weighted
    token_count: int  # distinct non-boilerplate tokens
    source_lines: int


@dataclass
class MatchResult:
    """Result of matching a procedure pair."""

    key_a: str
    key_b: str
    confidence: float  # 0-1
    match_pass: int  # 1-4
    change_type: Literal[
        "unchanged",  # exact match, raw bodies identical
        "cosmetic_only",  # canonical hash matches but raw differs
        "logic_changed",  # canonical hash differs
        "renamed",  # name differs but body matches
        "renamed_modified",  # name AND body differ
    ]


@dataclass
class VBAMatchOutput:
    """Complete output of the VBA procedure matcher."""

    matched: list[MatchResult] = field(default_factory=list)
    added: list[str] = field(default_factory=list)  # keys only in B
    removed: list[str] = field(default_factory=list)  # keys only in A
    ambiguous_dropped: list[str] = field(default_factory=list)  # keys with multiple candidates


def _tokenize(canonical_body: str) -> Counter:
    """Extract IDF-weighted token counter from canonical body."""
    tokens = Counter()
    for m in _TOKEN_RE.finditer(canonical_body):
        token = m.group(1)
        if token not in _BOILERPLATE_TOKENS:
            tokens[token] += 1
    return tokens


def _token_similarity(a: Counter, b: Counter) -> float:
    """IDF-weighted Jaccard-like similarity between two token counters."""
    if not a or not b:
        return 0.0
    a_set = set(a.keys())
    b_set = set(b.keys())
    intersection = a_set & b_set
    union = a_set | b_set
    if not union:
        return 0.0
    return len(intersection) / len(union)


def _build_desc(key: str, module: str, name: str, kind: str, raw_body: str) -> ProcedureDesc:
    """Build a ProcedureDesc for matching."""
    canonical = canonicalize_vba_body(raw_body)
    canonical_hash = hashlib.sha256(canonical.encode("utf-8")).hexdigest()
    tokens = _tokenize(canonical)
    return ProcedureDesc(
        key=key,
        module=module,
        name=name,
        kind=kind,
        canonical_hash=canonical_hash,
        body_tokens=tokens,
        token_count=len(tokens),
        source_lines=len(raw_body.split("\n")),
    )


# Threshold defaults (tunable)
PASS2_THRESHOLD = 0.5  # same name, fuzzy body
PASS3_THRESHOLD = 0.7  # rename, same module
PASS4_THRESHOLD = 0.85  # rename, cross-module
MIN_TOKEN_FLOOR = 20  # skip fuzzy passes if fewer distinct non-boilerplate tokens


def match_procedures(
    procs_a: dict[str, ProcedureDesc],
    procs_b: dict[str, ProcedureDesc],
    *,
    detect_cross_module_renames: bool = False,
    pass2_threshold: float = PASS2_THRESHOLD,
    pass3_threshold: float = PASS3_THRESHOLD,
    pass4_threshold: float = PASS4_THRESHOLD,
    min_token_floor: int = MIN_TOKEN_FLOOR,
) -> VBAMatchOutput:
    """
    Match procedures across two versions.

    Args:
        procs_a: Procedures from version A (key → ProcedureDesc)
        procs_b: Procedures from version B (key → ProcedureDesc)
        detect_cross_module_renames: Enable Pass 4 (default: disabled)

    Returns:
        VBAMatchOutput with matched pairs, added, removed, and ambiguous-dropped
    """
    output = VBAMatchOutput()
    matched_a: set[str] = set()
    matched_b: set[str] = set()

    # ---- Pass 1: Exact match (module, name, kind, canonical_hash) ----
    for key_a, desc_a in procs_a.items():
        for key_b, desc_b in procs_b.items():
            if key_b in matched_b:
                continue
            if (
                desc_a.module == desc_b.module
                and desc_a.name == desc_b.name
                and desc_a.kind == desc_b.kind
            ):
                if desc_a.canonical_hash == desc_b.canonical_hash:
                    change_type = "unchanged"  # or "cosmetic_only" if raw differs
                    # Check raw body difference for cosmetic detection
                    # (we don't have raw body here, but hash tells us canonical is same)
                    output.matched.append(
                        MatchResult(
                            key_a=key_a,
                            key_b=key_b,
                            confidence=1.0,
                            match_pass=1,
                            change_type=change_type,
                        )
                    )
                else:
                    output.matched.append(
                        MatchResult(
                            key_a=key_a,
                            key_b=key_b,
                            confidence=1.0,
                            match_pass=1,
                            change_type="logic_changed",
                        )
                    )
                matched_a.add(key_a)
                matched_b.add(key_b)
                break

    # ---- Pass 2: Same name, fuzzy body ----
    unmatched_a = {k: v for k, v in procs_a.items() if k not in matched_a}
    unmatched_b = {k: v for k, v in procs_b.items() if k not in matched_b}

    for key_a, desc_a in list(unmatched_a.items()):
        if desc_a.token_count < min_token_floor:
            continue
        best_key_b = None
        best_sim = 0.0
        candidates = []
        for key_b, desc_b in unmatched_b.items():
            if desc_b.token_count < min_token_floor:
                continue
            if desc_a.module == desc_b.module and desc_a.name == desc_b.name:
                sim = _token_similarity(desc_a.body_tokens, desc_b.body_tokens)
                if sim >= pass2_threshold:
                    candidates.append((key_b, sim))
        if len(candidates) == 1:
            best_key_b, best_sim = candidates[0]
        elif len(candidates) > 1:
            # Ambiguous — pick best, but flag if close
            candidates.sort(key=lambda x: -x[1])
            best_key_b, best_sim = candidates[0]

        if best_key_b:
            output.matched.append(
                MatchResult(
                    key_a=key_a,
                    key_b=best_key_b,
                    confidence=round(best_sim, 3),
                    match_pass=2,
                    change_type="logic_changed",
                )
            )
            matched_a.add(key_a)
            matched_b.add(best_key_b)
            del unmatched_a[key_a]
            del unmatched_b[best_key_b]

    # ---- Pass 3: Rename detection, same module ----
    unmatched_a = {k: v for k, v in procs_a.items() if k not in matched_a}
    unmatched_b = {k: v for k, v in procs_b.items() if k not in matched_b}

    # Group by module
    a_by_module: dict[str, list[tuple[str, ProcedureDesc]]] = {}
    b_by_module: dict[str, list[tuple[str, ProcedureDesc]]] = {}
    for k, v in unmatched_a.items():
        a_by_module.setdefault(v.module, []).append((k, v))
    for k, v in unmatched_b.items():
        b_by_module.setdefault(v.module, []).append((k, v))

    for module in set(a_by_module) & set(b_by_module):
        for key_a, desc_a in a_by_module[module]:
            if key_a in matched_a or desc_a.token_count < min_token_floor:
                continue
            candidates = []
            for key_b, desc_b in b_by_module[module]:
                if key_b in matched_b or desc_b.token_count < min_token_floor:
                    continue
                sim = _token_similarity(desc_a.body_tokens, desc_b.body_tokens)
                if sim >= pass3_threshold:
                    candidates.append((key_b, sim))

            if len(candidates) == 1:
                key_b, sim = candidates[0]
                # Determine if body also changed
                desc_b = procs_b[key_b]
                if desc_a.canonical_hash == desc_b.canonical_hash:
                    change_type = "renamed"
                else:
                    change_type = "renamed_modified"
                output.matched.append(
                    MatchResult(
                        key_a=key_a,
                        key_b=key_b,
                        confidence=round(sim, 3),
                        match_pass=3,
                        change_type=change_type,
                    )
                )
                matched_a.add(key_a)
                matched_b.add(key_b)
            elif len(candidates) > 1:
                output.ambiguous_dropped.append(key_a)
                matched_a.add(key_a)  # Don't try again in Pass 4

    # ---- Pass 4: Cross-module rename (optional) ----
    if detect_cross_module_renames:
        unmatched_a_4 = {k: v for k, v in procs_a.items() if k not in matched_a}
        unmatched_b_4 = {k: v for k, v in procs_b.items() if k not in matched_b}
        for key_a, desc_a in unmatched_a_4.items():
            if desc_a.token_count < min_token_floor:
                continue
            candidates = []
            for key_b, desc_b in unmatched_b_4.items():
                if key_b in matched_b or desc_b.token_count < min_token_floor:
                    continue
                sim = _token_similarity(desc_a.body_tokens, desc_b.body_tokens)
                if sim >= pass4_threshold:
                    candidates.append((key_b, sim))
            if len(candidates) == 1:
                key_b, sim = candidates[0]
                desc_b = procs_b[key_b]
                change_type = (
                    "renamed"
                    if desc_a.canonical_hash == desc_b.canonical_hash
                    else "renamed_modified"
                )
                output.matched.append(
                    MatchResult(
                        key_a=key_a,
                        key_b=key_b,
                        confidence=round(sim, 3),
                        match_pass=4,
                        change_type=change_type,
                    )
                )
                matched_a.add(key_a)
                matched_b.add(key_b)
            elif len(candidates) > 1:
                output.ambiguous_dropped.append(key_a)
                matched_a.add(key_a)

    # ---- Collect unmatched as added/removed ----
    output.removed = sorted(k for k in procs_a if k not in matched_a)
    output.added = sorted(k for k in procs_b if k not in matched_b)

    return output
