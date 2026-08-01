# ABOUTME: Deterministic JSON canonicalization for blob deduplication
# ABOUTME: Ensures byte-identical output for same content (sorted keys, stable float formatting)

"""
Canonical JSON

Provides deterministic JSON serialization for blob deduplication.

Rules (per design §6.5):
- UTF-8 encoding
- Sorted keys (recursively)
- No whitespace
- Stable float formatting (Python's default repr)
- Consistent handling of None/null

Design reference: §6.5 of memory_efficient_extraction_architecture.md
"""

import hashlib
import json
from typing import Any


def canonicalize_json(obj: Any) -> str:
    """
    Serialize object to canonical JSON string.

    Args:
        obj: Python object to serialize

    Returns:
        Canonical JSON string (UTF-8, sorted keys, no whitespace)
    """
    return json.dumps(
        obj,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,  # Reject NaN/Infinity for determinism
    )


def hash_json(obj: Any) -> str:
    """
    Compute SHA256 hash of canonical JSON.

    Args:
        obj: Python object to hash

    Returns:
        SHA256 hex digest
    """
    canonical = canonicalize_json(obj)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def canonicalize_and_hash(obj: Any) -> tuple[str, str]:
    """
    Canonicalize and hash in one pass.

    Args:
        obj: Python object

    Returns:
        Tuple of (canonical_json, sha256_hex)
    """
    canonical = canonicalize_json(obj)
    sha256 = hashlib.sha256(canonical.encode("utf-8")).hexdigest()
    return canonical, sha256
