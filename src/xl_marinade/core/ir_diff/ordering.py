# ABOUTME: Stage 5b — Deterministic 12-tier ordering for diff changes.
# ABOUTME: Within each tier, sort by canonical key then JSON fallback.

from __future__ import annotations

import json

from xl_marinade.core.ir_diff.change_types import ORDERING_TIERS
from xl_marinade.core.ir_diff.model import Change


def order_changes(changes: list[Change]) -> list[Change]:
    """Sort changes into deterministic order.

    Ordering tiers (lower = earlier):
    0: metadata, 1: root, 2: sheets, 3: rows/cols, 4: names, 5: tables,
    6: bindings, 7: cells, 8: label evidence/time, 9: edges,
    10: families, 11: resolution metrics

    Within a tier: sort by sort_key tuple, then by JSON of details as final tie-break.

    Args:
        changes: Unordered list of Change objects.

    Returns:
        New list sorted in deterministic order.
    """

    def sort_key(c: Change) -> tuple:
        tier = ORDERING_TIERS.get(c.type, 99)
        # Normalize sort_key to a comparable tuple of strings
        normalized_sk = tuple(str(x) for x in c.sort_key) if c.sort_key else ()
        # Final tie-break: deterministic JSON of details
        json_fb = json.dumps(c.details, sort_keys=True, separators=(",", ":"), default=str)
        return (tier, c.type, normalized_sk, json_fb)

    return sorted(changes, key=sort_key)
