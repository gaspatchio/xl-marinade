# ABOUTME: Derives a binding's hierarchical header_context (primary -> group -> axis ->
# ABOUTME: title -> sheet) from the label_candidates the IR already harvests, so the IR
# ABOUTME: exposes the header HIERARCHY ("Net Amount at Risk > Run Off > Mortality")
# ABOUTME: instead of collapsing it to a single flat label. See
# ABOUTME: research/header_selection_2026-06-27/HEADER_CONTEXT_DESIGN.md.
"""Header-context derivation (production of the scorecard prototype).

A ``header_context`` is an ORDERED list of layers, most-specific -> most-general:

  primary  - the immediate adjacent header (== bindings.label, the labeller's pick)
  group    - a section/group banner above/left of the primary, OR a merged cell spanning
             several bindings (a shared group header)
  axis     - a numeric/index axis label (a binomial step '2', a year) — NOT a name
  title    - a table/block title or defined-name slug
  sheet    - the sheet-name fallback

The layers are derived from the SAME ``label_candidates`` the labeller scores, so no new
scanning is needed. The ``primary`` layer is pinned to the labeller's selected label (not
re-derived) so header_context and bindings.label never disagree.
"""

from __future__ import annotations

import re

# Excel error values are never headers.
_ERROR_VALUES = {
    "#DIV/0!",
    "#REF!",
    "#N/A",
    "#VALUE!",
    "#NAME?",
    "#NULL!",
    "#NUM!",
    "#SPILL!",
    "#CALC!",
    "#GETTING_DATA",
}
# A footnote / cross-reference annotation like '(2)' or '(a)' — not a header.
_ANNOT_RE = re.compile(r"^\(\s*[0-9a-z]{1,3}\s*\)$", re.IGNORECASE)


def is_header_noise(text: str) -> bool:
    """True for tokens that must never become a header layer: blanks, Excel error
    values, footnote/annotation tags like '(2)', and pure punctuation."""
    s = (text or "").strip()
    if not s:
        return True
    if s.upper() in _ERROR_VALUES:
        return True
    if _ANNOT_RE.match(s):
        return True
    if not any(ch.isalnum() for ch in s):
        return True
    return False


def _is_numeric(s) -> bool:
    try:
        float(str(s).strip().replace(",", "").replace("%", ""))
        return True
    except (ValueError, AttributeError):
        return False


def _range_boundaries(a1: str):
    """(min_col, min_row, max_col, max_row) for a bare A1 cell or range; None on failure."""
    m = re.match(r"^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", a1.upper())
    if not m:
        return None

    def col(letters):
        n = 0
        for ch in letters:
            n = n * 26 + (ord(ch) - ord("A") + 1)
        return n

    c1, r1 = col(m.group(1)), int(m.group(2))
    if m.group(3):
        c2, r2 = col(m.group(3)), int(m.group(4))
    else:
        c2, r2 = c1, r1
    return (min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))


def _merge_span_for(addr, merged_ranges):
    """The merged (c1,r1,c2,r2) span containing addr, if any and wider/taller than 1x1."""
    if not addr or not merged_ranges:
        return None
    b = _range_boundaries(addr.split("!")[-1])
    if not b:
        return None
    c1, r1 = b[0], b[1]
    for m in merged_ranges:
        if m[0] <= c1 <= m[2] and m[1] <= r1 <= m[3] and (m[2] > m[0] or m[3] > m[1]):
            return m
    return None


def build_header_context(binding_addr, candidates, primary_label, primary_type, merged_ranges=None):
    """Derive the ordered header hierarchy for one binding.

    Args:
        binding_addr: the binding's A1 address ("Sheet!A1" or "Sheet!A1:B10").
        candidates: the binding's parsed ``label_candidates`` (type/literals/cells/address).
        primary_label: the label the labeller SELECTED (== bindings.label). Pinned as the
            ``primary`` layer so header_context never disagrees with bindings.label.
        primary_type: the selected candidate's type (e.g. 'scan_above', 'scan_below').
        merged_ranges: optional list of (c1,r1,c2,r2) merged spans on the sheet — used to
            tag a header that SPANS multiple columns as a shared ``group``.

    Returns:
        Ordered list of layers ``[{role, text, cell, distance, span}, ...]``, most-specific
        first. Always begins with the ``primary`` layer when a primary_label is given.
    """
    a1 = binding_addr.split("!")[-1]
    bnd = _range_boundaries(a1) or (1, 1, 1, 1)
    bc1, bc2 = bnd[0], bnd[2]
    layers: list[dict] = []
    seen: set = set()

    def add(role, text, cell, distance, span=None):
        t = (text or "").strip()
        key = (role, t.lower())
        if is_header_noise(t) or key in seen:
            return
        seen.add(key)
        layers.append({"role": role, "text": t, "cell": cell, "distance": distance, "span": span})

    # 0. PRIMARY — pinned to the labeller's pick. Locate its cell in the winning candidate.
    prim = (primary_label or "").strip()
    prim_cell = None
    if prim:
        for cand in candidates:
            if cand.get("type") != primary_type:
                continue
            for c in cand.get("cells", []):
                v = c.get("value")
                if v is not None and str(v).strip() == prim:
                    prim_cell = c.get("address")
                    break
            if prim_cell:
                break
        add("primary", prim, prim_cell, 1)

    # 1. scan_above / scan_left: the ordered non-empty text literals (closest-first) ARE the
    #    header stack. The closest is the primary (already added); a farther distinct TEXT
    #    banner is a group. Conservative v1: numerics grabbed by a scan that walked into a
    #    data row are NOT axis labels — skip them (real axis detection needs structural
    #    matrix work, tracked separately). A lone single character (a stray 'x' several rows
    #    up) is not a section banner either — require group text of length >= 2.
    for cand in candidates:
        ct = cand.get("type")
        if ct not in ("scan_above", "scan_left"):
            continue
        cells = cand.get("cells", [])
        lits = cand.get("literals", [])
        for i, lit in enumerate(lits):
            s = (lit or "").strip()
            if not s or _is_numeric(s):
                continue
            cell = cells[i].get("address") if i < len(cells) else None
            span = _merge_span_for(cell, merged_ranges)
            is_group = span and (span[2] - span[0]) > (bc2 - bc1)
            # closest text is the primary (pinned above); anything else is a group banner.
            if is_group or s != prim:
                if len(s) >= 2:
                    add("group", s, cell, i + 1, span)
            else:
                add("primary", s, cell, i + 1, span)

    # 2. merged headers explicitly emitted by the IR (_scan_merged_cells)
    for cand in candidates:
        if (cand.get("type") or "").startswith("merged"):
            for c in cand.get("cells", []):
                span = _merge_span_for(c.get("address"), merged_ranges)
                wide = span and (span[2] - span[0]) > (bc2 - bc1)
                for lit in cand.get("literals", []):
                    if (lit or "").strip() and not _is_numeric(lit):
                        add("group" if wide else "primary", lit, c.get("address"), 1, span)

    # 3. explicit axis segments (row_segment / col_segment) -> axis. These are the IR's
    #    structural axis candidates (a real index row/col), NOT numerics a name-scan walked
    #    into — so they are safe to surface as axis layers.
    for cand in candidates:
        if (cand.get("type") or "") in ("row_segment", "col_segment"):
            cells = cand.get("cells", [])
            for i, lit in enumerate(cand.get("literals", [])):
                if (lit or "").strip():
                    cell = cells[i].get("address") if i < len(cells) else None
                    add("axis", lit, cell, i + 1)

    # 4. defined names / tables -> title
    for cand in candidates:
        if (cand.get("type") or "").startswith(("named", "table")):
            for lit in cand.get("literals", []):
                if (lit or "").strip():
                    add("title", lit, cand.get("address"), None)

    # 5. sheet name -> last resort
    for cand in candidates:
        if cand.get("type") == "sheet_name":
            for lit in cand.get("literals", []):
                add("sheet", lit, cand.get("address"), None)

    role_order = {"primary": 0, "group": 1, "axis": 2, "title": 3, "sheet": 4}
    layers.sort(key=lambda x: (role_order.get(x["role"], 9), x["distance"] or 99))
    return layers
