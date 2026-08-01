"""
Ontology loader and configuration.

Centralizes how we locate and load the actuarial ontology so:
- runtime behavior is explicit and configurable
- semantic_index, agent orchestration, and other modules stay consistent

Domain parameter (P2-T0a): ``load_ontology()`` accepts a ``domain`` keyword
argument that is threaded through the call surface but not yet dispatched
— the loader ignores the value and continues to resolve the actuarial
ontology via ``get_ontology_path()`` / ``MARINADE_ONTOLOGY_PATH``.
The parameter exists as a preparatory hook for plugin-shaped ontologies
(see invariant I6). A later task will wire ``domain`` to a per-domain
ontology selector; until then every caller passes ``domain="actuarial"``
explicitly so the static-analysis regression can prove no caller relies
on an implicit default.
"""

from __future__ import annotations

import copy
import json
import os
from importlib.resources import files
from pathlib import Path
from typing import Any

from loguru import logger

_CACHE: dict[str, dict[str, Any]] = {}


def _shipped_ontology_text() -> str | None:
    """Read the ontology shipped as package data (zip-safe). None if unavailable."""
    try:
        res = files("xl_marinade.docs").joinpath("ontology", "basic.json")
        if res.is_file():
            return res.read_text(encoding="utf-8")
    except (ModuleNotFoundError, FileNotFoundError, OSError):
        return None
    return None


def get_ontology_path() -> Path | None:
    """Env-var override path, or None to signal 'use shipped package data'.

    Config: MARINADE_ONTOLOGY_PATH — absolute or CWD-relative JSON path.
    """
    raw = (os.getenv("MARINADE_ONTOLOGY_PATH") or "").strip()
    if raw:
        p = Path(raw)
        return p if p.is_absolute() else (Path.cwd() / p)
    return None


def load_ontology(
    *,
    path: Path | None = None,
    domain: str = "actuarial",
) -> dict[str, Any]:
    """
    Load ontology JSON from disk with a small in-process cache.

    Expected schema (minimum):
      - ontology_version: str
      - concepts: list[dict] where each concept has id/label/description/synonyms

    Precedence: explicit ``path`` -> ``MARINADE_ONTOLOGY_PATH`` env
    var -> shipped ``ontology/basic.json`` package data -> generic
    ``{"ontology_version": "0", "concepts": []}``. Never raises
    ``FileNotFoundError`` for a missing resolvable ontology — a missing or
    unreadable file degrades to the generic ontology with a warning log.

    Returns an independent deep copy on every call: the parsed ontology is
    cached (so repeat calls skip re-reading/re-parsing), but callers receive a
    fresh copy each time and may mutate it freely without corrupting the cache.

    Args:
        path: Optional explicit ontology JSON path. When ``None``, the
            loader resolves the path via ``get_ontology_path()`` (which
            honors ``MARINADE_ONTOLOGY_PATH``), falling back to
            the package-shipped ``basic.json`` when neither is set.
        domain: Preparatory hook for plugin-shaped ontologies (P2-T0a).
            Currently ignored — the loader always resolves the actuarial
            ontology. Callers must pass ``domain="actuarial"`` (or the
            detected domain for future multi-domain call sites) so the
            I6 static-analysis regression can prove no caller relies on
            an implicit default. A later task will wire this to a
            per-domain ontology dispatch.
    """
    # ``domain`` is intentionally unused at this stage — see module
    # docstring. The cache remains keyed on the resolved path.
    del domain

    resolved = path if path is not None else get_ontology_path()
    if resolved is not None:
        p = Path(resolved).resolve()
        cache_key = str(p)
        if cache_key in _CACHE:
            return copy.deepcopy(_CACHE[cache_key])
        if not p.exists():
            logger.warning("Ontology not found at {} — using generic labels", p)
            empty = {"ontology_version": "0", "concepts": []}
            _CACHE[cache_key] = empty
            return copy.deepcopy(empty)
        text = p.read_text(encoding="utf-8")
    else:
        cache_key = "<shipped:basic>"
        if cache_key in _CACHE:
            return copy.deepcopy(_CACHE[cache_key])
        text = _shipped_ontology_text()
        if text is None:
            logger.warning("No ontology available — using generic labels")
            empty = {"ontology_version": "0", "concepts": []}
            _CACHE[cache_key] = empty
            return copy.deepcopy(empty)

    obj = json.loads(text)
    if not isinstance(obj, dict):
        raise ValueError(f"Ontology must be a JSON object: {cache_key}")

    if "ontology_version" not in obj or "concepts" not in obj:
        raise ValueError(
            f"Ontology missing required keys {{'ontology_version','concepts'}}: {cache_key}"
        )

    concepts = obj.get("concepts") or []
    if not isinstance(concepts, list):
        raise ValueError(f"Ontology concepts must be a list: {cache_key}")

    for c in concepts:
        if not isinstance(c, dict):
            raise ValueError(f"Ontology concept must be an object: {cache_key}")
        for k in ("id", "label", "description", "synonyms"):
            if k not in c:
                raise ValueError(f"Ontology concept missing '{k}' (id={c.get('id')}): {cache_key}")

    _CACHE[cache_key] = obj
    return copy.deepcopy(obj)


def load_concept_labels(
    *,
    path: Path | None = None,
    domain: str = "actuarial",
) -> dict[str, str]:
    """Convenience helper: map concept_id -> label.

    ``domain`` is forwarded to :func:`load_ontology`; it is a preparatory
    hook for plugin-shaped ontologies (P2-T0a) and currently ignored.
    """
    ont = load_ontology(path=path, domain=domain)
    out: dict[str, str] = {}
    for c in ont.get("concepts") or []:
        cid = c.get("id")
        label = c.get("label")
        if cid and label:
            out[str(cid)] = str(label)
    return out
