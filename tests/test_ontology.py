"""Ontology loader tolerates absence and resolves shipped package data."""

import json

from xl_marinade.docs import ontology_loader as ol


def test_shipped_basic_ontology_loads(monkeypatch):
    monkeypatch.delenv("MARINADE_ONTOLOGY_PATH", raising=False)
    ol._CACHE.clear()
    ont = ol.load_ontology(domain="actuarial")
    assert ont["ontology_version"]
    assert isinstance(ont["concepts"], list)


def test_absent_env_path_degrades_to_generic(monkeypatch, tmp_path):
    ol._CACHE.clear()
    missing = tmp_path / "nope.json"
    monkeypatch.setenv("MARINADE_ONTOLOGY_PATH", str(missing))
    ont = ol.load_ontology(domain="actuarial")  # must NOT raise
    assert ont["concepts"] == []


def test_explicit_path_still_loads(tmp_path):
    ol._CACHE.clear()
    p = tmp_path / "ont.json"
    p.write_text(
        json.dumps(
            {
                "ontology_version": "test",
                "concepts": [{"id": "c1", "label": "Premium", "description": "d", "synonyms": []}],
            }
        )
    )
    ont = ol.load_ontology(path=p, domain="actuarial")
    assert ont["concepts"][0]["label"] == "Premium"


def test_degraded_result_is_not_globally_shared(monkeypatch, tmp_path):
    """Degrading for one missing path must not corrupt a degrade for a different path."""
    ol._CACHE.clear()
    monkeypatch.setenv("MARINADE_ONTOLOGY_PATH", str(tmp_path / "a.json"))
    ont_a = ol.load_ontology(domain="actuarial")
    ont_a["concepts"].append("CORRUPTION")  # mutate the degraded result

    monkeypatch.setenv("MARINADE_ONTOLOGY_PATH", str(tmp_path / "b.json"))
    ont_b = ol.load_ontology(domain="actuarial")
    assert ont_b["concepts"] == []  # independent, clean object


def test_degraded_missing_path_is_cached(monkeypatch, tmp_path):
    """Repeat degrade for the same missing path is served from cache (no re-stat)."""
    ol._CACHE.clear()
    missing = tmp_path / "missing.json"
    monkeypatch.setenv("MARINADE_ONTOLOGY_PATH", str(missing))
    first = ol.load_ontology(domain="actuarial")
    second = ol.load_ontology(domain="actuarial")
    assert first == second  # same content, served from cache (no re-stat)
    assert first is not second  # but each call returns an independent copy
    assert str(missing.resolve()) in ol._CACHE


def test_loaded_ontology_mutation_does_not_corrupt_cache(tmp_path):
    """Mutating a returned ontology (even nested) must not corrupt the cached copy."""
    ol._CACHE.clear()
    p = tmp_path / "ont.json"
    p.write_text(
        json.dumps(
            {
                "ontology_version": "test",
                "concepts": [{"id": "c1", "label": "Premium", "description": "d", "synonyms": []}],
            }
        )
    )
    first = ol.load_ontology(path=p, domain="actuarial")
    first["concepts"].append({"id": "HACK", "label": "x", "description": "", "synonyms": []})
    first["concepts"][0]["label"] = "CORRUPTED"

    second = ol.load_ontology(path=p, domain="actuarial")
    assert len(second["concepts"]) == 1
    assert second["concepts"][0]["label"] == "Premium"
