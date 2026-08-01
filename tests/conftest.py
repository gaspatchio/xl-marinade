"""Shared test fixtures.

The dependency graph has two interchangeable backends: a NetworkX ``DiGraph``
(used when ``networkx`` is importable) and a pure-dict adjacency fallback.
``networkx`` is NOT a runtime dependency, so the shipped default for users is
the dict path. Pin that path by default here so the suite is deterministic
regardless of whether ``networkx`` happens to be installed in the environment;
the tests that exercise the ``DiGraph`` backend opt in explicitly (see
``test_dependency_graph.py``).
"""

import pytest

import xl_marinade.docs.dependency_traversal as dependency_traversal


@pytest.fixture(autouse=True)
def _default_dict_graph_backend(monkeypatch):
    """Force the dict graph backend so the suite matches the runtime default."""
    monkeypatch.setattr(dependency_traversal, "HAS_NETWORKX", False)
