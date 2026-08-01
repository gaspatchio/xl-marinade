"""The docs tier is free/deterministic: importable without openai, never imports llm."""

import ast
import builtins
import importlib
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parents[1] / "src"


def test_docs_imports_without_openai(monkeypatch):
    # Simulate openai not installed.
    real_import = builtins.__import__

    def guarded(name, *a, **k):
        if name == "openai" or name.startswith("openai."):
            raise ImportError("openai blocked for test")
        return real_import(name, *a, **k)

    for m in list(sys.modules):
        if m == "xl_marinade" or m.startswith("xl_marinade."):
            del sys.modules[m]
    monkeypatch.setattr(builtins, "__import__", guarded)
    importlib.import_module("xl_marinade.docs")  # must not raise


def test_docs_graph_has_no_llm_import():
    docs_dir = SRC / "xl_marinade" / "docs"
    offenders = []
    for py in docs_dir.rglob("*.py"):
        tree = ast.parse(py.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            mod = None
            if isinstance(node, ast.ImportFrom):
                mod = node.module
            elif isinstance(node, ast.Import):
                mod = node.names[0].name if node.names else None
            if mod and mod.startswith("xl_marinade.llm"):
                offenders.append(f"{py.name}: {mod}")
    assert offenders == [], f"docs must not import llm: {offenders}"


def test_docs_graph_has_no_documentation_agent_import():
    docs_dir = SRC / "xl_marinade" / "docs"
    offenders = []
    for py in docs_dir.rglob("*.py"):
        tree = ast.parse(py.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            mod = None
            if isinstance(node, ast.ImportFrom):
                mod = node.module
            elif isinstance(node, ast.Import):
                mod = node.names[0].name if node.names else None
            if mod and (mod == "documentation_agent" or mod.startswith("documentation_agent.")):
                offenders.append(f"{py.name}: {mod}")
    assert offenders == [], f"docs must not import documentation_agent: {offenders}"
