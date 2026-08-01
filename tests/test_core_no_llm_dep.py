"""The Tier-0 core imports and extracts with neither openai nor xl_marinade.llm present.

Proves the free core is a clean unit: the LLM add-on ([llm] extra) is not a
dependency of the default extraction path.
"""

import builtins
import sqlite3
import sys

import pytest

from test_workbook_generator.cli import create_comprehensive_test_workbook


@pytest.fixture
def blocked_llm(monkeypatch):
    real_import = builtins.__import__

    def guard(name, *args, **kwargs):
        if (
            name == "openai"
            or name.startswith("openai.")
            or name == "xl_marinade.llm"
            or name.startswith("xl_marinade.llm.")
        ):
            raise ModuleNotFoundError(f"blocked in test: {name}")
        return real_import(name, *args, **kwargs)

    for mod in list(sys.modules):
        if mod == "openai" or mod.startswith("openai.") or mod.startswith("xl_marinade.llm"):
            monkeypatch.delitem(sys.modules, mod, raising=False)
    monkeypatch.setattr(builtins, "__import__", guard)


def test_core_extracts_without_openai_or_llm(tmp_path, blocked_llm):
    xlsx = tmp_path / "s.xlsx"
    create_comprehensive_test_workbook(xlsx)

    import xl_marinade

    out = xl_marinade.extract(xlsx, tmp_path / "ir.db")  # enrich=False default

    assert out.exists()
    conn = sqlite3.connect(out)
    (n_bindings,) = conn.execute("SELECT COUNT(*) FROM agent_bindings").fetchone()
    conn.close()
    assert n_bindings > 0
