"""The core extraction path must run with documentation_agent unavailable.

This is the load-bearing decoupling guarantee for the free Tier-0 core: the OSS
install ships no `documentation_agent` package, so a full extraction must succeed
even when importing it is impossible.
"""

import builtins
import sqlite3
import sys

import pytest

from test_workbook_generator.cli import create_comprehensive_test_workbook


@pytest.fixture
def blocked_documentation_agent(monkeypatch):
    real_import = builtins.__import__

    def guard(name, *args, **kwargs):
        if name == "documentation_agent" or name.startswith("documentation_agent."):
            raise ModuleNotFoundError(f"blocked in test: {name}")
        return real_import(name, *args, **kwargs)

    for mod in list(sys.modules):
        if mod == "documentation_agent" or mod.startswith("documentation_agent."):
            monkeypatch.delitem(sys.modules, mod, raising=False)
    monkeypatch.setattr(builtins, "__import__", guard)


def test_full_extraction_without_documentation_agent(tmp_path, blocked_documentation_agent):
    xlsx = tmp_path / "sample.xlsx"
    create_comprehensive_test_workbook(xlsx)

    from xl_marinade.core.new_arch.fast_extraction_pipeline import (
        run_full_workbook_extraction,
    )

    out_db = tmp_path / "ir.db"
    run_full_workbook_extraction(xlsx, out_db)

    assert out_db.exists()
    conn = sqlite3.connect(out_db)
    (n_bindings,) = conn.execute("SELECT COUNT(*) FROM agent_bindings").fetchone()
    conn.close()
    assert n_bindings > 0
