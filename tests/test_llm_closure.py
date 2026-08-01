"""The shippable LLM subset imports cleanly, with no documentation_agent dependency."""

import builtins
import importlib
import sys

# Deterministic modules now live in the docs tier (free, no LLM dependency).
_DOCS_MODULES = [
    "xl_marinade.docs.two_pass_labeller",
    "xl_marinade.docs.apply_actuarial_classification",
    "xl_marinade.docs.json_spec_generator",
    "xl_marinade.docs.generators.markdown",
    "xl_marinade.docs.confidence_scorer",
]
# enrichment_service stays in llm:
_LLM_MODULES = ["xl_marinade.llm.enrichment_service"]

ENTRY = _DOCS_MODULES + _LLM_MODULES


def test_llm_entry_modules_import():
    for module in ENTRY:
        importlib.import_module(module)


def test_no_documentation_agent_import(monkeypatch):
    """None of the entry modules may transitively import documentation_agent."""
    real_import = builtins.__import__

    def guard(name, *args, **kwargs):
        if name == "documentation_agent" or name.startswith("documentation_agent."):
            raise ModuleNotFoundError(f"blocked in test: {name}")
        return real_import(name, *args, **kwargs)

    for mod in list(sys.modules):
        if mod == "xl_marinade" or mod.startswith("xl_marinade."):
            monkeypatch.delitem(sys.modules, mod, raising=False)
    monkeypatch.setattr(builtins, "__import__", guard)

    for module in ENTRY:
        importlib.import_module(module)
