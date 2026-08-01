"""llm package exposes `document` lazily (PEP 562): no eager heavy-import, no submodule shadow."""

import importlib
import sys
import types


def _forget_package() -> None:
    for name in list(sys.modules):
        if name == "xl_marinade" or name.startswith("xl_marinade."):
            del sys.modules[name]


def test_import_llm_does_not_eagerly_pull_docs_chain():
    _forget_package()
    importlib.import_module("xl_marinade.llm")
    # The old eager `from ...document import document` dragged in the docs+sprint7 chain.
    assert "xl_marinade.llm._document" not in sys.modules
    assert "xl_marinade.docs.pipeline" not in sys.modules


def test_document_callable_and_impl_module_resolves():
    from xl_marinade.llm import document as document_callable

    assert callable(document_callable)  # PEP 562 __getattr__ yields the function

    mod = importlib.import_module("xl_marinade.llm._document")
    assert isinstance(mod, types.ModuleType)  # the impl module is not shadowed
    assert hasattr(mod, "document")


def test_document_callable_when_impl_module_imported_first():
    """Loading the impl module first must NOT shadow the package-level callable."""
    _forget_package()
    importlib.import_module("xl_marinade.llm._document")  # impl module touched first
    import xl_marinade.llm  # noqa: F401
    from xl_marinade.llm import document as doc

    assert callable(doc), type(doc)
