"""The core extraction entry point imports under the new package name."""


def test_core_entrypoint_imports():
    from xl_marinade.core.new_arch.fast_extraction_pipeline import (
        run_full_workbook_extraction,
    )

    assert callable(run_full_workbook_extraction)
