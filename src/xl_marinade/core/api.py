"""Public library API for XL Marinade.

The heavy extraction pipeline is imported lazily inside :func:`extract` so that
``import xl_marinade`` stays cheap and side-effect-free.
"""

from __future__ import annotations

from pathlib import Path

from xl_marinade.errors import DiffError, ExtractionError, MarinadeError


def extract(
    workbook: str | Path,
    out: str | Path,
    *,
    max_memory_mb: int = 1800,
    enrich: bool = False,
) -> Path:
    """Extract an Excel workbook's formula graph to a SQLite database.

    Deterministic and network-free by default. Returns the path to the written
    database.

    Args:
        enrich: opt-in LLM VBA enrichment (makes network calls; requires the
            ``xl-marinade[llm]`` extra). Defaults to ``False`` — a bare install
            never contacts a network, even if ``OPENAI_API_KEY`` is set.

    Raises:
        xl_marinade.errors.ExtractionError: if extraction fails.
    """
    from xl_marinade.core.new_arch.fast_extraction_pipeline import (
        run_full_workbook_extraction,
    )

    workbook_path, out_path = Path(workbook), Path(out)

    # Create the output's parent rather than letting SQLite fail on it. The
    # bare "unable to open database file" names neither the cause nor the path,
    # and it arrived *after* the CLI had already printed the output path as
    # though it were fine (issues #17, #31). The telemetry writer alongside
    # this one already created its parent, so the two disagreed.
    try:
        out_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise ExtractionError(
            f"cannot create the output directory {str(out_path.parent)!r}: {exc}"
        ) from exc

    try:
        run_full_workbook_extraction(
            workbook_path, out_path, max_memory_mb=max_memory_mb, enrich=enrich
        )
    except MarinadeError:
        raise
    except Exception as exc:  # noqa: BLE001 - re-raise as a typed error
        raise ExtractionError(str(exc)) from exc
    return out_path


def diff(db_a: str | Path, db_b: str | Path) -> dict:
    """Compare two IR databases and return a replay-complete changelist.

    Deterministic and network-free. Thin typed wrapper over
    :func:`xl_marinade.core.ir_diff.pipeline.diff_ir` so callers (and the CLI)
    see a typed error instead of a raw ``sqlite3.Error`` when an input database
    is unreadable, corrupt, or not an IR database.

    Raises:
        xl_marinade.errors.DiffError: if the diff fails for an unexpected reason
            (e.g. an unreadable or corrupt IR database).
        xl_marinade.errors.MarinadeError: for typed diff failures such as
            ``xl_marinade.core.ir_diff.model.DiffVerificationError``, which are
            re-raised unchanged.
    """
    from xl_marinade.core.ir_diff.pipeline import diff_ir

    try:
        return diff_ir(str(db_a), str(db_b))
    except MarinadeError:
        raise
    except Exception as exc:  # noqa: BLE001 - re-raise as a typed error
        raise DiffError(str(exc)) from exc
