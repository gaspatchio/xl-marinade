"""Thin Typer CLI for XL Marinade.

This layer only parses arguments, calls the ``xl_marinade`` library, renders
output, and maps the typed error hierarchy to process exit codes. It contains
no business logic — that all lives in ``xl_marinade.core``.

Exit codes: 0 success, 1 generic error, 3 memory-budget exceeded, 4 LLM unavailable.
"""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path

import typer
from rich.console import Console
from rich.markup import escape

import xl_marinade
from xl_marinade.errors import LLMUnavailable, MarinadeError, MemoryBudgetExceeded

app = typer.Typer(add_completion=False)
_err = Console(stderr=True)


@app.callback()
def _main() -> None:
    """XL Marinade - deterministic Excel formula-graph extractor."""
    # ASCII only: this docstring is rendered into --help on stdout, which is
    # strict-encoded under a non-UTF-8 locale (cp932/cp949 Windows pipes, legacy
    # latin-1 locales) -- a non-ASCII character there crashes `marinade --help`.


@app.command()
def extract(
    workbook: Path = typer.Argument(
        ..., exists=True, dir_okay=False, readable=True, help="Excel workbook (.xlsx/.xlsm)."
    ),
    out: Path = typer.Option(Path("ir.db"), "--out", "-o", help="Output SQLite database path."),
    enrich: bool = typer.Option(
        False,
        "--enrich",
        help="Opt-in LLM VBA enrichment (makes network calls; requires xl-marinade\\[llm]).",
    ),
    max_memory_mb: int = typer.Option(
        1800,
        "--max-memory-mb",
        help="Extraction memory budget in MB; extraction aborts above this to prevent OOM.",
    ),
) -> None:
    """Extract a workbook's formula graph to a SQLite database."""
    try:
        result = xl_marinade.extract(workbook, out, enrich=enrich, max_memory_mb=max_memory_mb)
    except MemoryBudgetExceeded as exc:
        _err.print(f"[red]error:[/] {escape(str(exc))}")
        _err.print("[yellow]hint:[/] re-run with a higher --max-memory-mb")
        raise typer.Exit(code=3)
    except MarinadeError as exc:
        _err.print(f"[red]error:[/] {escape(str(exc))}")
        raise typer.Exit(code=1)
    _err.print(f"[green]wrote[/] {result}")


@app.command()
def document(
    ir_db: Path = typer.Argument(
        ..., exists=True, dir_okay=False, readable=True, help="IR SQLite database from `extract`."
    ),
    out: Path = typer.Option(
        Path("docs_out"),
        "--out",
        "-o",
        help="Output directory (documentation.md + model_spec.json).",
    ),
    enrich: bool = typer.Option(
        False,
        "--enrich",
        help="Opt-in LLM enrichment (requires xl-marinade\\[llm] + a key; degrades to deterministic otherwise).",
    ),
) -> None:
    """Generate documentation for an extracted IR database."""
    try:
        # Detect the [llm] extra by its marker dependency rather than catching a broad
        # ImportError — the latter would also swallow a genuine import failure deep in the
        # llm chain and silently degrade, masking a real breakage as "extra not installed".
        use_llm = enrich and importlib.util.find_spec("openai") is not None
        if enrich and not use_llm:
            _err.print(
                "[yellow]note:[/] xl-marinade\\[llm] not installed — deterministic documentation only"
            )
        if use_llm:
            from xl_marinade.llm import document as _document
        else:
            from xl_marinade.docs import document as _document
        md = _document(ir_db, out)
    except MemoryBudgetExceeded as exc:
        _err.print(f"[red]error:[/] {escape(str(exc))}")
        raise typer.Exit(code=3)
    except LLMUnavailable as exc:
        _err.print(f"[red]error:[/] {escape(str(exc))}")
        raise typer.Exit(code=4)
    except MarinadeError as exc:
        _err.print(f"[red]error:[/] {escape(str(exc))}")
        raise typer.Exit(code=1)
    _err.print(f"[green]wrote[/] {md}")


@app.command()
def diff(
    db_a: Path = typer.Argument(
        ..., exists=True, dir_okay=False, readable=True, help="Version A (earlier) IR database."
    ),
    db_b: Path = typer.Argument(
        ..., exists=True, dir_okay=False, readable=True, help="Version B (later) IR database."
    ),
    out: Path | None = typer.Option(
        None, "--out", "-o", help="Write the JSON changelist here (default: stdout)."
    ),
) -> None:
    """Diff two IR databases; emit the changelist as JSON."""
    from xl_marinade.core.api import diff as _diff

    try:
        result = _diff(str(db_a), str(db_b))
    except MarinadeError as exc:
        _err.print(f"[red]error:[/] {escape(str(exc))}")
        raise typer.Exit(code=1)

    payload = json.dumps(result, indent=2, default=str)
    if out is not None:
        out.write_text(payload, encoding="utf-8", newline="\n")
        _err.print(f"[green]wrote[/] {out}")
    else:
        typer.echo(payload)


def main() -> None:
    """Console-script entry point (marinade = xl_marinade.cli.main:main)."""
    app()


if __name__ == "__main__":
    main()
