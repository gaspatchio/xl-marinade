"""Thin Typer CLI for XL Marinade.

This layer only parses arguments, calls the ``xl_marinade`` library, renders
output, and maps the typed error hierarchy to process exit codes. It contains
no business logic — that all lives in ``xl_marinade.core``.

Exit codes: 0 success, 1 generic error, 3 memory-budget exceeded, 4 LLM unavailable,
5 calc engine unavailable, 6 workbook refused for recalculation.
"""

from __future__ import annotations

import importlib.metadata
import importlib.util
import json
from datetime import date
from pathlib import Path

import typer
from rich.console import Console
from rich.markup import escape

import xl_marinade
from xl_marinade.errors import (
    CalcUnavailable,
    LLMUnavailable,
    MarinadeError,
    MemoryBudgetExceeded,
    WorkbookRefused,
)

app = typer.Typer(add_completion=False)
_err = Console(stderr=True)


def _version_callback(value: bool) -> None:
    """Print the installed version and exit, before any argument validation."""
    if not value:
        return
    # Read the installed distribution rather than a module constant: it is the
    # version the user actually has, and it cannot drift from the package
    # metadata the way a hand-maintained __version__ can.
    typer.echo(importlib.metadata.version("xl-marinade"))
    raise typer.Exit


@app.callback()
def _main(
    version: bool = typer.Option(  # noqa: ARG001 - consumed by the eager callback
        False,
        "--version",
        "-V",
        callback=_version_callback,
        is_eager=True,
        help="Show the installed version and exit.",
    ),
) -> None:
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


@app.command()
def values(
    workbook: Path = typer.Argument(
        ..., exists=True, dir_okay=False, readable=True, help="Excel workbook (.xlsx/.xlsm)."
    ),
    variables: Path = typer.Option(
        ..., "--variables", exists=True, dir_okay=False, readable=True, help="vars.json file."
    ),
    input_map: Path | None = typer.Option(
        None,
        "--map",
        exists=True,
        dir_okay=False,
        readable=True,
        help="map.json (policy id field).",
    ),
    ir: Path | None = typer.Option(
        None, "--ir", exists=True, dir_okay=False, readable=True, help="Reuse this IR database."
    ),
    out: Path = typer.Option(..., "--out", "-o", help="Output directory (rows.csv + run.json)."),
) -> None:
    """Experimental: write the workbook's stored values as variable/period rows."""
    from xl_marinade.calc import ValuesRequest
    from xl_marinade.calc import values as _values

    try:
        rows_path = _values(ValuesRequest(workbook, variables, out, input_map, ir))
    except MarinadeError as exc:
        _err.print(f"[red]error:[/] {escape(str(exc))}")
        raise typer.Exit(code=1)
    _err.print(f"[green]wrote[/] {rows_path}")


@app.command()
def calc(
    workbook: Path = typer.Argument(
        ..., exists=True, dir_okay=False, readable=True, help="Excel workbook (.xlsx/.xlsm)."
    ),
    policies: Path = typer.Option(
        ..., "--policies", exists=True, dir_okay=False, readable=True, help="Policies CSV file."
    ),
    input_map: Path = typer.Option(
        ..., "--map", exists=True, dir_okay=False, readable=True, help="map.json (inputs)."
    ),
    outputs: list[str] = typer.Option(
        [], "--outputs", help="Output cell or range reference; repeat for several."
    ),
    variables: Path | None = typer.Option(
        None, "--variables", exists=True, dir_okay=False, readable=True, help="vars.json file."
    ),
    check_against: Path | None = typer.Option(
        None,
        "--check-against",
        exists=True,
        dir_okay=False,
        readable=True,
        help="Input sets with stored workbook results to check the recalculation against.",
    ),
    ir: Path | None = typer.Option(
        None, "--ir", exists=True, dir_okay=False, readable=True, help="Reuse this IR database."
    ),
    as_of: str | None = typer.Option(None, "--as-of", help="Valuation date, YYYY-MM-DD."),
    tolerance: float = typer.Option(1e-9, "--tolerance", help="Numeric comparison tolerance."),
    probe_cap: int = typer.Option(50_000, "--probe-cap", help="Maximum number of probe cells."),
    out: Path = typer.Option(..., "--out", "-o", help="Output directory."),
) -> None:
    """Experimental: recalculate a workbook per policy and report how far to trust each value."""
    from xl_marinade.calc import CalcRequest
    from xl_marinade.calc import calc as _calc

    try:
        as_of_date = date.fromisoformat(as_of) if as_of is not None else None
    except ValueError:
        _err.print(f"[red]error:[/] --as-of must be YYYY-MM-DD, got {escape(as_of or '')}")
        raise typer.Exit(code=1)
    try:
        result = _calc(
            CalcRequest(
                workbook=workbook,
                policies=policies,
                input_map=input_map,
                out_dir=out,
                outputs=list(outputs),
                variables=variables,
                check_against=check_against,
                ir=ir,
                as_of=as_of_date,
                tolerance=tolerance,
                probe_cap=probe_cap,
            )
        )
    except CalcUnavailable as exc:
        _err.print(f"[red]error:[/] {escape(str(exc))}")
        _err.print("[yellow]hint:[/] pip install 'xl-marinade\\[calc]'")
        raise typer.Exit(code=5)
    except WorkbookRefused as exc:
        _err.print(f"[red]error:[/] {escape(str(exc))}")
        for reason in exc.reasons:
            _err.print(f"  - {escape(reason)}")
        raise typer.Exit(code=6)
    except MarinadeError as exc:
        _err.print(f"[red]error:[/] {escape(str(exc))}")
        raise typer.Exit(code=1)
    _err.print(f"[green]wrote[/] {result.results_csv.parent}")
    _warn_probe_cap(result.trust_report, probe_cap)


def _warn_probe_cap(trust_report: Path, cap: int) -> None:
    """One line on stderr when probes did not fit under ``--probe-cap``."""
    try:
        over = int(json.loads(trust_report.read_text(encoding="utf-8")).get("probes_over_cap", 0))
    except (OSError, ValueError, TypeError):
        return
    if over <= 0:
        return
    _err.print(
        f"[yellow]warning:[/] {over} probes did not fit under --probe-cap {cap}, so the "
        f"values behind them are not ok; rerun with --probe-cap {cap + over} to place them all",
        soft_wrap=True,
    )


def main() -> None:
    """Console-script entry point (marinade = xl_marinade.cli.main:main)."""
    app()


if __name__ == "__main__":
    main()
