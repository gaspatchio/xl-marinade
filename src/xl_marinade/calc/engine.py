"""The only module that imports ``ironcalc`` (lazily)."""

from __future__ import annotations

import os
import sys
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from xl_marinade.calc.models import CellRef
from xl_marinade.errors import CalcError, CalcInputError, CalcUnavailable

EngineValue = float | str | bool | None
"""A computed cell value: number, text, boolean, ``#CODE`` error text, or None (empty)."""

_LOCALE = "en"
_TIMEZONE = "UTC"
_LANGUAGE = "en"


def _import_ironcalc() -> Any:
    try:
        import ironcalc
    except ImportError as exc:
        raise CalcUnavailable(
            "the calc engine is not installed; run pip install 'xl-marinade[calc]'"
        ) from exc
    return ironcalc


@contextmanager
def quiet_native_output() -> Iterator[None]:
    """Send fds 1 and 2 to ``os.devnull`` for the duration.

    IronCalc's Rust code prints diagnostics (``Unexpected type (empty) ...``) straight
    to the process's stdout/stderr when it loads a cache-free copy; Python-level
    redirection cannot catch them.
    """
    sys.stdout.flush()
    sys.stderr.flush()
    saved = [os.dup(1), os.dup(2)]
    try:
        with open(os.devnull, "wb") as devnull:
            os.dup2(devnull.fileno(), 1)
            os.dup2(devnull.fileno(), 2)
        yield
    finally:
        sys.stdout.flush()
        sys.stderr.flush()
        os.dup2(saved[0], 1)
        os.dup2(saved[1], 2)
        os.close(saved[0])
        os.close(saved[1])


_PANIC = "PanicException"
"""pyo3_runtime's exception for a Rust panic: a BaseException, matched by name since
``pyo3_runtime`` is not importable on its own."""


def _is_panic(exc: BaseException) -> bool:
    return type(exc).__name__ == _PANIC


def _crash(exc: BaseException, version: str, where: str) -> CalcError:
    """A Rust panic from the engine as a typed error."""
    return CalcError(
        f"the calc engine (IronCalc {version}) crashed {where}: {exc}. A formula in the "
        "workbook hits an engine bug, so the run stopped"
    )


def engine_version() -> str | None:
    """The installed IronCalc version, or None when it is not installed."""
    try:
        ironcalc = _import_ironcalc()
    except CalcUnavailable:
        return None
    return str(getattr(ironcalc, "__version__", "unknown"))


class Engine:
    """Thin wrapper over an IronCalc model."""

    def __init__(self, model: Any, version: str, source: Path | None = None) -> None:
        self._model = model
        self.version: str = version
        self._source = source
        self._sheets: dict[str, int] = {}
        self._refresh_sheets()

    @classmethod
    def load(cls, path: Path, workbook: Path | None = None) -> Engine:
        """Load a workbook into the engine (not yet evaluated); ``workbook`` is the
        user's file ``path`` was prepared from, named in errors.

        Raises:
            CalcUnavailable: if ``ironcalc`` is not importable.
            CalcError: if the engine cannot load the workbook, or crashes loading it.
        """
        ironcalc = _import_ironcalc()
        version = str(getattr(ironcalc, "__version__", "unknown"))
        try:
            model = ironcalc.load_from_xlsx(str(path), _LOCALE, _TIMEZONE, _LANGUAGE)
        except Exception as exc:  # the binding raises WorkbookError or a bare Exception
            raise CalcError(f"the calc engine could not load {path}: {exc}") from exc
        except BaseException as exc:
            if not _is_panic(exc):
                raise
            raise _crash(exc, version, f"loading {workbook or path}") from exc
        return cls(model, version, workbook or path)

    def _native(self, what: str, call: Callable[..., Any], *args: Any) -> Any:
        """Call the engine, turning a Rust panic into a :class:`CalcError`; every
        other exception, interrupts included, passes through unchanged."""
        try:
            return call(*args)
        except BaseException as exc:
            if not _is_panic(exc):
                raise
            where = f"{what} in {self._source}" if self._source else what
            raise _crash(exc, self.version, where) from exc

    def _refresh_sheets(self) -> None:
        props = self._native("reading sheets", self._model.get_worksheets_properties)
        self._sheets = {str(p["name"]): i for i, p in enumerate(props)}

    def _sheet(self, cell: CellRef) -> int:
        idx = self._sheets.get(cell.sheet)
        if idx is None:
            raise CalcInputError(f"unknown sheet {cell.sheet!r} for cell {cell.a1()}")
        return idx

    def set_value(self, cell: CellRef, value: float | str | bool) -> None:
        """Set a constant: number, text or boolean (no formula parsing)."""
        idx, what = self._sheet(cell), f"setting {cell.a1()}"
        if isinstance(value, bool):
            self._native(what, self._model.update_cell_with_bool, idx, cell.row, cell.col, value)
            return
        if isinstance(value, str):
            self._native(what, self._model.update_cell_with_text, idx, cell.row, cell.col, value)
            return
        self._native(
            what, self._model.update_cell_with_number, idx, cell.row, cell.col, float(value)
        )

    def clear(self, cell: CellRef) -> None:
        """Delete a cell's contents (it reads back as empty)."""
        self._native(
            f"clearing {cell.a1()}",
            self._model.clear_cell_contents,
            self._sheet(cell),
            cell.row,
            cell.col,
        )

    def set_formula(self, cell: CellRef, formula: str) -> None:
        """Set a formula; a missing leading ``=`` is added."""
        text = formula if formula.startswith("=") else f"={formula}"
        self._native(
            f"setting {cell.a1()} to {text}",
            self._model.update_cell_with_formula,
            self._sheet(cell),
            cell.row,
            cell.col,
            text,
        )

    def add_hidden_sheet(self, name: str) -> None:
        """Add a sheet named ``name`` and hide it.

        Raises:
            CalcInputError: if the workbook has a sheet of that name (in any case).
        """
        taken = next((s for s in self._sheets if s.lower() == name.lower()), None)
        if taken is not None:
            raise CalcInputError(
                f"the workbook has a sheet named {taken!r}, the name calc gives its hidden "
                "probe sheet; rename that sheet in a copy of the workbook"
            )
        self._native("adding the probe sheet", self._model.add_sheet, name)
        self._refresh_sheets()
        self._native(
            "hiding the probe sheet", self._model.set_sheet_state, self._sheets[name], "hidden"
        )

    def evaluate(self) -> None:
        """Recalculate the whole workbook.

        Raises:
            CalcError: if the engine crashes (an IronCalc panic, such as on
                ``XNPV(1,1)``); it does not say which cell.
        """
        self._native("recalculating", self._model.evaluate)

    def is_error(self, cell: CellRef) -> bool:
        """True when the cell holds an error value (not text that starts with ``#``)."""
        kind = self._native(
            f"reading {cell.a1()}",
            self._model.get_cell_type,
            self._sheet(cell),
            cell.row,
            cell.col,
        )
        return str(getattr(kind, "name", kind)).endswith("ErrorValue")

    def value(self, cell: CellRef) -> EngineValue:
        """The computed value of a cell; errors come back as ``#CODE`` text."""
        got = self._native(
            f"reading {cell.a1()}",
            self._model.get_cell_value,
            self._sheet(cell),
            cell.row,
            cell.col,
        )
        if got is None or isinstance(got, bool | str):
            return got
        if isinstance(got, int | float):
            return float(got)
        return str(got)
