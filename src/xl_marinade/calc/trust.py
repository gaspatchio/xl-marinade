"""The trust run: does the engine reproduce what Excel stored, by really recalculating?

The engine must have loaded the cache-free engine copy and been evaluated. A cell is
reproduced only when its stored value exists and matches; everything else is a failure
with exactly one reason.
"""

from __future__ import annotations

import math
from dataclasses import dataclass

from xl_marinade.calc.engine import Engine, EngineValue
from xl_marinade.calc.graph import StoredValue, WorkbookGraph
from xl_marinade.calc.models import CellRef, Reason, ReasonKind

_NAME_ERROR = "#NAME?"


@dataclass(slots=True, frozen=True)
class TrustFailure:
    """A formula cell the engine did not reproduce."""

    cell: CellRef
    reason: Reason


@dataclass(slots=True)
class TrustResult:
    """Counts and failures of one trust run."""

    checked: int
    reproduced: int
    failures: list[TrustFailure]

    @property
    def share(self) -> float:
        """Share of checked cells reproduced (1.0 when nothing was checked)."""
        return self.reproduced / self.checked if self.checked else 1.0


def _is_error(value: object) -> bool:
    return isinstance(value, str) and value.startswith("#")


def _is_number(value: object) -> bool:
    return isinstance(value, int | float) and not isinstance(value, bool)


def same(stored: StoredValue, got: EngineValue, tol: float) -> bool:
    """True when the engine value matches the stored one.

    Numbers match within ``tol``, relative or absolute; NaN never matches. Booleans
    match only booleans, text (and ``#CODE`` errors) only identical text. Empty text and
    an empty cell match each other; an empty cell does not match ``0``.
    """
    if stored in (None, "") or got in (None, ""):
        return stored in (None, "") and got in (None, "")
    if isinstance(stored, bool) or isinstance(got, bool):
        return isinstance(stored, bool) and isinstance(got, bool) and stored == got
    if _is_number(stored) and _is_number(got):
        e, g = float(stored), float(got)  # type: ignore[arg-type]
        if math.isnan(e) or math.isnan(g):
            return False
        return abs(e - g) <= tol * max(1.0, abs(e), abs(g))
    if isinstance(stored, str) and isinstance(got, str):
        return stored == got
    return False


def _short(value: object) -> str:
    text = repr(value)
    return text if len(text) <= 40 else text[:37] + "..."


def _reason(cell: CellRef, graph: WorkbookGraph, stored: StoredValue, got: EngineValue) -> Reason:
    detail = f"stored {_short(stored)}, engine {_short(got)}"
    if not graph.has_stored(cell):
        return Reason(ReasonKind.no_stored_value, cell, f"no stored value, engine {_short(got)}")
    if got == _NAME_ERROR and cell in graph.udf_cells:
        return Reason(ReasonKind.macro_function, cell, detail)
    if got == _NAME_ERROR:
        return Reason(ReasonKind.unsupported_function, cell, detail)
    if _is_error(got) and not _is_error(stored):
        return Reason(ReasonKind.engine_error, cell, detail)
    return Reason(ReasonKind.silently_wrong, cell, detail)


def run_trust(
    engine: Engine, graph: WorkbookGraph, *, tolerance: float, skip: set[CellRef]
) -> TrustResult:
    """Compare every formula cell not in ``skip`` with its stored value.

    ``skip`` holds refused-pattern cells: they are neither checked nor counted. A cell
    with no stored value always fails (``no_stored_value``): it cannot be checked.
    """
    checked = 0
    reproduced = 0
    failures: list[TrustFailure] = []
    for cell in graph.formula_cells():
        if cell in skip:
            continue
        checked += 1
        stored = graph.stored(cell)
        got = engine.value(cell)
        if graph.has_stored(cell) and same(stored, got, tolerance):
            reproduced += 1
            continue
        failures.append(TrustFailure(cell, _reason(cell, graph, stored, got)))
    return TrustResult(checked, reproduced, failures)
