"""Turn the variable map into variable/period rows."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from xl_marinade.calc.graph import StoredValue, WorkbookGraph
from xl_marinade.calc.models import CellRef, Reason, Status, VariableSpec
from xl_marinade.errors import CalcInputError


@dataclass(slots=True)
class VariableCells:
    """A variable's cells in period order."""

    name: str
    cells: list[CellRef]
    period_start: int | None


@dataclass(slots=True)
class Row:
    """One output row: a variable at a period (None for a single cell) for one policy."""

    policy: str
    variable: str
    period: int | None
    value: StoredValue
    status: Status
    reasons: list[Reason]


def expand_variables(specs: list[VariableSpec], graph: WorkbookGraph) -> list[VariableCells]:
    """Resolve each spec to its cells; a range must be one row or one column of one sheet.

    Raises:
        CalcInputError: for a 2-D range, a range over more than one sheet (a name with
            areas on several sheets), or a ``cell`` entry that spans several cells.
    """
    out: list[VariableCells] = []
    for spec in specs:
        cells = graph.resolve_target(spec.target)
        if spec.period_start is None:
            if len(cells) != 1:
                raise CalcInputError(f"variable {spec.name!r}: 'cell' must be a single cell")
            out.append(VariableCells(spec.name, cells, None))
            continue
        sheets = sorted({c.sheet for c in cells})
        if len(sheets) > 1:
            # Periods follow rows, so cells from two sheets would interleave.
            raise CalcInputError(
                f"variable {spec.name!r}: range {spec.target!r} spans more than one sheet "
                f"({', '.join(sheets)}); give one row or one column of one sheet"
            )
        if len({c.col for c in cells}) > 1 and len({c.row for c in cells}) > 1:
            raise CalcInputError(
                f"variable {spec.name!r}: 2-D range {spec.target!r}; use one row or one column"
            )
        out.append(
            VariableCells(spec.name, sorted(cells, key=lambda c: (c.row, c.col)), spec.period_start)
        )
    return out


def period_0_cell(var: VariableCells) -> CellRef | None:
    """The cell ``period_start`` steps before the first cell, or None if off the sheet.

    A single-cell range is treated as a column (the step goes up the rows).
    """
    if var.period_start is None or not var.cells:
        return None
    first = var.cells[0]
    along_row = len(var.cells) > 1 and var.cells[0].row == var.cells[1].row
    row, col = (
        (first.row, first.col - var.period_start)
        if along_row
        else (
            first.row - var.period_start,
            first.col,
        )
    )
    if row < 1 or col < 1:
        return None
    return CellRef(first.sheet, row, col)


def rows_for(
    policy: str,
    variables: list[VariableCells],
    value_of: Callable[[CellRef], tuple[StoredValue, Status, list[Reason]]],
) -> list[Row]:
    """Build rows for one policy, variables in order and periods ascending."""
    rows: list[Row] = []
    for var in variables:
        for i, cell in enumerate(var.cells):
            value, status, reasons = value_of(cell)
            period = None if var.period_start is None else var.period_start + i
            rows.append(Row(policy, var.name, period, value, status, list(reasons)))
    return rows


def variables_record(
    specs: list[VariableSpec], variables: list[VariableCells]
) -> dict[str, dict[str, object]]:
    """``run.json``'s ``variables``: per variable its range, ``period_start`` and the
    cell period 0 falls on (None for a single cell or when it is off the sheet)."""
    by_name = {s.name: s for s in specs}
    entries: dict[str, dict[str, object]] = {}
    for var in variables:
        p0 = period_0_cell(var)
        entries[var.name] = {
            "range": by_name[var.name].target,
            "period_start": var.period_start,
            "period_0_cell": p0.a1() if p0 else None,
        }
    return entries
