"""``marinade values``: the workbook's stored values as variable/period rows."""

from __future__ import annotations

import json
import tempfile
from dataclasses import dataclass
from pathlib import Path

from xl_marinade.calc.config import load_input_map, load_variables
from xl_marinade.calc.graph import StoredValue, WorkbookGraph, open_graph
from xl_marinade.calc.models import CellRef, Reason, Status
from xl_marinade.calc.output import VALUES_OUTPUTS, clear_outputs, write_rows
from xl_marinade.calc.rows import expand_variables, rows_for, variables_record
from xl_marinade.errors import CalcInputError


@dataclass(slots=True)
class ValuesRequest:
    """Inputs for :func:`values`."""

    workbook: Path
    variables: Path
    out_dir: Path
    input_map: Path | None = None
    ir: Path | None = None


def _stored_policy_id(graph: WorkbookGraph, input_map: Path | None) -> str:
    if input_map is None:
        return "stored"
    imap = load_input_map(input_map)
    target = imap.fields.get(imap.id_field) if imap.id_field else None
    if target is None:
        return "stored"
    cells = graph.resolve_target(target)
    if len(cells) != 1:
        raise CalcInputError(f"id field {imap.id_field!r} must map to a single cell")
    value = graph.stored(cells[0])
    if value is None:
        return "stored"
    return str(int(value)) if isinstance(value, float) and value.is_integer() else str(value)


def values(req: ValuesRequest) -> Path:
    """Write ``rows.csv`` and ``run.json`` for the stored values; returns ``rows.csv``.

    Raises:
        CalcInputError: on bad input files, a 2-D range, or an IR from another workbook.
    """
    clear_outputs(req.out_dir, VALUES_OUTPUTS)
    specs = load_variables(req.variables)
    with tempfile.TemporaryDirectory(prefix="marinade-values-") as tmp:
        graph = open_graph(req.workbook, req.ir, Path(tmp))
    variables = expand_variables(specs, graph)
    policy = _stored_policy_id(graph, req.input_map)

    def value_of(cell: CellRef) -> tuple[StoredValue, Status, list[Reason]]:
        return graph.stored(cell), Status.STORED, []

    rows = rows_for(policy, variables, value_of)
    rows_path = req.out_dir / "rows.csv"
    write_rows(rows_path, rows)

    run: dict[str, object] = {"variables": variables_record(specs, variables)}
    (req.out_dir / "run.json").write_text(
        json.dumps(run, indent=2) + "\n", encoding="utf-8", newline="\n"
    )
    return rows_path
