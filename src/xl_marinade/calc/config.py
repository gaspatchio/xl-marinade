"""Loaders for the ``marinade calc`` input files.

Targets stay raw strings; a later step resolves them against the workbook graph.
Only file shape is validated here.

``check.json`` shape::

    {
      "results": {"sheet": "Results", "id_column": "A", "first_row": 2,
                  "outputs": {"Summary!B13": "C"}},
      "inputs": {"table": {"sheet": "Inputs", "id_column": "A", "first_row": 2,
                           "fields": {"age": "B"}}}
    }

or ``"inputs": {"selector": "Control!C4"}`` (exactly one of ``table``/``selector``).
``outputs`` maps an output cell target to the results-sheet column holding its
stored value per policy.
"""

from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any

from xl_marinade.calc.models import (
    CheckAgainst,
    InputMap,
    InputTable,
    Policy,
    VariableSpec,
)
from xl_marinade.errors import CalcInputError


def _read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except OSError as exc:
        raise CalcInputError(f"cannot read {path}: {exc}") from exc
    except UnicodeDecodeError as exc:
        raise CalcInputError(f"{path} is not UTF-8 text: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise CalcInputError(f"{path} is not valid JSON: {exc}") from exc


def _str_map(value: Any, what: str) -> dict[str, str]:
    if not isinstance(value, dict) or not value:
        raise CalcInputError(f"{what} must be a non-empty object")
    for k, v in value.items():
        if not isinstance(v, str) or not v.strip():
            raise CalcInputError(f"{what}[{k!r}] must be a non-empty string")
    return dict(value)


def load_input_map(path: Path) -> InputMap:
    """Load ``map.json``: ``{"id_field": ..., "fields": {name: target}}``."""
    data = _read_json(path)
    if not isinstance(data, dict):
        raise CalcInputError(f"{path} must contain a JSON object")
    fields = _str_map(data.get("fields"), "fields")
    id_field = data.get("id_field")
    if id_field is not None:
        if not isinstance(id_field, str) or not id_field:
            raise CalcInputError("id_field must be a non-empty string")
    return InputMap(fields=fields, id_field=id_field)


def load_policies(path: Path, input_map: InputMap) -> list[Policy]:
    """Load ``policies.csv``; ids stay text, and every mapped field must be filled."""
    id_col = input_map.id_field or "id"
    try:
        with path.open(encoding="utf-8-sig", newline="") as fh:
            reader = csv.DictReader(fh)
            header = reader.fieldnames or []
            dupes = sorted({h for h in header if header.count(h) > 1})
            if dupes:
                raise CalcInputError(f"{path} has duplicate column(s): {dupes}")
            if id_col not in header:
                raise CalcInputError(f"{path} has no id column {id_col!r}")
            missing = [f for f in input_map.fields if f not in header]
            if missing:
                raise CalcInputError(f"{path} is missing mapped field(s): {missing}")
            policies: list[Policy] = []
            seen: set[str] = set()
            for row in reader:
                pid = (row[id_col] or "").strip()
                if not pid:
                    raise CalcInputError(f"{path} has a row with a blank {id_col!r}")
                if pid in seen:
                    raise CalcInputError(f"duplicate policy id {pid!r} in {path}")
                seen.add(pid)
                values: dict[str, str] = {}
                for name in input_map.fields:
                    val = (row[name] or "").strip()
                    if not val:
                        raise CalcInputError(f"policy {pid!r}: blank value for field {name!r}")
                    values[name] = val
                policies.append(Policy(id=pid, values=values))
    except OSError as exc:
        raise CalcInputError(f"cannot read {path}: {exc}") from exc
    except UnicodeDecodeError as exc:
        raise CalcInputError(f"{path} is not UTF-8 text (save it as CSV UTF-8): {exc}") from exc
    except csv.Error as exc:
        raise CalcInputError(f"{path} is not a readable CSV file: {exc}") from exc
    return policies


def load_variables(path: Path) -> list[VariableSpec]:
    """Load ``vars.json`` in file order."""
    data = _read_json(path)
    if not isinstance(data, dict) or not data:
        raise CalcInputError(f"{path} must contain a non-empty JSON object")
    specs: list[VariableSpec] = []
    for name, entry in data.items():
        if not isinstance(entry, dict):
            raise CalcInputError(f"variable {name!r} must be an object")
        has_range, has_cell = "range" in entry, "cell" in entry
        if has_range == has_cell:
            raise CalcInputError(f"variable {name!r} needs exactly one of 'range' or 'cell'")
        target = entry["range"] if has_range else entry["cell"]
        if not isinstance(target, str) or not target.strip():
            raise CalcInputError(f"variable {name!r} target must be a non-empty string")
        if has_cell:
            if "period_start" in entry:
                raise CalcInputError(f"variable {name!r}: a cell entry must not have period_start")
            specs.append(VariableSpec(name=name, target=target, period_start=None))
            continue
        ps = entry.get("period_start")
        if not isinstance(ps, int) or isinstance(ps, bool) or ps < 0:
            raise CalcInputError(f"variable {name!r}: range entry needs period_start (int >= 0)")
        specs.append(VariableSpec(name=name, target=target, period_start=ps))
    return specs


def _sheet_block(data: Any, what: str) -> tuple[str, str, int]:
    if not isinstance(data, dict):
        raise CalcInputError(f"{what} must be an object")
    sheet, id_col, first = data.get("sheet"), data.get("id_column"), data.get("first_row")
    if not isinstance(sheet, str) or not sheet:
        raise CalcInputError(f"{what}.sheet must be a non-empty string")
    if not isinstance(id_col, str) or not id_col:
        raise CalcInputError(f"{what}.id_column must be a non-empty string")
    if not isinstance(first, int) or isinstance(first, bool) or first < 1:
        raise CalcInputError(f"{what}.first_row must be an int >= 1")
    return sheet, id_col, first


def load_check_against(path: Path) -> CheckAgainst:
    """Load ``check.json`` (shape documented in the module docstring)."""
    data = _read_json(path)
    if not isinstance(data, dict):
        raise CalcInputError(f"{path} must contain a JSON object")
    results: Any = data.get("results")
    sheet, id_col, first = _sheet_block(results, "results")
    outputs = _str_map(results["outputs"] if "outputs" in results else None, "results.outputs")
    inputs = data.get("inputs")
    if not isinstance(inputs, dict):
        raise CalcInputError("inputs must be an object")
    has_table, has_selector = "table" in inputs, "selector" in inputs
    if has_table == has_selector:
        raise CalcInputError("inputs needs exactly one of 'table' or 'selector'")
    if has_selector:
        selector = inputs["selector"]
        if not isinstance(selector, str) or not selector.strip():
            raise CalcInputError("inputs.selector must be a non-empty string")
        return CheckAgainst(sheet, id_col, first, outputs, None, selector)
    t_sheet, t_id, t_first = _sheet_block(inputs["table"], "inputs.table")
    fields = _str_map(inputs["table"].get("fields"), "inputs.table.fields")
    table = InputTable(t_sheet, t_id, t_first, fields)
    return CheckAgainst(sheet, id_col, first, outputs, table, None)


def parse_outputs(values: list[str]) -> list[str]:
    """Keep order and drop exact duplicates."""
    return list(dict.fromkeys(values))
