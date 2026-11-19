"""Read workbook calculation settings from ``xl/workbook.xml``."""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from xl_marinade.calc.models import WorkbookSettings
from xl_marinade.errors import CalcInputError

__all__ = ["WorkbookSettings", "read_settings", "refusal_reasons"]

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_TRUE = {"1", "true"}


def _flag(value: str | None, default: bool) -> bool:
    if value is None:
        return default
    return value.strip().lower() in _TRUE


def read_settings(path: Path) -> WorkbookSettings:
    """Read the calculation settings of an ``.xlsx`` workbook.

    Absent attributes take Excel's defaults (``fullPrecision`` is true).

    Raises:
        CalcInputError: if the file is not a readable workbook package.
    """
    try:
        with zipfile.ZipFile(path) as zf:
            root = ET.fromstring(zf.read("xl/workbook.xml"))
    except (zipfile.BadZipFile, KeyError, ET.ParseError, OSError) as exc:
        raise CalcInputError(f"cannot read workbook settings from {path}: {exc}") from exc

    pr = root.find(f"{_NS}workbookPr")
    calc = root.find(f"{_NS}calcPr")
    pr_attrs = pr.attrib if pr is not None else {}
    calc_attrs = calc.attrib if calc is not None else {}
    return WorkbookSettings(
        date1904=_flag(pr_attrs.get("date1904"), False),
        full_precision=_flag(calc_attrs.get("fullPrecision"), True),
        iterate=_flag(calc_attrs.get("iterate"), False),
        manual_calc=calc_attrs.get("calcMode") == "manual",
    )


def refusal_reasons(s: WorkbookSettings) -> list[str]:
    """Return the settings that make the calculator refuse the whole workbook."""
    reasons: list[str] = []
    if s.date1904:
        reasons.append("1904 date system")
    if not s.full_precision:
        reasons.append("precision as displayed")
    return reasons
