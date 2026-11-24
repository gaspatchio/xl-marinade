"""Dataclasses and enums shared by every calc module."""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import StrEnum
from typing import NamedTuple

from xl_marinade.core.ref_converter import quote_sheet_name
from xl_marinade.errors import CalcInputError

_REF_RE = re.compile(
    r"^(?:(?P<sheet>'(?:[^']|'')*'|[^'!]+)!)?\$?(?P<col>[A-Za-z]{1,3})\$?(?P<row>[0-9]+)$"
)


def _col_to_number(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _number_to_col(n: int) -> str:
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


class CellRef(NamedTuple):
    """A single cell: sheet name, 1-based row, 1-based column."""

    sheet: str
    row: int
    col: int

    def a1(self) -> str:
        """Return ``Sheet!A1`` with the sheet name quoted when needed."""
        return f"{quote_sheet_name(self.sheet)}!{_number_to_col(self.col)}{self.row}"

    @classmethod
    def parse(cls, text: str) -> CellRef:
        """Parse ``Sheet!A1`` / ``'It''s'!$B$3`` into a CellRef.

        Raises:
            CalcInputError: if the text is not a sheet-qualified single-cell reference.
        """
        m = _REF_RE.match(text.strip()) if isinstance(text, str) else None
        if m is None or m.group("sheet") is None:
            raise CalcInputError(f"not a sheet-qualified cell reference: {text!r}")
        sheet = m.group("sheet")
        if sheet.startswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
        row = int(m.group("row"))
        if row < 1:
            raise CalcInputError(f"row must be >= 1 in {text!r}")
        return cls(sheet, row, _col_to_number(m.group("col")))


class Status(StrEnum):
    """Trust status of a cell or run."""

    REFUSED = "refused"
    UNCERTAIN = "uncertain"
    UNVERIFIED = "unverified_path"
    OK = "ok"
    STORED = "stored"


class ReasonKind(StrEnum):
    """Why a cell carries a non-ok status."""

    silently_wrong = "silently_wrong"
    engine_error = "engine_error"
    macro_function = "macro_function"
    unsupported_function = "unsupported_function"
    no_stored_value = "no_stored_value"
    refused_pattern = "refused_pattern"
    check_against_mismatch = "check_against_mismatch"
    near_rounding_flip = "near_rounding_flip"
    near_comparison = "near_comparison"
    unexercised_branch = "unexercised_branch"
    unexercised_lookup_edge = "unexercised_lookup_edge"
    not_checked_path = "not_checked_path"
    not_checked_value = "not_checked_value"
    external_link_held = "external_link_held"
    input_outside_allowed_values = "input_outside_allowed_values"


@dataclass(slots=True, frozen=True, order=True)
class Reason:
    """One reason a cell is not fully trusted."""

    kind: ReasonKind
    root: CellRef | None
    detail: str


SEVERITY: dict[ReasonKind, Status] = {
    ReasonKind.silently_wrong: Status.REFUSED,
    ReasonKind.engine_error: Status.REFUSED,
    ReasonKind.macro_function: Status.REFUSED,
    ReasonKind.unsupported_function: Status.REFUSED,
    ReasonKind.no_stored_value: Status.REFUSED,
    ReasonKind.refused_pattern: Status.REFUSED,
    ReasonKind.check_against_mismatch: Status.REFUSED,
    ReasonKind.input_outside_allowed_values: Status.REFUSED,
    ReasonKind.near_rounding_flip: Status.UNCERTAIN,
    ReasonKind.near_comparison: Status.UNCERTAIN,
    ReasonKind.not_checked_value: Status.UNCERTAIN,
    ReasonKind.unexercised_branch: Status.UNVERIFIED,
    ReasonKind.unexercised_lookup_edge: Status.UNVERIFIED,
    ReasonKind.not_checked_path: Status.UNVERIFIED,
    ReasonKind.external_link_held: Status.OK,
}


@dataclass(slots=True, frozen=True)
class ArrayRange:
    """The cells one array formula fills, as the workbook stored them.

    ``anchor`` is the top-left cell, the only one holding the formula text; the range
    runs to ``last_row`` / ``last_col``. ``dynamic`` marks an Excel 365 dynamic array
    (a spill), whose size can change with the inputs; a legacy (Ctrl+Shift+Enter)
    array keeps its size.
    """

    anchor: CellRef
    last_row: int
    last_col: int
    dynamic: bool

    def cells(self) -> list[CellRef]:
        """Every cell of the range, row by row, the anchor first."""
        a = self.anchor
        return [
            CellRef(a.sheet, r, c)
            for r in range(a.row, self.last_row + 1)
            for c in range(a.col, self.last_col + 1)
        ]


@dataclass(slots=True, frozen=True)
class InputMap:
    """Maps policy-file field names to raw workbook targets (``Sheet!A1`` or a name)."""

    fields: dict[str, str]
    id_field: str | None


@dataclass(slots=True, frozen=True)
class Policy:
    """One policy row: its id (kept as text) and the mapped field values."""

    id: str
    values: dict[str, str]


@dataclass(slots=True, frozen=True)
class VariableSpec:
    """A requested output variable: a cell, or a range with a first period."""

    name: str
    target: str
    period_start: int | None


@dataclass(slots=True, frozen=True)
class InputTable:
    """Where a check-against workbook keeps its per-policy input rows."""

    sheet: str
    id_column: str
    first_row: int
    fields: dict[str, str]


@dataclass(slots=True, frozen=True)
class CheckAgainst:
    """Where stored results (and the inputs that produced them) live."""

    results_sheet: str
    id_column: str
    first_row: int
    outputs: dict[str, str]
    table: InputTable | None
    selector: str | None


@dataclass(slots=True, frozen=True)
class WorkbookSettings:
    """Workbook-level calculation settings that change what a recalculation means."""

    date1904: bool
    full_precision: bool
    iterate: bool
    manual_calc: bool
    strict_ooxml: bool = False
