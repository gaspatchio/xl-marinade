# ABOUTME: Name/Table Map module that enumerates and resolves Defined Names and ListObjects
# ABOUTME: Maps name references to concrete A1 ranges deterministically per ADR-002

import re
from collections import defaultdict
from dataclasses import dataclass

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table


@dataclass
class DefinedNameInfo:
    """
    Information about a Defined Name.

    Attributes:
        name: Name of the defined name
        range_string: Reference string (e.g., "Sheet1!$A$1:$A$10")
        scope: "workbook" or sheet name for sheet-scoped names
        is_external: True if name references external workbook
        ranges: List of concrete A1 ranges (decomposed if non-contiguous)
        is_dynamic: True if name's value is a formula (dynamic named range)
    """

    name: str
    range_string: str
    scope: str
    is_external: bool
    ranges: list[str]
    is_dynamic: bool = False


@dataclass
class TableInfo:
    """
    Information about a ListObject (Excel table).

    Attributes:
        name: Table name
        sheet: Sheet name containing the table
        range: Table range (including headers)
        columns: List of column names
        header_row: A1 range of header row
        data_range: A1 range of data body (excluding headers)
    """

    name: str
    sheet: str
    range: str
    columns: list[str]
    header_row: str
    data_range: str


class NameTableMap:
    """
    Map of Defined Names and ListObjects in a workbook.

    Provides:
    - Enumeration of all names and tables
    - Resolution of name references to concrete ranges
    - Resolution of structured table references
    - Deterministic ordering per ADR-000
    """

    def __init__(self, workbook: Workbook) -> None:
        """
        Initialize name/table map from workbook.

        Note: Reverse maps (cell → names, cell → tables) are built lazily on first access
        to avoid expanding all defined name ranges upfront. This significantly reduces
        initialization cost for workbooks with many large defined names.

        Args:
            workbook: openpyxl Workbook object
        """
        self.workbook = workbook
        self._names: dict[str, DefinedNameInfo] = {}
        self._tables: dict[str, TableInfo] = {}
        self._cell_to_names: dict[str, list[str]] = defaultdict(list)
        self._cell_to_tables: dict[str, str] = {}

        # Lazy initialization flag for reverse maps
        self._reverse_maps_built = False

        # Build forward maps (name → ranges, table → info) - these are cheap
        self._enumerate_names()
        self._enumerate_tables()

        # REMOVED: self._build_reverse_maps()  # Now built lazily on first access

    def _enumerate_names(self) -> None:
        """
        Enumerate all Defined Names in workbook.

        Handles:
        - Workbook-scoped names
        - Sheet-scoped names
        - Non-contiguous ranges (decomposed per ADR-002)
        - External references (marked as unresolved)
        """
        # Workbook-scoped defined names
        # DefinedNameDict supports iteration over keys (name strings)
        for name_key in sorted(self.workbook.defined_names):
            name = self.workbook.defined_names[name_key]
            self._process_defined_name(name, scope="workbook")

        # Sheet-scoped defined names
        for sheet in self.workbook.worksheets:
            if hasattr(sheet, "defined_names") and sheet.defined_names:
                for name_key in sorted(sheet.defined_names):
                    name = sheet.defined_names[name_key]
                    self._process_defined_name(name, scope=sheet.title)

    def _process_defined_name(self, name: DefinedName, scope: str) -> None:
        """
        Process a single defined name.

        Args:
            name: DefinedName object from openpyxl
            scope: "workbook" or sheet name
        """
        name_str = name.name
        range_string = str(name.value) if name.value else ""

        # Check for external references
        is_external = self._is_external_reference(range_string)

        # Detect formula-valued (dynamic) names; these must NOT be comma-split into garbage.
        is_dynamic = (not is_external) and self._is_formula_valued(range_string)

        # Decompose range into contiguous parts (skip for external/dynamic names)
        if is_external or is_dynamic:
            ranges = []
        else:
            ranges = self._decompose_range(range_string)

        # Store with deterministic key (workbook-scoped names take precedence)
        key = f"{scope}::{name_str}" if scope != "workbook" else name_str

        self._names[key] = DefinedNameInfo(
            name=name_str,
            range_string=range_string,
            scope=scope,
            is_external=is_external,
            ranges=ranges,
            is_dynamic=is_dynamic,
        )

    def _is_formula_valued(self, range_string: str) -> bool:
        """
        Detect whether a defined name's value is a formula rather than a plain range.

        A plain cell/range reference (including non-contiguous comma forms and quoted
        sheet names that may contain parentheses) has no unquoted '('. A formula-valued
        name (e.g. OFFSET(...), INDEX(...)) does.

        Args:
            range_string: The defined name's value string

        Returns:
            True if the value contains an unquoted '(' (i.e. is a formula)
        """
        if not range_string:
            return False
        # Mask single-quoted sheet-name segments so parens inside quoted names don't count.
        masked = re.sub(r"'[^']*'", "", range_string)
        return "(" in masked

    def _is_external_reference(self, range_string: str) -> bool:
        """
        Check if range references external workbook.

        Args:
            range_string: Reference string

        Returns:
            True if external reference (e.g., "[Book2.xlsx]Sheet1!A1")
        """
        # External references have pattern [*.xlsx] or [*.xls] before sheet name
        # More robust than just checking for brackets (avoids false positives with array formulas)
        if not range_string:
            return False
        # Look for workbook name in brackets followed by sheet reference
        return bool(re.search(r"\[.+\.(xlsx?|xlsm|xltx?|xlam)\]", range_string, re.IGNORECASE))

    def _decompose_range(self, range_string: str) -> list[str]:
        """
        Decompose range into contiguous parts.

        Handles:
        - Non-contiguous ranges (A1:A10,C1:C10) → ["A1:A10", "C1:C10"]
        - Single cells (A1) → ["A1"]
        - Contiguous ranges (A1:B10) → ["A1:B10"]

        Args:
            range_string: Reference string

        Returns:
            List of contiguous range strings (without $ characters)
        """
        if not range_string:
            return []

        # Remove $ characters first (absolute markers)
        range_string = range_string.replace("$", "")

        # Split on comma (non-contiguous separator)
        # Note: Each part may have its own sheet qualification
        ranges = []
        for part in range_string.split(","):
            part = part.strip()
            if part:
                ranges.append(part)

        return sorted(ranges)  # Deterministic ordering

    def _enumerate_tables(self) -> None:
        """
        Enumerate all ListObjects (Excel tables) in workbook.
        """
        for sheet in self.workbook.worksheets:
            if not hasattr(sheet, "_tables") or sheet._tables is None:
                continue

            # _tables is a dict mapping table name to table object
            table_names = sorted(sheet._tables.keys()) if isinstance(sheet._tables, dict) else []
            for table_name in table_names:
                table = sheet._tables[table_name]
                self._process_table(table, sheet.title)

    def _process_table(self, table: Table, sheet_name: str) -> None:
        """
        Process a single table.

        Args:
            table: Table object from openpyxl
            sheet_name: Sheet containing the table
        """
        # Extract table info
        name = table.name
        range_str = table.ref  # e.g., "A1:C10"

        # Parse range
        if ":" not in range_str:
            # Single cell table (unusual but possible)
            header_row = range_str
            data_range = ""
        else:
            # Extract header and data ranges
            start, end = range_str.split(":", 1)
            # Parse to get row numbers
            start_match = re.match(r"^(\$?)([A-Z]+)(\$?)(\d+)$", start, re.IGNORECASE)
            end_match = re.match(r"^(\$?)([A-Z]+)(\$?)(\d+)$", end, re.IGNORECASE)

            if start_match and end_match:
                start_col = start_match.group(2)
                start_row = int(start_match.group(4))
                end_col = end_match.group(2)
                end_row = int(end_match.group(4))

                # Header is first row
                header_row = f"{start_col}{start_row}:{end_col}{start_row}"

                # Data is remaining rows
                if end_row > start_row:
                    data_range = f"{start_col}{start_row + 1}:{end_col}{end_row}"
                else:
                    data_range = ""
            else:
                header_row = range_str
                data_range = ""

        # Get column names
        columns = []
        if table.tableColumns:
            columns = [col.name for col in table.tableColumns]

        # Qualify with sheet name
        qualified_range = f"{sheet_name}!{range_str}"
        qualified_header = f"{sheet_name}!{header_row}"
        qualified_data = f"{sheet_name}!{data_range}" if data_range else ""

        self._tables[name] = TableInfo(
            name=name,
            sheet=sheet_name,
            range=qualified_range,
            columns=columns,
            header_row=qualified_header,
            data_range=qualified_data,
        )

    def _ensure_reverse_maps_built(self) -> None:
        """
        Build reverse maps on first access (lazy initialization).

        This is called lazily when get_cell_defined_names() or get_cell_table_ref()
        is first called, avoiding the upfront cost of expanding all defined name
        and table ranges if these methods are never used.

        The reverse maps are expensive to build (O(cells in all names and tables))
        but only needed for populating extras.defined_name and extras.table_ref
        during cell serialization.

        Exception Safety: If build fails, clears partial state and re-raises.
        Next call will retry from clean state.
        """
        if self._reverse_maps_built:
            return

        try:
            self._build_reverse_maps()
            self._reverse_maps_built = True
        except Exception:
            # Clear partial state on failure to ensure clean retry
            self._cell_to_names.clear()
            self._cell_to_tables.clear()
            raise

    def _build_reverse_maps(self) -> None:
        """
        Build reverse maps: cell → names and cell → table.

        This enables populating extras.defined_name and extras.table_ref.
        Called lazily via _ensure_reverse_maps_built().
        """
        # Map cells to defined names
        for key, name_info in self._names.items():
            if name_info.is_external:
                continue

            for range_str in name_info.ranges:
                cells = self._expand_range_to_cells(range_str)
                for cell in cells:
                    self._cell_to_names[cell].append(name_info.name)

        # Sort name lists for determinism
        for cell in self._cell_to_names:
            self._cell_to_names[cell] = sorted(self._cell_to_names[cell])

        # Map cells to tables
        for table_name, table_info in self._tables.items():
            if not table_info.data_range:
                continue

            cells = self._expand_range_to_cells(table_info.data_range)

            # Map each cell to table.column
            for cell in cells:
                col_name = self._get_table_column_for_cell(cell, table_info)
                if col_name:
                    self._cell_to_tables[cell] = f"{table_name}.{col_name}"

    def _expand_range_to_cells(self, range_str: str) -> list[str]:
        """
        Expand A1 range to individual cell addresses.

        Args:
            range_str: A1 range (e.g., "Sheet1!A1:B2")

        Returns:
            List of cell addresses (e.g., ["Sheet1!A1", "Sheet1!B1", "Sheet1!A2", "Sheet1!B2"])
        """
        # Handle sheet qualification
        sheet_prefix = ""
        if "!" in range_str:
            parts = range_str.rsplit("!", 1)
            sheet_prefix = parts[0] + "!"
            range_str = parts[1]

        # Remove $ characters
        range_str = range_str.replace("$", "")

        # Single cell
        if ":" not in range_str:
            return [sheet_prefix + range_str]

        # Parse range
        start, end = range_str.split(":", 1)
        start_match = re.match(r"^([A-Z]+)(\d+)$", start, re.IGNORECASE)
        end_match = re.match(r"^([A-Z]+)(\d+)$", end, re.IGNORECASE)

        if not start_match or not end_match:
            return []

        start_col = self._col_letter_to_num(start_match.group(1))
        start_row = int(start_match.group(2))
        end_col = self._col_letter_to_num(end_match.group(1))
        end_row = int(end_match.group(2))

        # Expand to cells
        cells = []
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                col_letter = self._col_num_to_letter(col)
                cells.append(f"{sheet_prefix}{col_letter}{row}")

        return cells

    def _col_letter_to_num(self, col_letter: str) -> int:
        """Convert column letter to number (A=1, Z=26, AA=27)"""
        num = 0
        for char in col_letter.upper():
            num = num * 26 + (ord(char) - ord("A") + 1)
        return num

    def _col_num_to_letter(self, col_num: int) -> str:
        """Convert column number to letter (1=A, 26=Z, 27=AA)"""
        letters = []
        while col_num > 0:
            col_num -= 1
            letters.append(chr(col_num % 26 + ord("A")))
            col_num //= 26
        return "".join(reversed(letters))

    def _get_table_column_for_cell(self, cell: str, table_info: TableInfo) -> str | None:
        """
        Determine which column a cell belongs to in a table.

        Args:
            cell: Cell address (e.g., "Sheet1!B2")
            table_info: Table information

        Returns:
            Column name or None
        """
        # Extract column from cell address
        if "!" in cell:
            cell = cell.split("!", 1)[1]

        match = re.match(r"^([A-Z]+)(\d+)$", cell, re.IGNORECASE)
        if not match:
            return None

        col_letter = match.group(1).upper()
        col_num = self._col_letter_to_num(col_letter)

        # Parse table range to find column index
        table_range = table_info.range
        if "!" in table_range:
            table_range = table_range.split("!", 1)[1]

        start = table_range.split(":", 1)[0].replace("$", "")
        start_match = re.match(r"^([A-Z]+)(\d+)$", start, re.IGNORECASE)
        if not start_match:
            return None

        start_col_num = self._col_letter_to_num(start_match.group(1))
        col_index = col_num - start_col_num

        if 0 <= col_index < len(table_info.columns):
            return table_info.columns[col_index]

        return None

    def resolve_name(self, name: str, scope: str = "workbook") -> list[str] | None:
        """
        Resolve a defined name to concrete A1 ranges.

        Workbook-scoped names take precedence over sheet-scoped names.

        Args:
            name: Name to resolve
            scope: Sheet scope (if sheet-scoped) or "workbook"

        Returns:
            List of A1 ranges, or None if not found/unresolved

        Example:
            >>> name_map.resolve_name("Revenue")
            ['Sheet1!B5:B10']
        """
        # If an explicit sheet scope is provided, prefer the sheet-scoped name.
        if scope and scope != "workbook":
            sheet_key = f"{scope}::{name}"
            if sheet_key in self._names:
                info = self._names[sheet_key]
                if info.is_external:
                    return None
                return info.ranges

        # Otherwise, workbook-scoped names take precedence.
        if name in self._names:
            info = self._names[name]
            if info.is_external:
                return None  # External references are unresolved
            return info.ranges

        # Fallback: sheet scope (useful when caller passes scope="workbook" but only a sheet-scoped
        # name exists in the file).
        sheet_key = f"{scope}::{name}"
        if sheet_key in self._names:
            info = self._names[sheet_key]
            if info.is_external:
                return None
            return info.ranges

        return None

    def is_dynamic_name(self, name: str, scope: str = "workbook") -> bool:
        """
        Report whether a defined name is formula-valued (dynamic).

        Uses the same key-precedence as resolve_name so callers get a consistent
        answer for the name they just failed to resolve.

        Args:
            name: Name to check
            scope: Sheet scope (if sheet-scoped) or "workbook"

        Returns:
            True if the matched defined name is dynamic; False otherwise.
        """
        if scope and scope != "workbook":
            sheet_key = f"{scope}::{name}"
            if sheet_key in self._names:
                return self._names[sheet_key].is_dynamic

        if name in self._names:
            return self._names[name].is_dynamic

        sheet_key = f"{scope}::{name}"
        if sheet_key in self._names:
            return self._names[sheet_key].is_dynamic

        return False

    def resolve_table_reference(self, table_ref: str) -> list[str] | None:
        """
        Resolve structured table reference to concrete cells.

        Supports:
        - Table1[Column1] → cells in Column1
        - Table1[[#All],[Column1]] → all cells including header
        - Table1[#Headers] → header row
        - Table1[#Data] → data rows only

        Args:
            table_ref: Structured reference string

        Returns:
            List of cell addresses, or None if unresolved

        Example:
            >>> name_map.resolve_table_reference("Sales[Revenue]")
            ['Sheet1!B2:B10']
        """
        # Parse table reference
        match = re.match(r"^(\w+)\[(.+)\]$", table_ref)
        if not match:
            return None

        table_name = match.group(1)
        column_spec = match.group(2)

        # Look up table
        if table_name not in self._tables:
            return None

        table_info = self._tables[table_name]

        # Handle special specifiers
        if column_spec == "#Headers":
            return [table_info.header_row]
        elif column_spec == "#Data":
            return [table_info.data_range] if table_info.data_range else []
        elif column_spec == "#All":
            return [table_info.range]

        # Handle column reference
        # Remove special notation like [[#All],[Column]]
        column_spec = column_spec.replace("[[#All],", "").replace("[", "").replace("]", "")

        # Find column index
        if column_spec not in table_info.columns:
            return None

        col_index = table_info.columns.index(column_spec)

        # Build column range
        data_range = table_info.data_range
        if not data_range:
            return []

        # Parse data range to extract column
        if "!" in data_range:
            sheet, range_part = data_range.rsplit("!", 1)
        else:
            sheet = table_info.sheet
            range_part = data_range

        start, end = range_part.replace("$", "").split(":", 1)
        start_match = re.match(r"^([A-Z]+)(\d+)$", start, re.IGNORECASE)
        end_match = re.match(r"^([A-Z]+)(\d+)$", end, re.IGNORECASE)

        if not start_match or not end_match:
            return []

        start_col_num = self._col_letter_to_num(start_match.group(1))
        target_col_num = start_col_num + col_index
        target_col_letter = self._col_num_to_letter(target_col_num)

        start_row = int(start_match.group(2))
        end_row = int(end_match.group(2))

        column_range = f"{sheet}!{target_col_letter}{start_row}:{target_col_letter}{end_row}"
        return [column_range]

    def get_cell_defined_names(self, cell: str) -> list[str]:
        """
        Get all defined names that include a cell.

        Args:
            cell: Cell address (e.g., "Sheet1!A1")

        Returns:
            List of defined name strings (for extras.defined_name)

        Note: Triggers lazy reverse map build on first call.
        """
        self._ensure_reverse_maps_built()
        return self._cell_to_names.get(cell, [])

    def get_cell_table_ref(self, cell: str) -> str | None:
        """
        Get table reference for a cell.

        Args:
            cell: Cell address (e.g., "Sheet1!A1")

        Returns:
            Table reference string "TableName.ColumnName" or None

        Note: Triggers lazy reverse map build on first call.
        """
        self._ensure_reverse_maps_built()
        return self._cell_to_tables.get(cell)

    def get_all_names(self) -> list[DefinedNameInfo]:
        """Get all defined names (sorted deterministically)"""
        return [self._names[key] for key in sorted(self._names.keys())]

    def get_all_tables(self) -> list[TableInfo]:
        """Get all tables (sorted deterministically)"""
        return [self._tables[key] for key in sorted(self._tables.keys())]


def create_name_table_map(workbook: Workbook) -> NameTableMap:
    """
    Create name/table map from workbook.

    Args:
        workbook: openpyxl Workbook object

    Returns:
        NameTableMap instance
    """
    return NameTableMap(workbook)
