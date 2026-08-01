# ABOUTME: Fast streaming XML parser for Excel worksheets using iterparse
# ABOUTME: Replaces openpyxl object materialization with constant-memory streaming

import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Iterator
from pathlib import Path

# Excel XML namespaces
SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _ns(tag: str) -> str:
    """Prepend spreadsheet namespace to tag."""
    return f"{{{SPREADSHEET_NS}}}{tag}"


# Pre-computed namespaced tag constants for the streaming hot path.
# Avoids re-formatting these strings ~113M times per large-model run inside
# stream_worksheet_cells.
_NS_C = _ns("c")
_NS_F = _ns("f")
_NS_V = _ns("v")
_NS_IS = _ns("is")
_NS_T = _ns("t")
_NS_SI = _ns("si")


class SharedStringsTable:
    """Lazy loader for shared strings table."""

    def __init__(self, zipf: zipfile.ZipFile):
        self._zipf = zipf
        self._strings: list[str] | None = None

    def _load(self) -> None:
        """Load shared strings table from xl/sharedStrings.xml."""
        if self._strings is not None:
            return

        self._strings = []

        try:
            with self._zipf.open("xl/sharedStrings.xml") as f:
                # Use iterparse to handle large shared string tables
                for event, elem in ET.iterparse(f, events=("end",)):
                    if elem.tag == _ns("si"):  # String item
                        # Extract text from <t> elements
                        text_parts = []
                        for t_elem in elem.findall(f".//{_ns('t')}"):
                            if t_elem.text:
                                text_parts.append(t_elem.text)
                        self._strings.append("".join(text_parts))
                        elem.clear()  # Free memory
        except KeyError:
            # No shared strings table in workbook
            pass

    def get(self, index: int) -> str:
        """
        Get shared string by index.

        Args:
            index: 0-based index into shared strings table

        Returns:
            The shared string at the given index

        Raises:
            ValueError: If index is out of bounds
        """
        self._load()
        if self._strings is None or not (0 <= index < len(self._strings)):
            raise ValueError(
                f"Shared string index {index} out of bounds "
                f"(table has {len(self._strings) if self._strings else 0} entries)"
            )
        return self._strings[index]


def parse_cell_reference(cell_ref: str) -> tuple[int, int]:
    """
    Parse Excel cell reference (e.g., 'A1', 'AA100') to (row, col).
    Returns 1-based row and column numbers.

    Cell references must follow Excel A1 notation: column letters followed by row digits.

    Raises:
        ValueError: If cell reference is invalid or out of Excel bounds
    """
    if not cell_ref:
        raise ValueError("Cell reference cannot be empty")

    # Parse column letters and row digits
    # Excel A1 notation: letters first, then digits
    col_str = ""
    row_str = ""
    seen_digit = False

    for char in cell_ref:
        if char.isalpha():
            if seen_digit:
                # Letters after digits are invalid
                raise ValueError(
                    f"Invalid cell reference '{cell_ref}': "
                    f"letters cannot appear after digits in A1 notation"
                )
            col_str += char
        elif char.isdigit():
            seen_digit = True
            row_str += char
        else:
            raise ValueError(
                f"Invalid cell reference '{cell_ref}': contains invalid character '{char}'"
            )

    if not col_str:
        raise ValueError(f"Invalid cell reference '{cell_ref}': no column letters")

    # Convert column letters to number (A=1, Z=26, AA=27, etc.)
    col = 0
    for char in col_str.upper():
        col = col * 26 + (ord(char) - ord("A") + 1)

    # Validate column bounds (Excel max: XFD = 16384)
    if col < 1 or col > 16384:
        raise ValueError(
            f"Invalid cell reference '{cell_ref}': column {col} out of Excel bounds (1-16384)"
        )

    try:
        row = int(row_str) if row_str else 1
    except ValueError:
        raise ValueError(
            f"Invalid cell reference '{cell_ref}': row part '{row_str}' is not a valid number"
        )

    # Validate row bounds (Excel max: 1048576)
    if row < 1 or row > 1048576:
        raise ValueError(
            f"Invalid cell reference '{cell_ref}': row {row} out of Excel bounds (1-1048576)"
        )

    return row, col


def stream_worksheet_cells(
    workbook_path: str, sheet_name: str
) -> Iterator[tuple[int, int, str | None, str | None, str | None, int | None, bool, int | None]]:
    """
    Stream cells from a worksheet using iterparse.

    Yields tuples of (row, col, formula, value, cell_type, shared_index, shared_master, style_index) where:
    - row, col: 1-based integers
    - formula: formula string without leading '=' (or None)
    - value: cell value as string (or None)
    - cell_type: Excel cell type ('s' for shared string, 'str' for inline string,
                 'n' for number, 'b' for boolean, etc.)
    - shared_index: shared formula index (or None if not shared)
    - shared_master: True if this cell defines the shared formula text
    - style_index: integer index into the workbook's cellXfs style table (or None)

    Args:
        workbook_path: Path to .xlsx/.xlsm file
        sheet_name: Name of worksheet to stream

    Yields:
        (row, col, formula, value, cell_type, shared_index, shared_master, style_index) tuples
    """
    workbook_path = Path(workbook_path)

    with zipfile.ZipFile(workbook_path, "r") as zipf:
        # Load shared strings table
        shared_strings = SharedStringsTable(zipf)

        # Find sheet XML path
        sheet_xml_path = _find_sheet_xml_path(zipf, sheet_name)
        if not sheet_xml_path:
            # Get available sheet names for better error message
            available_sheets = get_sheet_names(str(workbook_path))
            raise ValueError(
                f"Sheet '{sheet_name}' not found in workbook. "
                f"Available sheets: {', '.join(available_sheets) if available_sheets else '(none)'}"
            )

        # Stream parse the worksheet
        with zipf.open(sheet_xml_path) as f:
            for event, elem in ET.iterparse(f, events=("end",)):
                if elem.tag == _NS_C:  # Cell element
                    # Fast path: empty placeholder cells like `<c r="A1" s="180"/>`
                    # have no children (no <f>, <v>, or <is> sub-element), so they
                    # cannot produce IR output. On a large model's main data sheet
                    # that's 31.95M of 31.96M cells. Skipping here avoids ~32M parse_cell_reference
                    # calls + ~64M Element.find() lookups + ~32M attribute fetches
                    # that the rest of the loop body would otherwise pay.
                    if len(elem) == 0:
                        elem.clear()
                        continue

                    cell_ref = elem.get("r")
                    if not cell_ref:
                        elem.clear()
                        continue

                    row, col = parse_cell_reference(cell_ref)
                    cell_type = elem.get("t")  # Cell type attribute

                    # Style index (s attribute) — maps into cellXfs in styles.xml
                    style_index_str = elem.get("s")
                    style_index: int | None = None
                    if style_index_str is not None:
                        try:
                            style_index = int(style_index_str)
                        except ValueError:
                            pass

                    # Extract formula
                    formula = None
                    shared_index = None
                    shared_master = False
                    formula_elem = elem.find(_NS_F)
                    if formula_elem is not None and formula_elem.text:
                        formula = formula_elem.text.strip()
                        # Remove leading '=' if present
                        if formula.startswith("="):
                            formula = formula[1:]
                        if formula_elem.get("t") == "shared":
                            shared_index_str = formula_elem.get("si")
                            if shared_index_str is not None:
                                try:
                                    shared_index = int(shared_index_str)
                                    shared_master = True
                                except ValueError:
                                    shared_index = None
                    elif formula_elem is not None and formula_elem.get("t") == "shared":
                        shared_index_str = formula_elem.get("si")
                        if shared_index_str is not None:
                            try:
                                shared_index = int(shared_index_str)
                            except ValueError:
                                shared_index = None

                    # Extract value
                    value = None
                    value_elem = elem.find(_NS_V)
                    if value_elem is not None and value_elem.text:
                        value_text = value_elem.text.strip()

                        # Resolve shared string references
                        if cell_type == "s":  # Shared string
                            try:
                                idx = int(value_text)
                                value = shared_strings.get(idx)
                            except ValueError:
                                # Malformed: either non-numeric index or out-of-bounds
                                # Fall back to raw value (will be the index as string)
                                value = value_text
                        else:
                            value = value_text

                    # Handle inline strings
                    if cell_type == "inlineStr":
                        is_elem = elem.find(_NS_IS)
                        if is_elem is not None:
                            t_elem = is_elem.find(_NS_T)
                            if t_elem is not None and t_elem.text:
                                value = t_elem.text

                    # Defensive fallback: cells that have children but no formula
                    # text, no value text, and no shared-formula reference still
                    # get filtered here. Rare in practice (the len(elem)==0 fast
                    # path covers the vast majority of empties), but preserves the
                    # original byte-equivalent behaviour for malformed/edge XML.
                    if formula is None and (value is None or value == "") and shared_index is None:
                        elem.clear()
                        continue

                    # Yield cell data
                    yield (
                        row,
                        col,
                        formula,
                        value,
                        cell_type,
                        shared_index,
                        shared_master,
                        style_index,
                    )

                    # Clear element to free memory
                    elem.clear()


def _find_sheet_xml_path(zipf: zipfile.ZipFile, sheet_name: str) -> str | None:
    """
    Find the XML path for a given sheet name.

    Parses xl/workbook.xml to find the sheet relationship ID,
    then resolves it via xl/_rels/workbook.xml.rels.
    """
    # Parse workbook.xml to find sheet relationship ID
    sheet_rel_id = None

    try:
        with zipf.open("xl/workbook.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()

            # Find sheets element
            sheets_elem = root.find(f".//{_ns('sheets')}")
            if sheets_elem is None:
                return None

            # Find matching sheet by name (case-sensitive)
            for sheet_elem in sheets_elem.findall(_ns("sheet")):
                if sheet_elem.get("name") == sheet_name:
                    # Get relationship ID
                    sheet_rel_id = sheet_elem.get(f"{{{RELATIONSHIPS_NS}}}id")
                    break
    except KeyError:
        return None

    if not sheet_rel_id:
        return None

    # Parse workbook.xml.rels to resolve relationship ID to target path
    try:
        with zipf.open("xl/_rels/workbook.xml.rels") as f:
            tree = ET.parse(f)
            root = tree.getroot()

            # Find relationship with matching ID
            for rel_elem in root.findall(f"{{{PACKAGE_RELATIONSHIPS_NS}}}Relationship"):
                if rel_elem.get("Id") == sheet_rel_id:
                    target = rel_elem.get("Target")
                    if target:
                        # Target may have leading / or be relative
                        target = target.lstrip("/")
                        # If target doesn't start with xl/, prepend it
                        if not target.startswith("xl/"):
                            target = f"xl/{target}"
                        return target
    except KeyError:
        pass

    return None


def get_sheet_names(workbook_path: str) -> list[str]:
    """
    Get list of sheet names from workbook in document order.

    Args:
        workbook_path: Path to .xlsx/.xlsm file

    Returns:
        List of sheet names in document order
    """
    workbook_path = Path(workbook_path)
    sheet_names = []

    with zipfile.ZipFile(workbook_path, "r") as zipf:
        try:
            with zipf.open("xl/workbook.xml") as f:
                tree = ET.parse(f)
                root = tree.getroot()

                # Find sheets element
                sheets_elem = root.find(f".//{_ns('sheets')}")
                if sheets_elem is not None:
                    # Collect sheet names in document order
                    for sheet_elem in sheets_elem.findall(_ns("sheet")):
                        name = sheet_elem.get("name")
                        if name:
                            sheet_names.append(name)
        except KeyError:
            pass

    return sheet_names
