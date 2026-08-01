# ABOUTME: Workbook catalog for deterministic sheet_id assignment from workbook.xml.
# ABOUTME: Parses sheet order, assigns 1-based sheet_id, and provides stable hashing.

"""
Workbook Catalog

Parses xl/workbook.xml to assign deterministic sheet_id values:
- Document order is authoritative (not sheetId attribute)
- sheet_id is 1-based, sequential, no gaps
- Only worksheets are included (not chartsheets, dialogsheets, macrosheets)
- Duplicate sheet names (case-insensitive) cause abort
- Missing relationships cause abort

The catalog produces a SHA256 hash for determinism diagnostics.
"""

import hashlib
import json
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any
from zipfile import ZipFile

# Excel namespace constants
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

# Relationship type for worksheets
WORKSHEET_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

# Import MAX_SHEET_ID from cell_identity for validation
from xl_marinade.core.new_arch.cell_identity import MAX_SHEET_ID


class WorkbookCatalog:
    """
    Workbook catalog with deterministic sheet_id assignment.

    Attributes:
        sheets: List of (sheet_id, sheet_name, rel_id, target) tuples in document order
        sheet_by_name: Dict mapping lowercase sheet_name -> sheet_id (for case-insensitive lookup)
        sheet_by_id: Dict mapping sheet_id -> sheet_name
    """

    def __init__(self, xlsx_path: str | Path):
        """
        Load workbook catalog from .xlsx/.xlsm file.

        Args:
            xlsx_path: Path to Excel workbook

        Raises:
            ValueError: if workbook.xml is invalid, sheets are duplicated, or relationships are missing
        """
        self.xlsx_path = Path(xlsx_path)
        self.sheets: list[tuple[int, str, str, str]] = []
        self.sheet_by_name: dict[str, int] = {}
        self.sheet_by_id: dict[int, str] = {}

        self._load_catalog()

    def _load_catalog(self) -> None:
        """Parse workbook.xml and assign sheet_id values."""
        with ZipFile(self.xlsx_path, "r") as zf:
            # Load workbook relationships
            try:
                rels_xml = zf.read("xl/_rels/workbook.xml.rels")
            except KeyError:
                raise ValueError("Missing xl/_rels/workbook.xml.rels")

            rels_root = ET.fromstring(rels_xml)
            rels_map = {}

            for rel in rels_root.findall(f"{{{NS_PKG_REL}}}Relationship"):
                rel_id = rel.get("Id")
                rel_type = rel.get("Type")
                target = rel.get("Target")

                if rel_id and rel_type and target:
                    rels_map[rel_id] = (rel_type, target)

            # Load workbook.xml
            try:
                workbook_xml = zf.read("xl/workbook.xml")
            except KeyError:
                raise ValueError("Missing xl/workbook.xml")

            workbook_root = ET.fromstring(workbook_xml)
            sheets_elem = workbook_root.find(f"{{{NS_MAIN}}}sheets")

            if sheets_elem is None:
                raise ValueError("No <sheets> element in workbook.xml")

            # Process sheets in document order
            sheet_id = 1
            seen_names = set()

            # Count total sheets first to validate against bit limit
            total_sheets = len(sheets_elem.findall(f"{{{NS_MAIN}}}sheet"))
            if total_sheets > MAX_SHEET_ID:
                raise ValueError(
                    f"Too many sheets: {total_sheets} (max {MAX_SHEET_ID} due to 20-bit sheet_id limit)"
                )

            for sheet_elem in sheets_elem.findall(f"{{{NS_MAIN}}}sheet"):
                sheet_name = sheet_elem.get("name")
                rel_id = sheet_elem.get(f"{{{NS_REL}}}id")

                if not sheet_name:
                    raise ValueError("Sheet element missing 'name' attribute")

                if not rel_id:
                    raise ValueError(f"Sheet '{sheet_name}' missing r:id attribute")

                # Check for duplicate names (case-insensitive)
                sheet_name_lower = sheet_name.lower()
                if sheet_name_lower in seen_names:
                    raise ValueError(f"Duplicate sheet name (case-insensitive): '{sheet_name}'")

                seen_names.add(sheet_name_lower)

                # Resolve relationship
                if rel_id not in rels_map:
                    raise ValueError(f"Sheet '{sheet_name}' has unresolved relationship: {rel_id}")

                rel_type, target = rels_map[rel_id]

                # Only include worksheets (not chartsheets, dialogsheets, macrosheets)
                # Use exact match to avoid substring false positives
                if rel_type != WORKSHEET_REL_TYPE:
                    continue

                # Store sheet info
                self.sheets.append((sheet_id, sheet_name, rel_id, target))
                self.sheet_by_name[sheet_name_lower] = sheet_id
                self.sheet_by_id[sheet_id] = sheet_name

                sheet_id += 1

            if not self.sheets:
                raise ValueError("No worksheets found in workbook")

    def get_sheet_id(self, sheet_name: str) -> int | None:
        """
        Get sheet_id for a sheet name (case-insensitive).

        Args:
            sheet_name: Sheet name to lookup

        Returns:
            sheet_id (1-based) or None if not found
        """
        return self.sheet_by_name.get(sheet_name.lower())

    def get_sheet_name(self, sheet_id: int) -> str | None:
        """
        Get sheet name for a sheet_id.

        Args:
            sheet_id: 1-based sheet identifier

        Returns:
            Sheet name or None if not found
        """
        return self.sheet_by_id.get(sheet_id)

    def compute_hash(self) -> str:
        """
        Compute deterministic SHA256 hash of the catalog.

        Returns:
            Hex-encoded SHA256 hash
        """
        # Serialize catalog as canonical JSON
        catalog_data = [
            {"sheet_id": sheet_id, "sheet_name": sheet_name, "rel_id": rel_id, "target": target}
            for sheet_id, sheet_name, rel_id, target in self.sheets
        ]

        canonical_json = json.dumps(
            catalog_data, ensure_ascii=False, sort_keys=True, separators=(",", ":")
        )

        return hashlib.sha256(canonical_json.encode("utf-8")).hexdigest()

    def to_dict(self) -> dict[str, Any]:
        """
        Export catalog as a dictionary.

        Returns:
            Dict with 'sheets' list and 'hash'
        """
        return {
            "sheets": [
                {"sheet_id": sheet_id, "sheet_name": sheet_name, "rel_id": rel_id, "target": target}
                for sheet_id, sheet_name, rel_id, target in self.sheets
            ],
            "hash": self.compute_hash(),
        }
