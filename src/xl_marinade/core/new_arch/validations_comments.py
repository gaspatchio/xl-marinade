# ABOUTME: Extracts data-validation rules (native + x14/extLst) and cell comments
# ABOUTME: (classic + threaded) from workbook XML into the data_validations / cell_comments tables.

"""
Data-validation + cell-comment extraction (2026-07-05 scope-gap channels #1/#2).

Additive channel: populates two standalone tables keyed by sheet_id. No
traversal / bindings / cells impact. Zero fabrication: sqref and formulas are
stored verbatim; range/name-sourced lists are NOT resolved to values at
extraction time (the ref itself is the declared semantics).

Parsed parts:
- per-sheet XML `<dataValidations>` (type/operator/sqref attrs, formula1/2 children)
- per-sheet `<extLst>` `<x14:dataValidation>` (xm:f formulas, xm:sqref) — the
  scenario-selection wiring in the corpus exists ONLY in this form
- `xl/comments*.xml` (classic notes; resolved to sheets via sheet _rels)
- `xl/threadedComments/*.xml` (ref attr carries the cell; dT ordering)
"""

import io
import json
import posixpath
import sqlite3
import xml.etree.ElementTree as ET
import zipfile
from typing import Any

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_XM = "http://schemas.microsoft.com/office/excel/2006/main"
NS_TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

_NS_ROW = f"{{{NS_MAIN}}}row"
_NS_DV = f"{{{NS_MAIN}}}dataValidation"
_NS_FORMULA1 = f"{{{NS_MAIN}}}formula1"
_NS_FORMULA2 = f"{{{NS_MAIN}}}formula2"
_NS_X14_DV = f"{{{NS_X14}}}dataValidation"
_NS_X14_FORMULA1 = f"{{{NS_X14}}}formula1"
_NS_X14_FORMULA2 = f"{{{NS_X14}}}formula2"
_NS_XM_F = f"{{{NS_XM}}}f"
_NS_XM_SQREF = f"{{{NS_XM}}}sqref"


def parse_literal_list(formula1: str | None) -> list[str] | None:
    """Parse a quoted literal list formula1 like '"CSO 2017, CSO 2015, UW"'.

    Returns the trimmed values (["CSO 2017", "CSO 2015", "UW"]) or None when
    formula1 is not a quoted literal (range / defined-name / expression source).
    """
    if not formula1:
        return None
    text = formula1.strip()
    if len(text) < 2 or not (text.startswith('"') and text.endswith('"')):
        return None
    return [part.strip() for part in text[1:-1].split(",")]


def _truthy_attr(value: str | None) -> int:
    return 1 if value in ("1", "true") else 0


def _validation_record(
    *,
    sqref: str | None,
    val_type: str | None,
    operator: str | None,
    formula1: str | None,
    formula2: str | None,
    allow_blank: int,
    prompt_title: str | None,
    prompt_text: str | None,
    source: str,
) -> dict[str, Any] | None:
    if not sqref:
        return None
    literal_values = parse_literal_list(formula1)
    return {
        "sqref": sqref,
        "val_type": val_type,
        "operator": operator,
        "formula1": formula1,
        "formula2": formula2,
        "literal_values_json": (
            json.dumps(literal_values, ensure_ascii=False) if literal_values is not None else None
        ),
        "allow_blank": allow_blank,
        "prompt_title": prompt_title,
        "prompt_text": prompt_text,
        "source": source,
    }


def parse_sheet_validations(sheet_xml: Any) -> list[dict[str, Any]]:
    """Parse native <dataValidation> + x14:dataValidation rules from sheet XML.

    Accepts bytes or a binary file-like. Streams with iterparse (sheet XML can
    be huge); row elements are cleared as they complete so memory stays bounded
    by one row of cells. Returns records in document order, native rules before
    x14 (native <dataValidations> precedes <extLst> in the sheet part).
    """
    records: list[dict[str, Any]] = []
    source = io.BytesIO(sheet_xml) if isinstance(sheet_xml, (bytes, bytearray)) else sheet_xml
    for _event, elem in ET.iterparse(source, events=("end",)):
        if elem.tag == _NS_ROW:
            elem.clear()
        elif elem.tag == _NS_DV:
            f1 = elem.find(_NS_FORMULA1)
            f2 = elem.find(_NS_FORMULA2)
            rec = _validation_record(
                sqref=elem.get("sqref"),
                val_type=elem.get("type"),
                operator=elem.get("operator"),
                formula1=f1.text if f1 is not None else None,
                formula2=f2.text if f2 is not None else None,
                allow_blank=_truthy_attr(elem.get("allowBlank")),
                prompt_title=elem.get("promptTitle"),
                prompt_text=elem.get("prompt"),
                source="native",
            )
            if rec:
                records.append(rec)
            elem.clear()
        elif elem.tag == _NS_X14_DV:
            sqref_elem = elem.find(_NS_XM_SQREF)
            f1_wrap = elem.find(_NS_X14_FORMULA1)
            f2_wrap = elem.find(_NS_X14_FORMULA2)
            f1 = f1_wrap.find(_NS_XM_F) if f1_wrap is not None else None
            f2 = f2_wrap.find(_NS_XM_F) if f2_wrap is not None else None
            rec = _validation_record(
                sqref=sqref_elem.text if sqref_elem is not None else None,
                val_type=elem.get("type"),
                operator=elem.get("operator"),
                formula1=f1.text if f1 is not None else None,
                formula2=f2.text if f2 is not None else None,
                allow_blank=_truthy_attr(elem.get("allowBlank")),
                prompt_title=elem.get("promptTitle"),
                prompt_text=elem.get("prompt"),
                source="x14",
            )
            if rec:
                records.append(rec)
            elem.clear()
    return records


def parse_persons(person_xml: bytes) -> dict[str, str]:
    """Parse xl/persons/person*.xml into {person_id: display_name}."""
    persons: dict[str, str] = {}
    root = ET.fromstring(person_xml)
    for person in root.iter(f"{{{NS_TC}}}person"):
        pid = person.get("id")
        name = person.get("displayName")
        if pid and name:
            persons[pid] = name
    return persons


def parse_threaded_comments(
    tc_xml: bytes, persons: dict[str, str] | None = None
) -> list[dict[str, Any]]:
    """Parse a threadedComments part into comment records with thread_order.

    thread_order is the 0-based position within the cell's thread, ordered by
    the dT timestamp attribute (document order breaks ties / missing dT).
    """
    persons = persons or {}
    raw: list[dict[str, Any]] = []
    root = ET.fromstring(tc_xml)
    for idx, tc in enumerate(root.iter(f"{{{NS_TC}}}threadedComment")):
        ref = tc.get("ref")
        if not ref:
            continue
        text_elem = tc.find(f"{{{NS_TC}}}text")
        text = (text_elem.text or "") if text_elem is not None else ""
        raw.append(
            {
                "a1": ref,
                "author": persons.get(tc.get("personId") or ""),
                "text": text,
                "kind": "threaded",
                "_dt": tc.get("dT") or "",
                "_doc_order": idx,
            }
        )
    raw.sort(key=lambda r: (r["a1"], r["_dt"], r["_doc_order"]))
    by_cell: dict[str, int] = {}
    for rec in raw:
        rec["thread_order"] = by_cell.get(rec["a1"], 0)
        by_cell[rec["a1"]] = rec["thread_order"] + 1
        del rec["_dt"]
        del rec["_doc_order"]
    return raw


def parse_classic_comments(comments_xml: bytes) -> list[dict[str, Any]]:
    """Parse a classic xl/comments*.xml part into comment records.

    Text is the concatenation of all <t> runs under <text> (rich-text runs
    already carry their own whitespace).
    """
    root = ET.fromstring(comments_xml)
    authors = [(a.text or "") for a in root.findall(f"{{{NS_MAIN}}}authors/{{{NS_MAIN}}}author")]
    records: list[dict[str, Any]] = []
    for comment in root.iter(f"{{{NS_MAIN}}}comment"):
        ref = comment.get("ref")
        if not ref:
            continue
        text_elem = comment.find(f"{{{NS_MAIN}}}text")
        parts: list[str] = []
        if text_elem is not None:
            for t in text_elem.iter(f"{{{NS_MAIN}}}t"):
                if t.text:
                    parts.append(t.text)
        author: str | None = None
        author_id = comment.get("authorId")
        if author_id is not None:
            try:
                author = authors[int(author_id)] or None
            except (ValueError, IndexError):
                author = None
        # Legacy mirrors of threaded comments encode the thread uid as the
        # author ("tc={...}"); that is not a human author name.
        if author and author.startswith("tc="):
            author = None
        records.append(
            {
                "a1": ref,
                "author": author,
                "text": "".join(parts),
                "kind": "classic",
                "thread_order": 0,
            }
        )
    return records


def _sheet_rel_targets(zipf: zipfile.ZipFile, sheet_target: str) -> dict[str, list[str]]:
    """Resolve a sheet's comments / threadedComment relationship targets.

    Returns {"comments": [zip paths], "threaded": [zip paths]} with targets
    normalized relative to the sheet part's directory.
    """
    out: dict[str, list[str]] = {"comments": [], "threaded": []}
    sheet_dir = posixpath.dirname(sheet_target)
    rels_path = posixpath.join(sheet_dir, "_rels", posixpath.basename(sheet_target) + ".rels")
    try:
        rels_xml = zipf.read(rels_path)
    except KeyError:
        return out
    root = ET.fromstring(rels_xml)
    for rel in root.findall(f"{{{NS_PKG_REL}}}Relationship"):
        rel_type = rel.get("Type") or ""
        target = rel.get("Target")
        if not target:
            continue
        resolved = posixpath.normpath(posixpath.join(sheet_dir, target.lstrip("/")))
        if rel_type.endswith("/comments"):
            out["comments"].append(resolved)
        elif rel_type.endswith("/threadedComment"):
            out["threaded"].append(resolved)
    return out


def _normalize_part_path(target: str) -> str:
    """Normalize a workbook-rels sheet target like 'worksheets/sheet2.xml' to a zip path."""
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = f"xl/{target}"
    return target


def _zip_part_contains(zipf: zipfile.ZipFile, part: str, needle: bytes) -> bool:
    """Chunked byte scan for `needle` in a zip part without XML parsing.

    Perf guard, not correctness: sheet XML can be hundreds of MB (a large
    model's main data sheet); a full iterparse pass just to discover there
    are no dataValidations would dominate the stage. Decompression-only scan
    is cheap.
    """
    tail = b""
    with zipf.open(part) as handle:
        while True:
            chunk = handle.read(1 << 20)
            if not chunk:
                return False
            if needle in tail + chunk:
                return True
            tail = chunk[-(len(needle) - 1) :] if len(needle) > 1 else b""


def extract_validations_and_comments(
    workbook_path: Any,
    sheets: list[tuple[int, str, str, str]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Extract all data-validation rules and cell comments for a workbook.

    Args:
        workbook_path: path to the .xlsx/.xlsm file
        sheets: WorkbookCatalog.sheets tuples (sheet_id, sheet_name, rel_id, target)

    Returns:
        (validation_records, comment_records) — each record carries sheet_id;
        IDs are assigned at persist time.
    """
    validations: list[dict[str, Any]] = []
    comments: list[dict[str, Any]] = []
    with zipfile.ZipFile(workbook_path, "r") as zipf:
        part_names = set(zipf.namelist())

        persons: dict[str, str] = {}
        for name in sorted(part_names):
            if name.startswith("xl/persons/") and name.endswith(".xml"):
                try:
                    persons.update(parse_persons(zipf.read(name)))
                except ET.ParseError:
                    continue

        for sheet_id, _sheet_name, _rel_id, target in sheets:
            part = _normalize_part_path(target)
            if part not in part_names:
                continue

            # b"dataValidation" also matches the x14 form inside extLst.
            if _zip_part_contains(zipf, part, b"dataValidation"):
                with zipf.open(part) as sheet_handle:
                    sheet_records = parse_sheet_validations(sheet_handle)
                for rec in sheet_records:
                    rec["sheet_id"] = sheet_id
                    validations.append(rec)

            rel_targets = _sheet_rel_targets(zipf, part)
            threaded_cells: set[str] = set()
            for tc_part in rel_targets["threaded"]:
                if tc_part not in part_names:
                    continue
                try:
                    tc_records = parse_threaded_comments(zipf.read(tc_part), persons)
                except ET.ParseError:
                    continue
                for rec in tc_records:
                    rec["sheet_id"] = sheet_id
                    threaded_cells.add(rec["a1"])
                    comments.append(rec)
            for c_part in rel_targets["comments"]:
                if c_part not in part_names:
                    continue
                try:
                    c_records = parse_classic_comments(zipf.read(c_part))
                except ET.ParseError:
                    continue
                for rec in c_records:
                    # Classic mirrors of threaded comments are Excel-generated
                    # placeholder boilerplate — the threaded part is canonical.
                    if rec["a1"] in threaded_cells:
                        continue
                    rec["sheet_id"] = sheet_id
                    comments.append(rec)
    return validations, comments


def store_validations_comments(
    conn: sqlite3.Connection,
    validations: list[dict[str, Any]],
    comments: list[dict[str, Any]],
) -> dict[str, int]:
    """Persist extracted records into data_validations / cell_comments."""
    conn.execute("DELETE FROM data_validations")
    conn.execute("DELETE FROM cell_comments")
    conn.executemany(
        """
        INSERT INTO data_validations (
            validation_id, sheet_id, sqref, val_type, operator,
            formula1, formula2, literal_values_json, allow_blank,
            prompt_title, prompt_text, source
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        [
            (
                idx + 1,
                rec["sheet_id"],
                rec["sqref"],
                rec["val_type"],
                rec["operator"],
                rec["formula1"],
                rec["formula2"],
                rec["literal_values_json"],
                rec["allow_blank"],
                rec["prompt_title"],
                rec["prompt_text"],
                rec["source"],
            )
            for idx, rec in enumerate(validations)
        ],
    )
    conn.executemany(
        """
        INSERT INTO cell_comments (
            comment_id, sheet_id, a1, author, text, kind, thread_order
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        [
            (
                idx + 1,
                rec["sheet_id"],
                rec["a1"],
                rec["author"],
                rec["text"],
                rec["kind"],
                rec["thread_order"],
            )
            for idx, rec in enumerate(comments)
        ],
    )
    conn.commit()
    return {"validations": len(validations), "comments": len(comments)}


def extract_and_store_validations_comments(
    workbook_path: Any,
    sheets: list[tuple[int, str, str, str]],
    conn: sqlite3.Connection,
) -> dict[str, int]:
    """One-call pipeline stage: extract from the workbook and persist."""
    validations, comments = extract_validations_and_comments(workbook_path, sheets)
    return store_validations_comments(conn, validations, comments)
