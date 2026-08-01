# ABOUTME: Parses Excel xl/styles.xml to identify date-formatted cell styles
# ABOUTME: Classifies style indices as date vs non-date for format metadata extraction

"""
Excel Styles Parser — Date Format Detection

Parses xl/styles.xml from an xlsx/xlsm zip to classify which cell style
indices correspond to date formats.  Used by the fast extraction pipeline
to populate the format_blob_id column with {"is_date": true}.

Detection logic:
  1. Built-in numFmtIds 14-22, 27-36, 45-47 are always date formats.
  2. Custom numFmt codes containing d/m/y tokens (but not *only* h/m/s)
     are date formats.
  3. The 1904 date system flag is read from xl/workbook.xml.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

# Excel XML namespace
_SS_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _ns(tag: str) -> str:
    return f"{{{_SS_NS}}}{tag}"


# Built-in numFmtId values that Excel uses for date/datetime formats.
# Reference: ECMA-376, Part 1, §18.8.30 (numFmt).
_BUILTIN_DATE_FMT_IDS: frozenset[int] = frozenset(
    list(range(14, 23))  # 14-22: various date/time
    + list(range(27, 37))  # 27-36: CJK date formats
    + [45, 46, 47]  # 45-47: mm:ss, [h]:mm:ss, mm:ss.0
)

# Built-in numFmtId → format code (ECMA-376, Part 1, §18.8.30 "implied
# formats"). These codes are NOT present in the <numFmts> XML block —
# only custom ids (>=164) are serialized there — so the unit-bearing
# built-ins (9/10 = percent, 5-8/37-44 = accounting/paren-negative)
# must come from this table. Ids 23-36 are reserved/locale-dependent
# and omitted (the CJK date ids among them are already covered by
# _BUILTIN_DATE_FMT_IDS).
_BUILTIN_FORMAT_CODES: dict[int, str] = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    5: "$#,##0_);($#,##0)",
    6: "$#,##0_);[Red]($#,##0)",
    7: "$#,##0.00_);($#,##0.00)",
    8: "$#,##0.00_);[Red]($#,##0.00)",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "m/d/yyyy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yyyy h:mm",
    37: "#,##0_);(#,##0)",
    38: "#,##0_);[Red](#,##0)",
    39: "#,##0.00_);(#,##0.00)",
    40: "#,##0.00_);[Red](#,##0.00)",
    41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
    44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mm:ss.0",
    48: "##0.0E+0",
    49: "@",
}

# Regex to detect date tokens in a custom format code.
# We look for any of d, m, y (case-insensitive) that appear outside
# literal string sections (delimited by " or \).
# A format that ONLY has h/m/s tokens without d or y is a pure time
# format, not a date format — but we still classify it as date since
# Excel stores times as fractional serials too.
_DATE_TOKEN_RE = re.compile(r"[dDmMyY]")
_TIME_ONLY_RE = re.compile(r"^[\[\]hHmMsS0.:, ]+$")


@dataclass(frozen=True)
class DateFormatInfo:
    """Result of parsing date format metadata from a workbook."""

    date_style_indices: frozenset[int] = field(default_factory=frozenset)
    is_1904: bool = False
    # cellXf index → raw number-format signal flags (subset of
    # {"percent": True, "sign_parens": True, "unit": "<word>"}).
    # Only indices with at least one flag appear; interpretation
    # (e.g. percent |value|<=1 guard) is the consumer's job.
    format_flags_by_style: dict[int, dict] = field(default_factory=dict)


def _strip_literals(fmt_code: str) -> str:
    """Remove quoted literal sections and escaped characters from a format code."""
    # Remove sections in double quotes
    result = re.sub(r'"[^"]*"', "", fmt_code)
    # Remove backslash-escaped characters
    result = re.sub(r"\\.", "", result)
    return result


# Quoted text literal that plausibly names a unit: letters (optionally
# space-separated words), 2-20 chars. Rejects "%", "-", "₿", punctuation.
_UNIT_LITERAL_RE = re.compile(r"^[A-Za-z][A-Za-z ]{0,18}[A-Za-z]$")


def derive_format_flags(fmt_code: str | None) -> dict:
    """Derive raw signal flags from a number-format code.

    Returns a subset of {"percent": True, "sign_parens": True,
    "unit": "<word>"}. Raw signal only — no interpretation (the
    percent |value|<=1 guard lives in the consumer). Codes containing
    the anonymization placeholder "₿" (seen in one client model) yield NO
    flags; other currency symbols are simply never emitted as a field.
    """
    if not fmt_code or fmt_code == "General":
        return {}
    if "₿" in fmt_code:
        return {}
    flags: dict = {}
    stripped = _strip_literals(fmt_code)
    # Percent-scaling token (a literal "%" outside quotes multiplies by 100).
    if "%" in stripped:
        flags["percent"] = True
    # Accounting/paren-negative convention: the negative (second ;-separated)
    # section wraps the number in parentheses.
    sections = stripped.split(";")
    if len(sections) >= 2:
        # Drop color tags ([Red]) and _x padding chars before checking parens.
        neg = re.sub(r"\[[^\]]*\]", "", sections[1])
        neg = re.sub(r"_.", "", neg)
        if "(" in neg and ")" in neg:
            flags["sign_parens"] = True
    # Labelled unit literal, e.g. 0" months" — skipped for date formats,
    # whose literals are connectives ("of", "at"), not units. "BTC" is
    # the ₿ anonymization placeholder spelled out (observed as 0.0" BTC"
    # in one model) — same guard as ₿: never emit it.
    if not _is_date_format_code(fmt_code):
        for literal in re.findall(r'"([^"]*)"', fmt_code):
            word = literal.strip()
            if _UNIT_LITERAL_RE.match(word) and word.upper() != "BTC":
                flags["unit"] = word
                break
    return flags


def _is_date_format_code(fmt_code: str) -> bool:
    """Check whether a custom numFmt code represents a date/datetime format."""
    if not fmt_code:
        return False
    stripped = _strip_literals(fmt_code)
    # Must contain at least one date token (d, m, y)
    # Pure time formats (only h:m:s) are also date-like since Excel stores
    # them as serial fractions, but for our purposes we want to detect
    # cells that hold date serials, so we require d or y tokens.
    has_date_token = bool(re.search(r"[dDyY]", stripped))
    return has_date_token


def _parse_is_1904(zipf: zipfile.ZipFile) -> bool:
    """Check xl/workbook.xml for the date1904 attribute."""
    try:
        with zipf.open("xl/workbook.xml") as f:
            for event, elem in ET.iterparse(f, events=("end",)):
                if elem.tag == _ns("workbookPr"):
                    val = (elem.get("date1904") or "").lower()
                    return val in ("1", "true")
                elem.clear()
    except (KeyError, ET.ParseError):
        pass
    return False


def parse_date_format_info(workbook_path: str | Path) -> DateFormatInfo:
    """
    Parse an xlsx/xlsm workbook to identify which cell style indices are
    date-formatted.

    Args:
        workbook_path: Path to the .xlsx/.xlsm file.

    Returns:
        DateFormatInfo with the set of style indices that are date formats
        and whether the workbook uses the 1904 date system.
    """
    workbook_path = Path(workbook_path)
    try:
        with zipfile.ZipFile(workbook_path, "r") as zipf:
            is_1904 = _parse_is_1904(zipf)
            date_indices, flags_by_style = _parse_styles_xml(zipf)
            return DateFormatInfo(
                date_style_indices=frozenset(date_indices),
                is_1904=is_1904,
                format_flags_by_style=flags_by_style,
            )
    except (zipfile.BadZipFile, FileNotFoundError, OSError):
        return DateFormatInfo()


def _parse_styles_xml(zipf: zipfile.ZipFile) -> tuple[set[int], dict[int, dict]]:
    """Parse xl/styles.xml: (date-format cellXf indices, cellXf index → signal flags)."""
    try:
        raw = zipf.read("xl/styles.xml")
    except KeyError:
        return set(), {}

    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return set(), {}

    # Step 1: Build numFmtId → formatCode mapping from <numFmts>.
    custom_fmt_codes: dict[int, str] = {}
    num_fmts_elem = root.find(_ns("numFmts"))
    if num_fmts_elem is not None:
        for nf in num_fmts_elem.findall(_ns("numFmt")):
            fmt_id_str = nf.get("numFmtId")
            fmt_code = nf.get("formatCode")
            if fmt_id_str is not None and fmt_code is not None:
                try:
                    custom_fmt_codes[int(fmt_id_str)] = fmt_code
                except ValueError:
                    pass

    # Step 2: Walk <cellXfs> to classify each xf index.
    date_indices: set[int] = set()
    flags_by_style: dict[int, dict] = {}
    flags_by_fmt_id: dict[int, dict] = {}  # memoize per numFmtId
    cell_xfs_elem = root.find(_ns("cellXfs"))
    if cell_xfs_elem is None:
        return set(), {}

    for idx, xf in enumerate(cell_xfs_elem.findall(_ns("xf"))):
        fmt_id_str = xf.get("numFmtId")
        if fmt_id_str is None:
            continue
        try:
            fmt_id = int(fmt_id_str)
        except ValueError:
            continue

        if fmt_id in _BUILTIN_DATE_FMT_IDS:
            date_indices.add(idx)
        elif fmt_id in custom_fmt_codes:
            if _is_date_format_code(custom_fmt_codes[fmt_id]):
                date_indices.add(idx)

        if fmt_id not in flags_by_fmt_id:
            # Custom <numFmts> entry wins; built-ins come from the table.
            fmt_code = custom_fmt_codes.get(fmt_id) or _BUILTIN_FORMAT_CODES.get(fmt_id)
            flags_by_fmt_id[fmt_id] = derive_format_flags(fmt_code)
        if flags_by_fmt_id[fmt_id]:
            flags_by_style[idx] = flags_by_fmt_id[fmt_id]

    return date_indices, flags_by_style
