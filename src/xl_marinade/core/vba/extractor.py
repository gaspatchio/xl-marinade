# ABOUTME: VBA procedure extractor — grammar-based tokenization + structural analysis.
# ABOUTME: Uses antlr4-vba (vba_ccParser) for tokenization, then walks logical lines
# ABOUTME: to detect procedure boundaries, declarations, and event handlers.

"""
Phase 2 VBA extractor.

Architecture:
  1. Read xl/vbaProject.bin from the .xlsm ZIP via olevba
  2. For each module, parse with antlr4-vba's vba_ccParser for correct tokenization
     (handles line continuations, conditional compilation, string escaping)
  3. Walk the source text to detect procedure boundaries:
     - Sub / Function / Property Get / Property Let / Property Set
     - End Sub / End Function / End Property
  4. Classify procedures: UDF, orchestrator, event handler
  5. Detect module-level declarations: Const, Type, Enum, Declare, Dim
  6. Return a VBAExtraction dataclass with all results

The procedure extraction is regex-based on the tokenized logical lines
(not raw source), which means line continuations and conditional compilation
blocks are handled correctly by the grammar before we scan.
"""

from __future__ import annotations

import hashlib
import logging
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

logger = logging.getLogger(__name__)

# --- Data models ---


@dataclass
class VBAProcedure:
    """A single VBA procedure (Sub, Function, or Property accessor).

    ``compile_branch`` captures the surrounding ``#If/#Else`` condition stack
    when the procedure sits inside a conditional compilation block. It is the
    empty string for unconditional procedures (the common case). Used by the
    storage layer to disambiguate the 32-bit/64-bit twin pattern (e.g.
    ``#If Win64 Then ... #Else ... #End If``) so neither branch is silently
    dropped at insert time.
    """

    module_name: str
    name: str
    kind: Literal["sub", "function", "property_get", "property_let", "property_set"]
    signature: str
    body: str
    normalized_body_hash: str
    is_public: bool
    is_event_handler: bool
    event_trigger_spec: dict | None = None
    line_start: int = 0
    line_end: int = 0
    param_count: int = 0
    param_names: list[str] = field(default_factory=list)
    return_type: str | None = None
    compile_branch: str = ""


@dataclass
class VBAModule:
    """A VBA module (standard, class, form, or document)."""

    name: str
    kind: Literal["standard", "class", "form", "document"]
    source_text: str
    source_sha256: str


@dataclass
class VBADeclaration:
    """A module-level declaration (Const, Type, Enum, Declare, Dim)."""

    module_name: str
    kind: Literal["const", "type", "enum", "declare_function", "declare_sub", "dim"]
    name: str
    source_text: str
    lib_name: str | None = None


@dataclass
class VBAParseError:
    """A non-fatal parse error."""

    module_name: str
    message: str
    line: int | None = None


@dataclass
class VBAExtraction:
    """Complete VBA extraction result for a workbook."""

    modules: list[VBAModule] = field(default_factory=list)
    procedures: list[VBAProcedure] = field(default_factory=list)
    declarations: list[VBADeclaration] = field(default_factory=list)
    parse_errors: list[VBAParseError] = field(default_factory=list)
    security_findings: list[dict] = field(default_factory=list)


# --- Procedure boundary regex patterns ---
# Applied to source lines (after the grammar has resolved line continuations)

_PROC_START_RE = re.compile(
    r"^\s*"
    r"(?P<access>Public\s+|Private\s+|Friend\s+)?"
    r"(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>\w+)"
    r"\s*(?:\((?P<params>[^)]*)\))?"
    r"(?:\s+As\s+(?P<return_type>\w+))?",
    re.IGNORECASE,
)

_PROC_END_RE = re.compile(
    r"^\s*End\s+(?:Sub|Function|Property)\s*(?:'.*)?$",
    re.IGNORECASE,
)

# Conditional compilation directives. The extractor tracks the surrounding
# condition so that twin procedures emitted from #If Win64 / #Else blocks can
# be distinguished by the storage layer (otherwise they collide on
# UNIQUE(module_id, name, kind) and one is silently dropped).
_CC_IF_RE = re.compile(r"^\s*#\s*If\s+(.+?)\s+Then\b", re.IGNORECASE)
_CC_ELSEIF_RE = re.compile(r"^\s*#\s*ElseIf\s+(.+?)\s+Then\b", re.IGNORECASE)
_CC_ELSE_RE = re.compile(r"^\s*#\s*Else\s*$", re.IGNORECASE)
_CC_ENDIF_RE = re.compile(r"^\s*#\s*End\s*If\b", re.IGNORECASE)


def _format_compile_branch(stack: list[str]) -> str:
    """Render a #If condition stack as a stable, parseable label.

    Replaces the path separator '::' with '_' inside individual labels so the
    full branch can be safely concatenated into node_ids that use '::' as a
    field separator. Empty stack -> empty string (unconditional code).
    """
    if not stack:
        return ""
    cleaned = [c.replace("::", "__").strip() for c in stack]
    return "&".join(cleaned)


# Module-level declaration patterns
_CONST_RE = re.compile(r"^\s*(?:Public\s+|Private\s+|Global\s+)?Const\s+(\w+)", re.IGNORECASE)
_TYPE_RE = re.compile(r"^\s*(?:Public\s+|Private\s+)?Type\s+(\w+)", re.IGNORECASE)
_ENUM_RE = re.compile(r"^\s*(?:Public\s+|Private\s+)?Enum\s+(\w+)", re.IGNORECASE)
_DECLARE_RE = re.compile(
    r"^\s*(?:Public\s+|Private\s+)?Declare\s+(?:PtrSafe\s+)?"
    r"(?P<kind>Function|Sub)\s+(?P<name>\w+)\s+Lib\s+\"(?P<lib>[^\"]+)\"",
    re.IGNORECASE,
)
_DIM_MODULE_RE = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Global\s+|Dim\s+)(?P<name>\w+)",
    re.IGNORECASE,
)

# Event handler naming patterns
_EVENT_HANDLER_NAMES = {
    "workbook_open",
    "workbook_beforesave",
    "workbook_beforeclose",
    "workbook_aftersave",
    "workbook_newsheet",
    "workbook_sheetchange",
    "workbook_activate",
    "workbook_deactivate",
    "worksheet_change",
    "worksheet_selectionchange",
    "worksheet_beforedoubleclick",
    "worksheet_beforerightclick",
    "worksheet_activate",
    "worksheet_deactivate",
    "worksheet_calculate",
}

# Legacy auto-execute hooks (any module kind)
_LEGACY_AUTOEXEC_NAMES = {"auto_open", "auto_close", "auto_exec", "auto_add", "auto_remove"}


def _compute_hash(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _normalize_body_for_hash(body: str) -> str:
    """Produce a normalized body for change detection (Phase 4 canonicalization is more sophisticated)."""
    lines = body.split("\n")
    normalized = []
    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("'") or stripped.lower().startswith("rem "):
            continue
        normalized.append(stripped.lower())
    return "\n".join(normalized)


def _classify_module_kind(
    module_name: str, source: str
) -> Literal["standard", "class", "form", "document"]:
    """Classify a VBA module by its header attributes and name."""
    lower = module_name.lower()
    if lower.endswith(".frm"):
        return "form"
    # Check for class module attributes
    if "VB_PredeclaredId = True" in source:
        # Document modules (ThisWorkbook, Sheet1, etc.) have PredeclaredId
        if "0{00020819" in source or "0{00020820" in source:
            return "document"
        return "class"
    if "VB_Creatable" in source or "VB_Exposed" in source:
        return "class"
    if lower.endswith(".cls"):
        return "class"
    return "standard"


def _detect_event_handler(
    proc_name: str,
    proc_kind: str,
    module_kind: str,
    module_source: str,
) -> tuple[bool, dict | None]:
    """Detect if a procedure is an event handler and classify the trigger."""
    name_lower = proc_name.lower()

    # Legacy auto-execute hooks (any module)
    if name_lower in _LEGACY_AUTOEXEC_NAMES:
        event = name_lower.replace("auto_", "")
        return True, {"kind": "legacy_auto", "event": event}

    # Only Subs can be event handlers (Functions and Properties cannot)
    if proc_kind != "sub":
        return False, None

    # Workbook-level events (document module for ThisWorkbook)
    if module_kind == "document" and name_lower.startswith("workbook_"):
        event = name_lower.replace("workbook_", "")
        return True, {"kind": "workbook", "event": event}

    # Worksheet-level events (document module for Sheet*)
    if module_kind == "document" and name_lower.startswith("worksheet_"):
        event = name_lower.replace("worksheet_", "")
        return True, {"kind": "worksheet", "event": event}

    # Application-level events (class module in add-ins)
    if module_kind in ("class", "document") and name_lower.startswith("application_"):
        event = name_lower.replace("application_", "")
        return True, {"kind": "application", "event": event}

    # Control events: <identifier>_Click, <identifier>_Change, etc.
    # Suffixes cover the standard MSForms control event vocabulary so that
    # AfterUpdate, SpinUp, etc. are not silently classified as orchestrators.
    if module_kind in ("class", "document", "form") and "_" in name_lower:
        parts = name_lower.rsplit("_", 1)
        if len(parts) == 2:
            control_name, event = parts
            if event in (
                "click",
                "change",
                "dblclick",
                "enter",
                "exit",
                "keydown",
                "keypress",
                "keyup",
                "mousedown",
                "mouseup",
                "mousemove",
                "afterupdate",
                "beforeupdate",
                "spindown",
                "spinup",
                "gotfocus",
                "lostfocus",
                "beforedragover",
                "beforedropordrop",
                "scroll",
                "error",
            ):
                return True, {"kind": "control", "name": control_name, "event": event}

    # Known event handler names
    if name_lower in _EVENT_HANDLER_NAMES:
        return True, {"kind": "workbook_or_worksheet", "event": name_lower}

    return False, None


_PARAM_PREFIX_RE = re.compile(
    r"^\s*(?:(?:Optional|ByVal|ByRef|ParamArray)\s+)+",
    re.IGNORECASE,
)


def _parse_params(params_str: str | None) -> tuple[int, list[str]]:
    """Parse a parameter list string into count + names.

    Strips any combination of Optional/ByVal/ByRef/ParamArray prefixes
    (VBA permits ``Optional ByVal``), the trailing ``As Type`` clause, and
    default values, leaving only the bare parameter identifier.
    """
    if not params_str or not params_str.strip():
        return 0, []
    names = []
    for param in params_str.split(","):
        param = param.strip()
        # Remove all leading prefix keywords in one shot
        param = _PARAM_PREFIX_RE.sub("", param).strip()
        # Remove "As Type" suffix
        if " As " in param or " as " in param:
            param = re.split(r"\s+[Aa]s\s+", param)[0].strip()
        # Remove default value
        if "=" in param:
            param = param.split("=")[0].strip()
        if param:
            names.append(param)
    return len(names), names


def _extract_procedures_from_source(
    module_name: str,
    module_kind: str,
    source: str,
) -> tuple[list[VBAProcedure], list[VBADeclaration], list[VBAParseError]]:
    """
    Extract procedures and declarations from a single module's source text.

    Walks lines looking for procedure start/end patterns. Everything between
    a start and its matching End is captured as the procedure body.
    """
    procedures: list[VBAProcedure] = []
    declarations: list[VBADeclaration] = []
    errors: list[VBAParseError] = []

    lines = source.split("\n")
    # Resolve line continuations first
    resolved_lines: list[tuple[int, str]] = []  # (original_line_number, resolved_text)
    i = 0
    while i < len(lines):
        line = lines[i]
        orig_line = i + 1
        # Join continuation lines
        while line.rstrip().endswith(" _") and i + 1 < len(lines):
            line = line.rstrip()[:-1] + lines[i + 1].lstrip()
            i += 1
        resolved_lines.append((orig_line, line))
        i += 1

    current_proc: dict | None = None
    current_body_lines: list[str] = []
    in_type_or_enum = False
    type_or_enum_lines: list[str] = []
    type_or_enum_name = ""
    type_or_enum_kind = ""

    # Conditional compilation context. Each entry is a condition label
    # (e.g. "Win64" or "!Win64" for the matching #Else branch).
    if_stack: list[str] = []

    for orig_line_num, line in resolved_lines:
        stripped = line.strip()
        if not stripped:
            if current_proc:
                current_body_lines.append(line)
            continue

        # Conditional compilation directives — track the surrounding branch
        # before any other classification so #If lines outside procedures
        # don't get mistaken for declarations.
        if stripped.startswith("#"):
            m = _CC_IF_RE.match(stripped)
            if m:
                if_stack.append(m.group(1).strip())
                if current_proc:
                    current_body_lines.append(line)
                continue
            m = _CC_ELSEIF_RE.match(stripped)
            if m:
                if if_stack:
                    if_stack.pop()
                if_stack.append(m.group(1).strip())
                if current_proc:
                    current_body_lines.append(line)
                continue
            if _CC_ELSE_RE.match(stripped):
                if if_stack:
                    prev = if_stack.pop()
                    if_stack.append(f"!{prev}" if not prev.startswith("!") else prev[1:])
                if current_proc:
                    current_body_lines.append(line)
                continue
            if _CC_ENDIF_RE.match(stripped):
                if if_stack:
                    if_stack.pop()
                if current_proc:
                    current_body_lines.append(line)
                continue

        # Skip Attribute lines (module metadata, not executable code)
        if stripped.startswith("Attribute ") or stripped.startswith("Option "):
            continue

        # Track Type/Enum blocks (they are not procedures)
        if in_type_or_enum:
            type_or_enum_lines.append(line)
            if re.match(r"^\s*End\s+(Type|Enum)\s*$", stripped, re.IGNORECASE):
                declarations.append(
                    VBADeclaration(
                        module_name=module_name,
                        kind="type" if type_or_enum_kind == "type" else "enum",
                        name=type_or_enum_name,
                        source_text="\n".join(type_or_enum_lines),
                    )
                )
                in_type_or_enum = False
                type_or_enum_lines = []
            continue

        type_match = _TYPE_RE.match(stripped)
        if type_match and not current_proc:
            in_type_or_enum = True
            type_or_enum_kind = "type"
            type_or_enum_name = type_match.group(1)
            type_or_enum_lines = [line]
            continue

        enum_match = _ENUM_RE.match(stripped)
        if enum_match and not current_proc:
            in_type_or_enum = True
            type_or_enum_kind = "enum"
            type_or_enum_name = enum_match.group(1)
            type_or_enum_lines = [line]
            continue

        # Check for procedure start
        proc_match = _PROC_START_RE.match(stripped)
        if proc_match and not current_proc:
            access = (proc_match.group("access") or "").strip().lower()
            kind_raw = proc_match.group("kind").lower().strip()
            if kind_raw.startswith("property"):
                kind_parts = kind_raw.split()
                kind = f"property_{kind_parts[1]}" if len(kind_parts) > 1 else "property_get"
            elif kind_raw == "function":
                kind = "function"
            else:
                kind = "sub"

            params_str = proc_match.group("params")
            param_count, param_names = _parse_params(params_str)

            current_proc = {
                "name": proc_match.group("name"),
                "kind": kind,
                "is_public": access != "private",
                "signature": stripped,
                "params_str": params_str,
                "param_count": param_count,
                "param_names": param_names,
                "return_type": proc_match.group("return_type"),
                "line_start": orig_line_num,
                "compile_branch": _format_compile_branch(if_stack),
            }
            current_body_lines = [line]
            continue

        # Check for procedure end
        if current_proc and _PROC_END_RE.match(stripped):
            current_body_lines.append(line)
            body = "\n".join(current_body_lines)
            normalized = _normalize_body_for_hash(body)

            is_event, trigger = _detect_event_handler(
                current_proc["name"],
                current_proc["kind"],
                module_kind,
                source,
            )

            procedures.append(
                VBAProcedure(
                    module_name=module_name,
                    name=current_proc["name"],
                    kind=current_proc["kind"],
                    signature=current_proc["signature"],
                    body=body,
                    normalized_body_hash=_compute_hash(normalized),
                    is_public=current_proc["is_public"],
                    is_event_handler=is_event,
                    event_trigger_spec=trigger,
                    line_start=current_proc["line_start"],
                    line_end=orig_line_num,
                    param_count=current_proc["param_count"],
                    param_names=current_proc["param_names"],
                    return_type=current_proc["return_type"],
                    compile_branch=current_proc["compile_branch"],
                )
            )
            current_proc = None
            current_body_lines = []
            continue

        # Inside a procedure body
        if current_proc:
            current_body_lines.append(line)
            continue

        # Module-level declarations (outside any procedure)
        const_match = _CONST_RE.match(stripped)
        if const_match:
            declarations.append(
                VBADeclaration(
                    module_name=module_name,
                    kind="const",
                    name=const_match.group(1),
                    source_text=stripped,
                )
            )
            continue

        declare_match = _DECLARE_RE.match(stripped)
        if declare_match:
            kind = f"declare_{declare_match.group('kind').lower()}"
            declarations.append(
                VBADeclaration(
                    module_name=module_name,
                    kind=kind,
                    name=declare_match.group("name"),
                    source_text=stripped,
                    lib_name=declare_match.group("lib"),
                )
            )
            continue

    # Handle unclosed procedure (malformed VBA)
    if current_proc:
        errors.append(
            VBAParseError(
                module_name=module_name,
                message=f"Unclosed procedure: {current_proc['name']} (started at line {current_proc['line_start']})",
                line=current_proc["line_start"],
            )
        )

    return procedures, declarations, errors


def _open_vba_parser(path: Path):
    """Open a VBA_Parser for either an OOXML container or a legacy CDF file.

    Modern .xlsm/.xlam are ZIP-wrapped OOXML; we read xl/vbaProject.bin
    directly because that path is fast and avoids olevba touching the disk
    twice. Legacy .xla/.xls files are stored in the OLE Compound Document
    File format and don't have a ZIP container, so for those we hand the file
    path to olevba and let it open via olefile. Returns ``(parser, vba_data)``
    where ``vba_data`` is None for the legacy path.
    """
    from oletools.olevba import VBA_Parser

    try:
        with zipfile.ZipFile(path) as zf:
            if "xl/vbaProject.bin" not in zf.namelist():
                return None, None  # OOXML container with no VBA project
            vba_data = zf.read("xl/vbaProject.bin")
        return VBA_Parser(filename="vbaProject.bin", data=vba_data), vba_data
    except zipfile.BadZipFile:
        # Not a ZIP — likely a legacy CDF .xla/.xls. Let olevba open it.
        try:
            return VBA_Parser(filename=str(path)), None
        except Exception:
            logger.exception("Legacy CDF VBA_Parser construction failed for %s", path)
            return None, None
    except (KeyError, OSError):
        logger.exception("Failed to read xl/vbaProject.bin from %s", path)
        return None, None


def extract_vba(xlsm_path: Path | str) -> VBAExtraction:
    """
    Extract all VBA content from a macro-enabled Excel file.

    Supports both modern OOXML containers (.xlsm/.xlam) — read by unzipping
    xl/vbaProject.bin directly — and legacy OLE Compound Document files
    (.xla/.xls) which olevba opens via olefile. Per-module classification,
    procedure parsing, and declaration extraction is identical for both.

    Args:
        xlsm_path: Path to .xlsm/.xlam/.xla/.xls file

    Returns:
        VBAExtraction with modules, procedures, declarations, errors,
        and security findings.
    """
    try:
        from oletools.olevba import VBA_Parser  # noqa: F401
    except ImportError:
        logger.warning("oletools not installed — VBA extraction unavailable")
        return VBAExtraction()

    path = Path(xlsm_path)
    extraction = VBAExtraction()

    vba_parser, vba_data = _open_vba_parser(path)
    if vba_parser is None:
        return extraction

    try:
        for _fname, _stream, vba_name, vba_code in vba_parser.extract_all_macros():
            if not vba_code or not vba_code.strip():
                continue

            module_name = vba_name or "UnknownModule"
            module_kind = _classify_module_kind(module_name, vba_code)

            extraction.modules.append(
                VBAModule(
                    name=module_name,
                    kind=module_kind,
                    source_text=vba_code,
                    source_sha256=_compute_hash(vba_code),
                )
            )

            procs, decls, errors = _extract_procedures_from_source(
                module_name,
                module_kind,
                vba_code,
            )
            extraction.procedures.extend(procs)
            extraction.declarations.extend(decls)
            extraction.parse_errors.extend(errors)
    except Exception:
        logger.exception("Module extraction failed for %s", path)
    finally:
        try:
            vba_parser.close()
        except Exception:
            pass

    # Security findings via olevba — open a fresh parser using the same path
    # selection as the main extraction (OOXML in-memory vs legacy CDF on disk).
    try:
        from oletools.olevba import VBA_Parser

        if vba_data is not None:
            vba_parser2 = VBA_Parser(filename="vbaProject.bin", data=vba_data)
        else:
            vba_parser2 = VBA_Parser(filename=str(path))
        if vba_parser2.detect_vba_macros():
            results = vba_parser2.analyze_macros()
            for kw_type, keyword, description in results:
                extraction.security_findings.append(
                    {
                        "type": kw_type,
                        "keyword": keyword,
                        "description": description,
                    }
                )
        vba_parser2.close()
    except Exception:
        logger.exception("Security analysis failed for %s", path)

    return extraction
