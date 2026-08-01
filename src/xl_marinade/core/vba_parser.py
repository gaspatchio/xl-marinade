# ABOUTME: VBA module parser that extracts UDF metadata without execution
# ABOUTME: Captures UDF names, parameters, volatility, and source hashes per IR Spec §10

import hashlib
import json
import logging
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    from oletools.olevba import VBA_Parser
except ImportError:  # pragma: no cover - optional dependency
    VBA_Parser = None
from openpyxl import Workbook

logger = logging.getLogger(__name__)


@dataclass
class UDFMetadata:
    """Metadata for a user-defined function extracted from VBA module"""

    name: str
    module: str
    param_count: int
    param_names: list[str]  # Empty if not extractable
    declared_volatile: bool
    source_text: str
    source_hash: str  # SHA-256 hex64

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for database storage"""
        return {
            "name": self.name,
            "module": self.module,
            "param_count": self.param_count,
            "param_names_json": json.dumps(self.param_names, sort_keys=True),
            "declared_volatile": self.declared_volatile,
            "source_text": self.source_text,
            "source_hash": self.source_hash,
        }


def compute_source_hash(source_text: str) -> str:
    """
    Compute deterministic SHA-256 hash of VBA source text.

    Args:
        source_text: VBA source code as string

    Returns:
        SHA-256 hash (lowercase hex, 64 characters)
    """
    hash_obj = hashlib.sha256(source_text.encode("utf-8"))
    return hash_obj.hexdigest()


def extract_function_signature(function_line: str) -> tuple[str, list[str]]:
    """
    Extract function name and parameter names from Function declaration line.

    Args:
        function_line: VBA function declaration (e.g., "Function MyFunc(a, b As Double)")

    Returns:
        (function_name, param_names) tuple

    Example:
        >>> extract_function_signature("Function MyFunc(a, b As Double)")
        ('MyFunc', ['a', 'b'])
        >>> extract_function_signature("Public Function Test()")
        ('Test', [])
    """
    # Match: [Public|Private] Function <name>([params])
    pattern = r"^\s*(?:Public\s+|Private\s+)?Function\s+(\w+)\s*\((.*?)\)"
    match = re.match(pattern, function_line, re.IGNORECASE)

    if not match:
        return "", []

    func_name = match.group(1)
    params_str = match.group(2).strip()

    if not params_str:
        return func_name, []

    # Parse parameter list: "a, b As Double, c" → ['a', 'b', 'c']
    # Split by comma, extract parameter name (before "As" or whitespace)
    param_names = []
    for param in params_str.split(","):
        param = param.strip()
        # Extract name before "As" keyword
        if " As " in param or " as " in param:
            param = re.split(r"\s+[Aa]s\s+", param)[0].strip()
        # Remove default value (e.g., "Optional x = 5" → "x")
        if "=" in param:
            param = param.split("=")[0].strip()
        # Remove Optional/ByVal/ByRef keywords
        param = re.sub(r"^\s*(Optional|ByVal|ByRef)\s+", "", param, flags=re.IGNORECASE).strip()
        if param:
            param_names.append(param)

    return func_name, param_names


def detect_application_volatile(function_body: str) -> bool:
    """
    Detect if function contains Application.Volatile call.

    Args:
        function_body: Complete function source text

    Returns:
        True if Application.Volatile is present (case-insensitive)
    """
    # Match: Application.Volatile (with optional True/False argument)
    pattern = r"\bApplication\.Volatile\b"
    return bool(re.search(pattern, function_body, re.IGNORECASE))


def extract_function_body(lines: list[str], start_idx: int) -> tuple[str, int]:
    """
    Extract complete function body starting from Function declaration.

    Args:
        lines: List of source lines
        start_idx: Index of "Function" declaration line

    Returns:
        (function_body, end_idx) - complete function text and index of "End Function"
    """
    function_lines = [lines[start_idx]]
    idx = start_idx + 1

    while idx < len(lines):
        line = lines[idx]
        function_lines.append(line)

        # Check for "End Function"
        if re.match(r"^\s*End\s+Function\s*$", line, re.IGNORECASE):
            return "\n".join(function_lines), idx

        idx += 1

    # Function not closed (malformed VBA)
    return "\n".join(function_lines), idx


def parse_vba_module(module_name: str, source_code: str) -> list[UDFMetadata]:
    """
    Parse VBA module source and extract all UDF metadata.

    Args:
        module_name: Name of the VBA module (e.g., "Module1")
        source_code: Complete VBA module source code

    Returns:
        List of UDFMetadata objects for all public functions
    """
    udfs = []
    lines = source_code.split("\n")

    idx = 0
    while idx < len(lines):
        line = lines[idx]

        # Match Function declaration (Public or no modifier means public)
        # Exclude Private functions
        if re.match(r"^\s*(?:Public\s+)?Function\s+\w+\s*\(", line, re.IGNORECASE):
            # Extract function signature
            func_name, param_names = extract_function_signature(line)

            if not func_name:
                idx += 1
                continue

            # Extract function body
            function_body, end_idx = extract_function_body(lines, idx)

            # Detect volatility
            declared_volatile = detect_application_volatile(function_body)

            # Compute source hash
            source_hash = compute_source_hash(function_body)

            # Create UDF metadata
            udf = UDFMetadata(
                name=func_name,
                module=module_name,
                param_count=len(param_names),
                param_names=param_names,
                declared_volatile=declared_volatile,
                source_text=function_body,
                source_hash=source_hash,
            )
            udfs.append(udf)

            # Skip to end of function
            idx = end_idx + 1
        else:
            idx += 1

    return udfs


def extract_udfs_from_workbook(workbook: Workbook) -> list[UDFMetadata]:
    """
    Extract all UDF metadata from workbook VBA project.

    Args:
        workbook: openpyxl Workbook object (must be loaded with keep_vba=True)

    Returns:
        List of UDFMetadata objects for all UDFs in all modules

    Implementation:
        Uses oletools.olevba to parse vbaProject.bin from the workbook's VBA archive.
        Extracts all public Function definitions from all modules.
    """
    if VBA_Parser is None:
        return []

    # Check if workbook has VBA (vba_archive attribute)
    if not hasattr(workbook, "vba_archive") or workbook.vba_archive is None:
        return []

    try:
        # Extract vbaProject.bin from the workbook's VBA archive
        # The vba_archive is a ZipFile-like object containing xl/vbaProject.bin
        vba_data = workbook.vba_archive.read("xl/vbaProject.bin")

        # Parse VBA using oletools
        # VBA_Parser can accept filename or data parameter
        vba_parser = VBA_Parser(filename="vbaProject.bin", data=vba_data)

        all_udfs: list[UDFMetadata] = []

        # Extract all VBA modules
        # vba_parser.extract_all_macros() returns (filename, stream_path, vba_filename, vba_code)
        for filename, stream_path, vba_filename, vba_code in vba_parser.extract_all_macros():
            if vba_code:
                # Parse this module to extract UDFs
                # Use stream_path or vba_filename as module name
                # Prefer vba_filename as it's more meaningful (e.g., "Module1", "ThisWorkbook")
                module_name = vba_filename if vba_filename else stream_path

                # Parse the VBA code to extract UDF metadata
                udfs = parse_vba_module(module_name, vba_code)
                all_udfs.extend(udfs)

        # Clean up
        vba_parser.close()

        return all_udfs

    except KeyError:
        # vbaProject.bin not found in archive
        return []
    except Exception:
        logger.exception("extract_udfs_from_workbook failed")
        return []


def extract_udfs_from_path(xlsm_path: Path | str) -> list[UDFMetadata]:
    """
    Extract UDFs directly from an .xlsm/.xlam/.xla file by reading
    ``xl/vbaProject.bin`` from the ZIP container. No openpyxl dependency.

    Used by the new-arch extraction pipeline which streams cell data via
    LazyValueFetcher and never materializes an openpyxl Workbook object.

    Args:
        xlsm_path: Path to a macro-enabled Excel file

    Returns:
        List of UDFMetadata objects extracted from the workbook's VBA project.
        Returns an empty list on: missing VBA project, corrupt ZIP, parse errors,
        or if the oletools dependency is unavailable.
    """
    if VBA_Parser is None:
        return []

    path = Path(xlsm_path)
    try:
        with zipfile.ZipFile(path) as zf:
            if "xl/vbaProject.bin" not in zf.namelist():
                # .xlsx or .xlsm-saved-as-xlsx with no VBA project — not an error.
                return []
            vba_data = zf.read("xl/vbaProject.bin")
    except (zipfile.BadZipFile, KeyError, OSError):
        logger.exception("Failed to read xl/vbaProject.bin from %s", path)
        return []

    try:
        vba_parser = VBA_Parser(filename="vbaProject.bin", data=vba_data)
    except Exception:
        logger.exception("VBA_Parser construction failed for %s", path)
        return []

    try:
        all_udfs: list[UDFMetadata] = []
        for _filename, _stream_path, vba_filename, vba_code in vba_parser.extract_all_macros():
            if vba_code:
                module_name = vba_filename or _stream_path or "UnknownModule"
                all_udfs.extend(parse_vba_module(module_name, vba_code))
        return all_udfs
    except Exception:
        logger.exception("extract_all_macros failed for %s", path)
        return []
    finally:
        try:
            vba_parser.close()
        except Exception:
            pass


def build_udf_map(udfs: list[UDFMetadata]) -> dict[str, UDFMetadata]:
    """
    Build mapping from UDF name (case-insensitive) to UDF metadata.

    Args:
        udfs: List of UDFMetadata objects

    Returns:
        Dictionary mapping uppercase UDF name to UDFMetadata
    """
    return {udf.name.upper(): udf for udf in udfs}


def detect_udf_calls_in_formula(
    formula: str, udf_map: dict[str, UDFMetadata]
) -> list[dict[str, str]]:
    """
    Detect UDF calls in formula and return udf_calls list.

    Args:
        formula: Formula string (e.g., "=MyFunc(A1, B1)")
        udf_map: Mapping from uppercase UDF name to UDFMetadata

    Returns:
        List of udf_call dicts: [{"name": "MyFunc", "module": "Module1"}, ...]
    """
    if not formula or not udf_map:
        return []

    udf_calls = []

    # Extract all function names from formula
    # Pattern: word followed by open paren (case-insensitive)
    pattern = r"\b([A-Za-z_]\w*)\s*\("
    matches = re.finditer(pattern, formula)

    for match in matches:
        func_name = match.group(1).upper()
        if func_name in udf_map:
            udf = udf_map[func_name]
            udf_calls.append(
                {
                    "name": udf.name,  # Original case from VBA
                    "module": udf.module,
                }
            )

    # Deduplicate while preserving order
    seen = set()
    unique_calls = []
    for call in udf_calls:
        key = (call["name"], call["module"])
        if key not in seen:
            seen.add(key)
            unique_calls.append(call)

    return unique_calls


def is_formula_volatile_due_to_udfs(
    udf_calls: list[dict[str, str]], udf_map: dict[str, UDFMetadata]
) -> bool:
    """
    Check if formula is volatile due to calling volatile UDFs.

    Args:
        udf_calls: List of udf_call dicts from detect_udf_calls_in_formula()
        udf_map: Mapping from uppercase UDF name to UDFMetadata

    Returns:
        True if any called UDF is declared volatile
    """
    for call in udf_calls:
        func_name_upper = call["name"].upper()
        if func_name_upper in udf_map and udf_map[func_name_upper].declared_volatile:
            return True
    return False
