# ABOUTME: Reads VBA module source from a vbaProject.bin (MS-OVBA) using only olefile.
# ABOUTME: Replaces oletools for extraction: decompression, dir stream, PROJECT stream, keyword scan.

"""
VBA project reader, implemented from Microsoft's published [MS-OVBA] spec.

A VBA project is an OLE compound file (``xl/vbaProject.bin`` inside an
``.xlsm``, or a storage inside a legacy ``.xls``/``.xla``). This module reads
it with `olefile` (BSD-2-Clause, pure Python) and nothing else:

- ``VBA/dir`` (compressed) lists the modules: name, stream name, and where
  the module's compressed source starts in its stream (§2.3.4.2);
- each module stream holds that source, compressed (§2.4.1);
- ``PROJECT`` (plain text) says which modules are standard (``.bas``),
  class/document (``.cls``) or form (``.frm``) modules (§2.3.1).

It previously came from oletools, whose required dependency pcodedmp is
GPLv3; xl-marinade is Apache-2.0 and keeps GPL out of its install.
Module names, file extensions, ordering and decoded source match what
oletools' ``VBA_Parser.extract_all_macros`` returned for the project's own
modules. oletools additionally scanned streams outside the project for
orphaned code and disassembled P-code; neither was ever stored.
"""

from __future__ import annotations

import codecs
import re
import struct
from dataclasses import dataclass
from pathlib import Path

import olefile


class VBAProjectError(Exception):
    """The data is not a readable VBA project (not OLE, corrupt, truncated)."""


@dataclass(slots=True, frozen=True)
class VBAModuleSource:
    """One module of a VBA project, with its decompressed source."""

    name: str  # module name, e.g. "Sheet1"
    filename: str  # name plus extension, e.g. "Sheet1.cls" (oletools' vba_filename)
    stream_path: str  # OLE path of the module stream, e.g. "VBA/Sheet1"
    code: str


# ---------------------------------------------------------------------------
# Decompression (MS-OVBA 2.4.1)
# ---------------------------------------------------------------------------

_CHUNK_RAW_SIZE = 4096


def decompress(data: bytes) -> bytes:
    """Decompress an MS-OVBA CompressedContainer.

    The container is a signature byte (0x01) followed by chunks. Each chunk
    has a 2-byte header: bits 0-11 are the chunk size minus 3, bits 12-14
    the signature 0b011, and bit 15 whether the chunk is compressed. An
    uncompressed chunk is 4096 literal bytes. A compressed chunk is a run of
    flag bytes, each followed by 8 tokens: a literal byte (flag bit 0) or a
    2-byte copy token (flag bit 1) pointing back into the chunk's output.

    Raises:
        VBAProjectError: on a bad signature or a token pointing outside the
            chunk.
    """
    if not data or data[0] != 0x01:
        raise VBAProjectError("compressed container does not start with 0x01")
    out = bytearray()
    pos = 1
    end = len(data)
    while pos < end:
        if pos + 2 > end:
            raise VBAProjectError("truncated chunk header")
        (header,) = struct.unpack_from("<H", data, pos)
        if (header >> 12) & 0x07 != 0b011:
            raise VBAProjectError(f"bad chunk signature in header 0x{header:04X}")
        chunk_end = min(pos + (header & 0x0FFF) + 3, end)
        pos += 2
        chunk_start = len(out)
        if not header & 0x8000:
            out += data[pos : pos + _CHUNK_RAW_SIZE]
            pos += _CHUNK_RAW_SIZE
            continue
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if not flags & (1 << bit):
                    out.append(data[pos])
                    pos += 1
                    continue
                if pos + 2 > chunk_end:
                    raise VBAProjectError("truncated copy token")
                (token,) = struct.unpack_from("<H", data, pos)
                pos += 2
                # The split between offset and length bits grows with the
                # amount of output this chunk has produced (§2.4.1.3.19.1).
                decompressed = len(out) - chunk_start
                offset_bits = max((decompressed - 1).bit_length(), 4)
                length = (token & (0xFFFF >> offset_bits)) + 3
                source = len(out) - ((token >> (16 - offset_bits)) + 1)
                if source < chunk_start:
                    raise VBAProjectError("copy token points before the chunk")
                for i in range(length):  # byte by byte: the ranges may overlap
                    out.append(out[source + i])
    return bytes(out)


# ---------------------------------------------------------------------------
# dir stream (MS-OVBA 2.3.4.2)
# ---------------------------------------------------------------------------

_PROJECTVERSION = 0x0009
_PROJECTCODEPAGE = 0x0003
_MODULENAME = 0x0019
_MODULENAMEUNICODE = 0x0047
_MODULESTREAMNAME = 0x001A
_MODULESTREAMNAMEUNICODE = 0x0032
_MODULEOFFSET = 0x0031
_MODULE_TERMINATOR = 0x002B


@dataclass(slots=True)
class _ModuleRecord:
    name: bytes = b""
    name_unicode: str | None = None
    stream_name: bytes | None = None
    stream_name_unicode: str | None = None
    text_offset: int = 0


def _records(dir_bytes: bytes) -> list[tuple[int, bytes]]:
    """Split the decompressed dir stream into (id, data) records.

    Every record is an id (u16), a size (u32) and that many bytes -- except
    PROJECTVERSION, whose size field always reads 4 but which carries 6
    bytes (major u32, minor u16).
    """
    records: list[tuple[int, bytes]] = []
    pos = 0
    end = len(dir_bytes)
    while pos + 6 <= end:
        record_id, size = struct.unpack_from("<HI", dir_bytes, pos)
        pos += 6
        if record_id == _PROJECTVERSION:
            size = 6
        if pos + size > end:
            raise VBAProjectError(f"dir record 0x{record_id:04X} overruns the stream")
        records.append((record_id, dir_bytes[pos : pos + size]))
        pos += size
    return records


def _parse_dir(dir_bytes: bytes) -> tuple[int, list[_ModuleRecord]]:
    """The project code page and its module records, in dir order."""
    codepage = 1252
    modules: list[_ModuleRecord] = []
    current: _ModuleRecord | None = None
    for record_id, data in _records(dir_bytes):
        if record_id == _PROJECTCODEPAGE and len(data) >= 2:
            (codepage,) = struct.unpack_from("<H", data)
        elif record_id == _MODULENAME:
            current = _ModuleRecord(name=data)
            modules.append(current)
        elif current is None:
            continue
        elif record_id == _MODULENAMEUNICODE:
            current.name_unicode = data.decode("utf-16-le", "replace")
        elif record_id == _MODULESTREAMNAME:
            current.stream_name = data
        elif record_id == _MODULESTREAMNAMEUNICODE:
            current.stream_name_unicode = data.decode("utf-16-le", "replace")
        elif record_id == _MODULEOFFSET and len(data) >= 4:
            (current.text_offset,) = struct.unpack_from("<I", data)
        elif record_id == _MODULE_TERMINATOR:
            current = None
    return codepage, modules


# Code pages whose Python codec is not simply "cp<number>".
_CODEPAGE_CODECS = {
    1200: "utf_16_le",
    1201: "utf_16_be",
    10000: "mac_roman",
    10006: "mac_greek",
    10007: "mac_cyrillic",
    10029: "mac_latin2",
    10079: "mac_iceland",
    10081: "mac_turkish",
    20127: "ascii",
    65000: "utf_7",
    65001: "utf_8",
    **{28590 + n: f"iso8859_{n}" for n in range(1, 16) if n != 12},
}


def _codec(codepage: int) -> str:
    """Python codec for a Windows code page; UTF-8 when there is none."""
    codec = _CODEPAGE_CODECS.get(codepage, f"cp{codepage}")
    try:
        codecs.lookup(codec)
    except LookupError:
        return "utf_8"
    return codec


# ---------------------------------------------------------------------------
# PROJECT stream (MS-OVBA 2.3.1): module kinds -> file extensions
# ---------------------------------------------------------------------------

_EXTENSION_BY_KEY = {"Document": "cls", "Module": "bas", "Class": "cls", "BaseClass": "frm"}


def _module_extensions(project_text: str) -> dict[str, str]:
    """Lower-cased module name -> extension, from ``Module=``/``Class=``/... lines."""
    extensions: dict[str, str] = {}
    for line in project_text.splitlines():
        key, sep, value = line.strip().partition("=")
        if not sep or key not in _EXTENSION_BY_KEY:
            continue
        value = value.lower()
        if key == "Document":
            value = value.split("/", 1)[0]  # "Sheet1/&H00000000"
        extensions[value] = _EXTENSION_BY_KEY[key]
    return extensions


# ---------------------------------------------------------------------------
# Projects and modules
# ---------------------------------------------------------------------------


def _project_roots(ole: olefile.OleFileIO) -> list[str]:
    """Storage paths holding a VBA project (MS-OVBA 2.2.1), with trailing "/".

    "" for vbaProject.bin, whose root is the project; "_VBA_PROJECT_CUR/" in
    a legacy .xls.
    """
    roots: list[str] = []
    for storage in ole.listdir(streams=False, storages=True):
        if storage[-1].upper() != "VBA":
            continue
        root = "/".join(storage[:-1])
        root = f"{root}/" if root else ""
        if all(
            ole.exists(root + name) and ole.get_type(root + name) == olefile.STGTY_STREAM
            for name in ("PROJECT", "VBA/_VBA_PROJECT", "VBA/dir")
        ):
            roots.append(root)
    return roots


def _read_project(ole: olefile.OleFileIO, root: str) -> list[VBAModuleSource]:
    dir_bytes = decompress(ole.openstream(root + "VBA/dir").read())
    codepage, records = _parse_dir(dir_bytes)
    codec = _codec(codepage)
    extensions = _module_extensions(
        ole.openstream(root + "PROJECT").read().decode(codec, "replace")
    )

    modules: list[VBAModuleSource] = []
    for record in records:
        name = record.name.decode(codec, "replace")
        candidates = [
            record.stream_name.decode(codec, "replace") if record.stream_name else None,
            record.stream_name_unicode,
            name,
            record.name_unicode,
        ]
        stream_path = next(
            (
                f"{root}VBA/{c}"
                for c in candidates
                if c is not None and ole.exists(f"{root}VBA/{c}")
            ),
            None,
        )
        if stream_path is None:
            continue  # declared in dir but absent: nothing to read
        compressed = ole.openstream(stream_path).read()[record.text_offset :]
        if not compressed:
            continue  # a module with no source text
        code = decompress(compressed).decode(codec, "replace")
        extension = extensions.get(name.lower(), "vba")
        modules.append(
            VBAModuleSource(
                name=name, filename=f"{name}.{extension}", stream_path=stream_path, code=code
            )
        )
    return modules


def is_ole(data: bytes) -> bool:
    """Whether ``data`` is an OLE compound file (as vbaProject.bin must be)."""
    return bool(olefile.isOleFile(data=data))


def read_vba_modules(source: bytes | str | Path) -> list[VBAModuleSource]:
    """All modules of every VBA project in an OLE file, in dir order.

    Args:
        source: The bytes of a vbaProject.bin, or the path of a legacy OLE
            workbook (.xls/.xla).

    Returns:
        The modules; empty when the file holds no VBA project.

    Raises:
        VBAProjectError: if ``source`` is not an OLE file or its VBA project
            cannot be read.
    """
    if isinstance(source, bytes):
        if not olefile.isOleFile(data=source):
            raise VBAProjectError("not an OLE compound file")
    elif not olefile.isOleFile(str(source)):
        raise VBAProjectError("not an OLE compound file")
    try:
        ole = (
            olefile.OleFileIO(source)
            if isinstance(source, bytes)
            else olefile.OleFileIO(str(source))
        )
        with ole:
            modules: list[VBAModuleSource] = []
            for root in _project_roots(ole):
                modules.extend(_read_project(ole, root))
            return modules
    except VBAProjectError:
        raise
    except (OSError, struct.error, ValueError, IndexError) as exc:
        raise VBAProjectError(f"unreadable VBA project: {exc}") from exc


# ---------------------------------------------------------------------------
# Security keyword scan
# ---------------------------------------------------------------------------

# Procedures Office runs on its own when a document opens or closes.
_AUTOEXEC = re.compile(
    r"^\s*(?:Public\s+|Private\s+)?Sub\s+(Auto_?Open|Auto_?Close|AutoExec|AutoNew|AutoExit"
    r"|Workbook_Open|Workbook_Activate|Workbook_BeforeClose|Document_Open|Document_Close"
    r"|Document_New)\b",
    re.IGNORECASE | re.MULTILINE,
)

# (keyword, pattern, description): calls that reach outside the workbook.
_SUSPICIOUS: tuple[tuple[str, str, str], ...] = (
    ("Shell", r"\bShell\s*[\(\"]", "May run an executable or a system command"),
    ("WScript.Shell", r"\bWScript\.Shell\b", "May run an executable or a system command"),
    ("Shell.Application", r"\bShell\.Application\b", "May run an executable or a system command"),
    ("PowerShell", r"\bpowershell\b", "May run a PowerShell command"),
    ("CreateObject", r"\bCreateObject\b", "May create an OLE object outside the workbook"),
    ("GetObject", r"\bGetObject\b", "May attach to an object outside the workbook"),
    ("Declare", r"\bDeclare\s+(?:PtrSafe\s+)?(?:Function|Sub)\b", "Calls a function in a DLL"),
    ("URLDownloadToFile", r"\bURLDownloadToFile\b", "May download a file"),
    ("XMLHTTP", r"\bXMLHTTP\b", "May make a network request"),
    ("WinHttp", r"\bWinHttp", "May make a network request"),
    ("ADODB.Stream", r"\bADODB\.Stream\b", "May read or write binary files"),
    ("SaveToFile", r"\bSaveToFile\b", "May write a file"),
    ("Kill", r"^\s*Kill\b", "May delete a file"),
    ("Open For Output", r"\bOpen\b[^\n]*\bFor\s+(?:Output|Append|Binary)\b", "May write a file"),
    ("Environ", r"\bEnviron\$?\s*\(", "May read system environment variables"),
    ("CallByName", r"\bCallByName\b", "May call a function whose name is computed"),
    ("ExecuteExcel4Macro", r"\bExecuteExcel4Macro\b", "May run an Excel 4 macro"),
    ("SendKeys", r"\bSendKeys\b", "May send keystrokes to another program"),
    ("VBProject", r"\bVBProject\b|\bVBComponents\b", "May read or modify VBA code"),
    ("Registry", r"\bReg(?:Write|Read|Delete)\b", "May access the Windows registry"),
)
_SUSPICIOUS_RE = tuple(
    (keyword, re.compile(pattern, re.IGNORECASE | re.MULTILINE), description)
    for keyword, pattern, description in _SUSPICIOUS
)


def scan_security(codes: list[str]) -> list[dict[str, str]]:
    """Workbook-wide keyword findings, as ``{"type", "keyword", "description"}``.

    A deterministic, sorted list: ``AutoExec`` for procedures Office runs on
    its own, ``Suspicious`` for calls that reach outside the workbook (shell,
    network, file system, DLLs, registry, the VBA project itself). A finding
    flags code to review; it is not a verdict.
    """
    findings: set[tuple[str, str, str]] = set()
    for code in codes:
        for match in _AUTOEXEC.finditer(code):
            findings.add(
                (
                    "AutoExec",
                    match.group(1),
                    "Runs automatically when Office opens or closes the document",
                )
            )
        for keyword, pattern, description in _SUSPICIOUS_RE:
            if pattern.search(code):
                findings.add(("Suspicious", keyword, description))
    return [{"type": t, "keyword": k, "description": d} for t, k, d in sorted(findings)]
