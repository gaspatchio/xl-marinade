"""The olefile-based VBA project reader (xl_marinade.core.vba.ovba).

VBA source used to come from oletools, whose required dependency pcodedmp is
GPLv3. The reader is implemented from Microsoft's [MS-OVBA] spec on olefile
alone. These tests pin the spec behaviour without a binary fixture: the
decompressor against a hand-worked container and a spec compressor, and the
dir/PROJECT parsing against a fake OLE container built here. Parity with
oletools on real workbooks (80 modules across the lifelib examples, byte for
byte) was checked when the reader was written.
"""

import struct
from dataclasses import dataclass, field

import pytest

from xl_marinade.core.vba import ovba
from xl_marinade.core.vba.ovba import VBAProjectError, decompress, read_vba_modules, scan_security

# ---------------------------------------------------------------------------
# Decompression (MS-OVBA 2.4.1)
# ---------------------------------------------------------------------------


def test_decompress_hand_worked_container() -> None:
    """Nine "a"s: one literal, then a copy of 8 bytes from 1 back.

    Chunk = header(2) + flag byte + "a" + copy token(2) = 6 bytes, so the
    header's size field is 6 - 3 = 3, plus signature 0b011 (0x3000) and the
    compressed bit (0x8000): 0xB003. After one byte of output the copy token
    uses 4 offset bits: length 8 - 3 = 5, offset 1 - 1 = 0, token 0x0005. The
    flag byte marks token 1 (the copy): 0x02.
    """
    assert decompress(b"\x01\x03\xb0\x02a\x05\x00") == b"a" * 9


def test_decompress_raw_chunk() -> None:
    """An uncompressed chunk: header bit 15 clear, then 4096 literal bytes."""
    raw = bytes(range(256)) * 16
    assert decompress(b"\x01" + struct.pack("<H", 0x3FFF) + raw) == raw


def _compress(data: bytes, window: int = 512) -> bytes:
    """A small MS-OVBA compressor (greedy matching), to round-trip with."""
    out = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        chunk = data[start : start + 4096]
        body = bytearray()
        i = 0
        while i < len(chunk):
            flag_at = len(body)
            body.append(0)
            for bit in range(8):
                if i >= len(chunk):
                    break
                offset_bits = max((i - 1).bit_length(), 4) if i else 4
                max_len = (0xFFFF >> offset_bits) + 3
                max_off = min(i, 1 << (16 - offset_bits), window)
                best_len = best_off = 0
                for off in range(1, max_off + 1):
                    n = 0
                    while n < max_len and i + n < len(chunk) and chunk[i + n - off] == chunk[i + n]:
                        n += 1
                    if n > best_len:
                        best_len, best_off = n, off
                if best_len >= 3:
                    token = ((best_off - 1) << (16 - offset_bits)) | (best_len - 3)
                    body += struct.pack("<H", token)
                    body[flag_at] |= 1 << bit
                    i += best_len
                else:
                    body.append(chunk[i])
                    i += 1
        assert len(body) <= 4096, "test data must be compressible"
        out += struct.pack("<H", 0xB000 | (len(body) - 1)) + body
    return bytes(out)


def test_decompress_round_trips_multi_chunk_source() -> None:
    """Long, repetitive VBA: several chunks, every offset/length bit split."""
    lines = [
        f'    total = total + Worksheets("Calc").Cells({r}, {c}).Value * {r * c}'
        for r in range(1, 40)
        for c in range(1, 6)
    ]
    source = ("Sub Run()\r\n" + "\r\n".join(lines) + "\r\nEnd Sub\r\n").encode("cp1252")
    assert len(source) > 3 * 4096
    assert decompress(_compress(source)) == source


@pytest.mark.parametrize(
    "data",
    [
        b"",
        b"\x02\x03\xb0\x02a\x05\x00",  # bad container signature
        b"\x01\x03\x00\x02a\x05\x00",  # bad chunk signature
        b"\x01\x03\xb0\x02a\x05\xf0",  # copy token 16 back, with 1 byte of output
    ],
)
def test_decompress_rejects_malformed_input(data: bytes) -> None:
    with pytest.raises(VBAProjectError):
        decompress(data)


# ---------------------------------------------------------------------------
# dir / PROJECT streams, against a fake OLE container
# ---------------------------------------------------------------------------


def _rec(record_id: int, data: bytes) -> bytes:
    return struct.pack("<HI", record_id, len(data)) + data


def _dir_stream(modules: list[dict], codepage: int = 1252) -> bytes:
    """A dir stream with the records a real project carries, in spec order."""
    out = bytearray()
    out += _rec(0x0001, struct.pack("<I", 3))  # PROJECTSYSKIND
    out += _rec(0x0002, struct.pack("<I", 0x409))  # PROJECTLCID
    out += _rec(0x0014, struct.pack("<I", 0x409))  # PROJECTLCIDINVOKE
    out += _rec(0x0003, struct.pack("<H", codepage))  # PROJECTCODEPAGE
    out += _rec(0x0004, b"VBAProject")  # PROJECTNAME
    out += _rec(0x0005, b"") + _rec(0x0040, b"")  # PROJECTDOCSTRING (+ unicode)
    out += _rec(0x0006, b"") + _rec(0x003D, b"")  # PROJECTHELPFILEPATH
    out += _rec(0x0007, struct.pack("<I", 0))  # PROJECTHELPCONTEXT
    out += _rec(0x0008, struct.pack("<I", 0))  # PROJECTLIBFLAGS
    # PROJECTVERSION: its size field reads 4 but 6 bytes follow.
    out += struct.pack("<HI", 0x0009, 4) + struct.pack("<IH", 1, 2)
    out += _rec(0x000C, b"") + _rec(0x003C, b"")  # PROJECTCONSTANTS
    out += _rec(0x0016, b"stdole") + _rec(0x003E, "stdole".encode("utf-16-le"))  # REFERENCENAME
    out += _rec(0x000D, b"\x00" * 20)  # REFERENCEREGISTERED
    out += _rec(0x000F, struct.pack("<H", len(modules)))  # PROJECTMODULES
    out += _rec(0x0013, struct.pack("<H", 0xFFFF))  # PROJECTCOOKIE
    for m in modules:
        out += _rec(0x0019, m["name"].encode("cp1252"))
        out += _rec(0x0047, m["name"].encode("utf-16-le"))
        if m.get("stream") is not None:
            out += _rec(0x001A, m["stream"].encode("cp1252"))
            out += _rec(0x0032, m.get("stream_unicode", m["stream"]).encode("utf-16-le"))
        out += _rec(0x001C, b"") + _rec(0x0048, b"")  # MODULEDOCSTRING
        out += _rec(0x0031, struct.pack("<I", m["offset"]))  # MODULEOFFSET
        out += _rec(0x001E, struct.pack("<I", 0))  # MODULEHELPCONTEXT
        out += _rec(0x002C, struct.pack("<H", 0xFFFF))  # MODULECOOKIE
        out += _rec(0x0021 if m["kind"] == "bas" else 0x0022, b"")  # MODULETYPE
        out += _rec(0x002B, b"")  # MODULE terminator
    out += _rec(0x0010, b"")  # dir terminator
    return bytes(out)


@dataclass
class _FakeOle:
    """Just the olefile surface the reader uses."""

    streams: dict[str, bytes] = field(default_factory=dict)

    def listdir(self, streams: bool = True, storages: bool = False) -> list[list[str]]:
        storages_seen = {tuple(p.split("/")[:-1]) for p in self.streams if "/" in p}
        return [list(s) for s in sorted(storages_seen)]

    def exists(self, path: str) -> bool:
        return path.lower() in {p.lower() for p in self.streams}

    def get_type(self, path: str) -> int:
        return ovba.olefile.STGTY_STREAM

    def openstream(self, path: str) -> "_Stream":
        for p, data in self.streams.items():
            if p.lower() == path.lower():
                return _Stream(data)
        raise OSError(path)


@dataclass
class _Stream:
    data: bytes

    def read(self) -> bytes:
        return self.data


def _project(modules: list[dict], project_text: str, codepage: int = 1252) -> _FakeOle:
    streams = {
        "PROJECT": project_text.encode("cp1252"),
        "VBA/_VBA_PROJECT": b"\xcc\x61\xff\xff\x00",
        "VBA/dir": _compress(_dir_stream(modules, codepage)),
    }
    for m in modules:
        if m.get("source") is None:
            continue
        cache = b"\x00" * m["offset"]  # the performance cache precedes the source
        compressed = _compress(m["source"].encode("cp1252")) if m["source"] else b""
        streams[f"VBA/{m.get('stream_unicode') or m['stream'] or m['name']}"] = cache + compressed
    return _FakeOle(streams)


_PROJECT_TEXT = (
    'ID="{00000000-0000-0000-0000-000000000000}"\r\n'
    "Document=ThisWorkbook/&H00000000\r\n"
    "Document=Sheet1/&H00000000\r\n"
    "Module=Calc\r\n"
    "Class=Ledger\r\n"
    "BaseClass=Dialog\r\n"
    'Name="VBAProject"\r\n'
)


def test_reads_modules_in_dir_order_with_extensions() -> None:
    modules = [
        {
            "name": "ThisWorkbook",
            "stream": "ThisWorkbook",
            "offset": 40,
            "kind": "cls",
            "source": 'Attribute VB_Name = "ThisWorkbook"\r\n',
        },
        {
            "name": "Sheet1",
            "stream": "Sheet1",
            "offset": 0,
            "kind": "cls",
            "source": 'Attribute VB_Name = "Sheet1"\r\n',
        },
        {
            "name": "Calc",
            "stream": "Calc",
            "offset": 123,
            "kind": "bas",
            "source": "Function Price(r As Double) As Double\r\n    Price = r * 2 ' € cost\r\nEnd Function\r\n",
        },
        {
            "name": "Ledger",
            "stream": "Ledger",
            "offset": 8,
            "kind": "cls",
            "source": "Private x As Long\r\n",
        },
        {
            "name": "Dialog",
            "stream": "Dialog",
            "offset": 8,
            "kind": "cls",
            "source": "Sub Show()\r\nEnd Sub\r\n",
        },
        {
            "name": "Loose",
            "stream": "Loose",
            "offset": 8,
            "kind": "bas",
            "source": "Sub Go()\r\nEnd Sub\r\n",
        },
    ]
    ole = _project(modules, _PROJECT_TEXT)
    assert ovba._project_roots(ole) == [""]  # type: ignore[arg-type]
    result = ovba._read_project(ole, "")  # type: ignore[arg-type]
    assert [m.filename for m in result] == [
        "ThisWorkbook.cls",
        "Sheet1.cls",
        "Calc.bas",
        "Ledger.cls",
        "Dialog.frm",
        "Loose.vba",  # not listed in PROJECT: oletools' default extension
    ]
    calc = result[2]
    assert calc.code == modules[2]["source"]  # "€" is byte 0x80 in the project's cp1252
    assert calc.stream_path == "VBA/Calc"


def test_unicode_stream_name_is_the_fallback() -> None:
    modules = [
        {
            "name": "Calc",
            "stream": "Mangled?",
            "stream_unicode": "Calc",
            "offset": 4,
            "kind": "bas",
            "source": "Sub A()\r\nEnd Sub\r\n",
        },
    ]
    ole = _project(modules, "Module=Calc\r\n")
    (module,) = ovba._read_project(ole, "")  # type: ignore[arg-type]
    assert module.stream_path == "VBA/Calc"
    assert module.code == "Sub A()\r\nEnd Sub\r\n"


def test_module_without_source_is_skipped() -> None:
    modules = [
        {"name": "Empty", "stream": "Empty", "offset": 16, "kind": "bas", "source": ""},
        {
            "name": "Calc",
            "stream": "Calc",
            "offset": 0,
            "kind": "bas",
            "source": "Sub A()\r\nEnd Sub\r\n",
        },
    ]
    ole = _project(modules, "Module=Empty\r\nModule=Calc\r\n")
    assert [m.filename for m in ovba._read_project(ole, "")] == ["Calc.bas"]  # type: ignore[arg-type]


def test_legacy_project_root_is_found() -> None:
    ole = _project(
        [
            {
                "name": "Calc",
                "stream": "Calc",
                "offset": 0,
                "kind": "bas",
                "source": "Sub A()\r\nEnd Sub\r\n",
            }
        ],
        "Module=Calc\r\n",
    )
    ole.streams = {f"_VBA_PROJECT_CUR/{p}": d for p, d in ole.streams.items()}
    assert ovba._project_roots(ole) == ["_VBA_PROJECT_CUR/"]  # type: ignore[arg-type]
    (module,) = ovba._read_project(ole, "_VBA_PROJECT_CUR/")  # type: ignore[arg-type]
    assert module.filename == "Calc.bas"


@pytest.mark.parametrize(
    "codepage, codec", [(1252, "cp1252"), (65001, "utf_8"), (10000, "mac_roman"), (99999, "utf_8")]
)
def test_codepage_codecs(codepage: int, codec: str) -> None:
    assert ovba._codec(codepage) == codec


def test_non_ole_input_is_an_error() -> None:
    with pytest.raises(VBAProjectError):
        read_vba_modules(b"not an OLE compound file")


# ---------------------------------------------------------------------------
# Security keyword scan
# ---------------------------------------------------------------------------


def test_scan_flags_autoexec_and_outside_calls() -> None:
    code = (
        "Private Sub Workbook_Open()\r\n"
        '    Shell "cmd /c calc"\r\n'
        '    Set o = CreateObject("Scripting.FileSystemObject")\r\n'
        "End Sub\r\n"
    )
    findings = scan_security([code])
    assert [(f["type"], f["keyword"]) for f in findings] == [
        ("AutoExec", "Workbook_Open"),
        ("Suspicious", "CreateObject"),
        ("Suspicious", "Shell"),
    ]
    assert all(f["description"] for f in findings)


def test_scan_ignores_ordinary_model_code() -> None:
    """The words olevba flagged on the lifelib models: "run" in a comment,
    a column named pv_av_change, a calculation loop."""
    code = (
        "' one row per run from row 13\r\n"
        "' pv_av_change, pv_net_cf (cols 3-9)\r\n"
        "Sub RunCurveReconciliation()\r\n"
        "    Application.Calculate\r\n"
        "End Sub\r\n"
    )
    assert scan_security([code]) == []
