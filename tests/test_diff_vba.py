"""`marinade diff` reports VBA changes (issue #47).

The diff compared cells, bindings, names, tables and edges but never read the
VBA tables, so a macro edit -- even deleting every module -- diffed to nothing.

Each test extracts one small macro-free workbook, then writes VBA rows into two
copies of that database, the way extraction records them: one `vba_modules`
row per module, holding its full source, and one `vba_procedures` row per
procedure, holding its body and 1-based inclusive line range. Only the VBA
differs between the two databases, so every `VBA_*` change is attributable to
the edit under test.
"""

import hashlib
import re
import shutil
import sqlite3
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import pytest

import xl_marinade

_PROC_START = re.compile(
    r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+Get)\s+(\w+)",
    re.IGNORECASE,
)
_PROC_END = re.compile(r"^End\s+(Sub|Function|Property)\b", re.IGNORECASE)
_KIND = {"sub": "sub", "function": "function", "property get": "property_get"}


@dataclass(slots=True)
class _Mod:
    name: str
    source: str
    kind: str = "standard"
    # compile_branch per procedure, in source order; '' when omitted.
    branches: list[str] = field(default_factory=list)


def _sha(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _procedures(mod: _Mod) -> list[tuple[str, str, str, str, int, int]]:
    """(name, kind, signature, body, line_start, line_end) for each procedure."""
    lines = mod.source.splitlines()
    procs: list[tuple[str, str, str, str, int, int]] = []
    start: int | None = None
    name = kind = signature = ""
    for i, line in enumerate(lines, start=1):
        stripped = line.strip()
        m = _PROC_START.match(stripped)
        if m and start is None:
            start, name, signature = i, m.group(2), stripped
            kind = _KIND[re.sub(r"\s+", " ", m.group(1).lower())]
        elif _PROC_END.match(stripped) and start is not None:
            procs.append((name, kind, signature, "\n".join(lines[start - 1 : i]), start, i))
            start = None
    return procs


def _write_vba(db: Path, modules: list[_Mod]) -> None:
    with sqlite3.connect(db) as conn:
        for mod in modules:
            cur = conn.execute(
                "INSERT INTO vba_modules (name, kind, source_sha256, source_text) "
                "VALUES (?, ?, ?, ?)",
                (mod.name, mod.kind, _sha(mod.source), mod.source),
            )
            for i, (name, kind, sig, body, start, end) in enumerate(_procedures(mod)):
                branch = mod.branches[i] if i < len(mod.branches) else ""
                conn.execute(
                    "INSERT INTO vba_procedures (module_id, name, kind, signature, body, "
                    "normalized_body_hash, line_start, line_end, compile_branch) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (cur.lastrowid, name, kind, sig, body, _sha(body), start, end, branch),
                )


@pytest.fixture(scope="module")
def base_db(tmp_path_factory: pytest.TempPathFactory) -> Path:
    tmp = tmp_path_factory.mktemp("vba_base")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calc"
    ws["A1"], ws["B1"] = "Rate", 0.05
    ws["A2"], ws["B2"] = "Result", "=B1*100"
    wb.save(tmp / "base.xlsx")
    return xl_marinade.extract(tmp / "base.xlsx", tmp / "base.db")


def _diff(base_db: Path, tmp_path: Path, mods_a: list[_Mod], mods_b: list[_Mod]) -> dict:
    db_a, db_b = tmp_path / "a.db", tmp_path / "b.db"
    shutil.copy(base_db, db_a)
    shutil.copy(base_db, db_b)
    _write_vba(db_a, mods_a)
    _write_vba(db_b, mods_b)
    return xl_marinade.diff(db_a, db_b)


def _vba(result: dict) -> list[dict]:
    return [c for c in result["changes"] if c["type"].startswith("VBA_")]


def _types(result: dict) -> list[str]:
    return [c["type"] for c in _vba(result)]


# A body with well over the matcher's 20-distinct-token floor, so the fuzzy
# passes (rename detection) are exercised rather than skipped.
_TOKENS = [f"term{chr(97 + i)}{chr(97 + j)}" for i in range(3) for j in range(9)]


def _calc_body(name: str, extra: str = "") -> str:
    lines = [f"Sub {name}()", "    Dim total As Double"]
    lines += [f"    total = total + {t} * {k}" for k, t in enumerate(_TOKENS, start=1)]
    if extra:
        lines.append(extra)
    lines += ['    Worksheets("Calc").Range("B1").Value = total', "End Sub"]
    return "\n".join(lines)


_HEADER = 'Attribute VB_Name = "Calc"\nOption Explicit\nConst RATE_CAP As Double = 0.1\n'

_HELPER = """Function Discount(ByVal r As Double) As Double
    Discount = 1 / (1 + r)
End Function"""


def _calc(*procs: str, header: str = _HEADER, name: str = "Calc.bas") -> _Mod:
    return _Mod(name=name, source=header + "\n".join(procs) + "\n")


# ---------------------------------------------------------------------------


def test_identical_vba_diffs_to_nothing(base_db: Path, tmp_path: Path) -> None:
    mods = [_calc(_calc_body("Run"), _HELPER)]
    result = _diff(base_db, tmp_path, mods, mods)
    assert _vba(result) == []
    assert all(v == 0 for k, v in result["summary"].items() if k.startswith("vba_"))


def test_vba_free_databases_report_zero_vba_counts(base_db: Path, tmp_path: Path) -> None:
    result = _diff(base_db, tmp_path, [], [])
    assert result["changes"] == []
    vba_keys = {k for k in result["summary"] if k.startswith("vba_")}
    assert vba_keys == {
        "vba_modules_added",
        "vba_modules_removed",
        "vba_declarations_changed",
        "vba_procedures_added",
        "vba_procedures_removed",
        "vba_procedures_renamed",
        "vba_procedures_logic_changed",
        "vba_cosmetic_changes",
    }


def test_removing_all_vba_is_reported(base_db: Path, tmp_path: Path) -> None:
    """The #47 reproduction: every module deleted diffed to zero changes."""
    sheet = _Mod(name="Sheet1.cls", kind="document", source='Attribute VB_Name = "Sheet1"\n')
    result = _diff(base_db, tmp_path, [_calc(_calc_body("Run"), _HELPER), sheet], [])

    assert sorted(_types(result)) == [
        "VBA_MODULE_REMOVED",
        "VBA_MODULE_REMOVED",
        "VBA_PROCEDURE_REMOVED",
        "VBA_PROCEDURE_REMOVED",
    ]
    removed = {c["module"]: c for c in _vba(result) if c["type"] == "VBA_MODULE_REMOVED"}
    assert removed["Calc.bas"]["procedure_count"] == 2
    assert removed["Sheet1.cls"]["module_kind"] == "document"
    s = result["summary"]
    assert (s["vba_modules_removed"], s["vba_procedures_removed"]) == (2, 2)
    assert {c["layer"] for c in _vba(result)} == {"workbook"}


def test_added_module_and_procedure(base_db: Path, tmp_path: Path) -> None:
    result = _diff(
        base_db,
        tmp_path,
        [_calc(_calc_body("Run"))],
        [_calc(_calc_body("Run"), _HELPER), _Mod(name="Util.bas", source="Sub Ping()\nEnd Sub\n")],
    )
    added = sorted((c["type"], c.get("module"), c.get("procedure")) for c in _vba(result))
    assert added == [
        ("VBA_MODULE_ADDED", "Util.bas", None),
        ("VBA_PROCEDURE_ADDED", "Calc.bas", "Discount"),
        ("VBA_PROCEDURE_ADDED", "Util.bas", "Ping"),
    ]
    discount = next(c for c in _vba(result) if c.get("procedure") == "Discount")
    assert discount["kind"] == "function"
    assert discount["signature"].startswith("Function Discount(")


def test_logic_change(base_db: Path, tmp_path: Path) -> None:
    result = _diff(
        base_db,
        tmp_path,
        [_calc(_calc_body("Run"), _HELPER)],
        [_calc(_calc_body("Run", extra="    total = total * 2"), _HELPER)],
    )
    (change,) = _vba(result)
    assert change["type"] == "VBA_PROCEDURE_LOGIC_CHANGED"
    assert (change["module"], change["procedure"], change["kind"]) == ("Calc.bas", "Run", "sub")
    assert change["layer"] == "workbook"
    assert "signature_a" not in change
    assert result["summary"]["vba_procedures_logic_changed"] == 1


def test_signature_change(base_db: Path, tmp_path: Path) -> None:
    helper_b = _HELPER.replace("ByVal r As Double", "ByVal r As Double, ByVal t As Long")
    result = _diff(
        base_db, tmp_path, [_calc(_calc_body("Run"), _HELPER)], [_calc(_calc_body("Run"), helper_b)]
    )
    (change,) = _vba(result)
    assert change["type"] == "VBA_PROCEDURE_SIGNATURE_CHANGED"
    assert change["signature_a"] == "Function Discount(ByVal r As Double) As Double"
    assert change["signature_b"].endswith("ByVal t As Long) As Double")


def test_comment_only_edit_is_cosmetic(base_db: Path, tmp_path: Path) -> None:
    result = _diff(
        base_db,
        tmp_path,
        [_calc(_calc_body("Run"), _HELPER)],
        [_calc(_calc_body("Run", extra="    ' sum the terms"), _HELPER)],
    )
    assert _types(result) == ["VBA_PROCEDURE_COSMETIC_CHANGED"]
    assert result["summary"]["vba_cosmetic_changes"] == 1
    assert result["summary"]["vba_procedures_logic_changed"] == 0


def test_rename(base_db: Path, tmp_path: Path) -> None:
    result = _diff(
        base_db,
        tmp_path,
        [_calc(_calc_body("Run"), _HELPER)],
        [_calc(_calc_body("RunAll"), _HELPER)],
    )
    (change,) = _vba(result)
    assert change["type"] == "VBA_PROCEDURE_RENAMED"
    assert (change["procedure_a"], change["procedure_b"]) == ("Run", "RunAll")
    assert result["summary"]["vba_procedures_renamed"] == 1


def test_rename_with_edit(base_db: Path, tmp_path: Path) -> None:
    result = _diff(
        base_db,
        tmp_path,
        [_calc(_calc_body("Run"), _HELPER)],
        [_calc(_calc_body("RunAll", extra="    total = total * 2"), _HELPER)],
    )
    (change,) = _vba(result)
    assert change["type"] == "VBA_PROCEDURE_RENAMED_AND_MODIFIED"
    assert result["summary"]["vba_procedures_renamed"] == 1


def test_declarations_change(base_db: Path, tmp_path: Path) -> None:
    """Code outside every procedure is compared too: a module constant here."""
    result = _diff(
        base_db,
        tmp_path,
        [_calc(_calc_body("Run"), _HELPER)],
        [_calc(_calc_body("Run"), _HELPER, header=_HEADER.replace("0.1", "0.2"))],
    )
    (change,) = _vba(result)
    assert change["type"] == "VBA_DECLARATIONS_CHANGED"
    assert change["module"] == "Calc.bas"
    assert result["summary"]["vba_declarations_changed"] == 1


def test_declarations_comment_is_cosmetic(base_db: Path, tmp_path: Path) -> None:
    result = _diff(
        base_db,
        tmp_path,
        [_calc(_calc_body("Run"), _HELPER)],
        [_calc(_calc_body("Run"), _HELPER, header=_HEADER + "' rates are annual\n")],
    )
    assert _types(result) == ["VBA_DECLARATIONS_COSMETIC_CHANGED"]
    assert result["summary"]["vba_cosmetic_changes"] == 1


def test_compile_branch_twins_are_compared_branch_to_branch(base_db: Path, tmp_path: Path) -> None:
    """`#If VBA7` twins share module, name and kind; only the branch tells them apart."""

    def twins(win64_extra: str) -> _Mod:
        src = (
            "#If VBA7 Then\n"
            + _calc_body("Run", extra=win64_extra)
            + "\n#Else\n"
            + _calc_body("Run")
            + "\n#End If\n"
        )
        return _Mod(name="Calc.bas", source=src, branches=["VBA7", "NOT VBA7"])

    unchanged = _diff(base_db, tmp_path, [twins("")], [twins("")])
    assert _vba(unchanged) == []

    result = _diff(base_db, tmp_path, [twins("")], [twins("    total = -total")])
    (change,) = _vba(result)
    assert change["type"] == "VBA_PROCEDURE_LOGIC_CHANGED"
    assert change["compile_branch"] == "VBA7"


def test_ambiguous_rename_is_reported_as_removed(base_db: Path, tmp_path: Path) -> None:
    """Two equally good rename targets: the old procedure must not vanish silently."""
    result = _diff(
        base_db,
        tmp_path,
        [_calc(_calc_body("Run"))],
        [_calc(_calc_body("RunA"), _calc_body("RunB"))],
    )
    by_type = sorted((c["type"], c["procedure"]) for c in _vba(result))
    assert by_type == [
        ("VBA_PROCEDURE_ADDED", "RunA"),
        ("VBA_PROCEDURE_ADDED", "RunB"),
        ("VBA_PROCEDURE_REMOVED", "Run"),
    ]
    removed = next(c for c in _vba(result) if c["type"] == "VBA_PROCEDURE_REMOVED")
    assert removed["ambiguous_rename"] is True


def test_database_without_vba_tables_reads_as_no_vba(base_db: Path, tmp_path: Path) -> None:
    db_a, db_b = tmp_path / "a.db", tmp_path / "b.db"
    shutil.copy(base_db, db_a)
    shutil.copy(base_db, db_b)
    with sqlite3.connect(db_a) as conn:
        for table in ("vba_chunks", "vba_procedure_cell_refs", "vba_procedure_edges"):
            conn.execute(f"DROP TABLE IF EXISTS {table}")
        conn.execute("DROP TABLE vba_procedures")
        conn.execute("DROP TABLE vba_modules")
    _write_vba(db_b, [_calc(_calc_body("Run"))])

    result = xl_marinade.diff(db_a, db_b)
    assert sorted(_types(result)) == ["VBA_MODULE_ADDED", "VBA_PROCEDURE_ADDED"]


def test_vba_changes_are_ordered_module_first(base_db: Path, tmp_path: Path) -> None:
    result = _diff(base_db, tmp_path, [], [_calc(_calc_body("Run"), _HELPER)])
    types = [c["type"] for c in result["changes"]]
    assert types == ["VBA_MODULE_ADDED", "VBA_PROCEDURE_ADDED", "VBA_PROCEDURE_ADDED"]
    assert [c["seq"] for c in result["changes"]] == [1, 2, 3]
