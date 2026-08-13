"""No shipped text I/O may depend on the platform's locale encoding.

Ruff's PLW1514 covers `open()` and `tempfile`, but NOT `Path.read_text()`,
`Path.write_text()` or `Path.open()` — verified against the pinned ruff. Those
are the same defect class (locale-dependent decode/encode; the Windows cp1252
crash this suite exists for), so the gate is completed here with an AST scan
that treats the Path methods as first-class.

Scope is the shipped code: `src/` plus the packages a user can run.
"""

import ast
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SCANNED_DIRS = ("src", "test_workbook_generator", "scripts")

# Methods whose text mode decodes/encodes with locale.getpreferredencoding()
# unless an explicit encoding is passed.
PATH_TEXT_METHODS = {"read_text", "write_text", "open"}


def _binary_mode(call: ast.Call) -> bool:
    for kw in call.keywords:
        if kw.arg == "mode" and isinstance(kw.value, ast.Constant):
            return "b" in str(kw.value.value)
    for arg in call.args:
        if isinstance(arg, ast.Constant) and isinstance(arg.value, str) and "b" in arg.value:
            return True
    return False


def _has_encoding(call: ast.Call) -> bool:
    return any(kw.arg == "encoding" for kw in call.keywords)


def _receiver_name(call: ast.Call) -> str:
    """Best-effort name of the object the method is called on."""
    recv = call.func.value  # type: ignore[union-attr]
    if isinstance(recv, ast.Name):
        return recv.id
    if isinstance(recv, ast.Attribute):
        return recv.attr
    return ""


def _offenders(path: Path) -> list[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    out: list[str] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Attribute):
            continue
        if node.func.attr not in PATH_TEXT_METHODS:
            continue
        recv = _receiver_name(node)
        # zipfile/tarfile members are binary handles, and `self.open()` is a
        # class's own method — neither is locale-decoded text I/O.
        if "zip" in recv.lower() or recv == "self":
            continue
        if _binary_mode(node) or _has_encoding(node):
            continue
        out.append(f"{path.relative_to(REPO_ROOT)}:{node.lineno} .{node.func.attr}()")
    return out


def test_no_locale_dependent_path_text_io():
    offenders: list[str] = []
    for directory in SCANNED_DIRS:
        root = REPO_ROOT / directory
        if not root.is_dir():
            continue
        for py in sorted(root.rglob("*.py")):
            offenders.extend(_offenders(py))
    assert not offenders, (
        "text I/O without an explicit encoding (decodes with the platform code page "
        "on Windows — see the cp1252 crash in tests/test_locale_independence.py):\n  "
        + "\n  ".join(offenders)
    )
