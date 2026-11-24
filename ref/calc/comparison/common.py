"""Shared pieces of the trust test: ground truth and a cache-free workbook copy."""

import math
import re
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

C_RE = re.compile(rb"<c\b([^>]*?)(/>|>(.*?)</c>)", re.S)
V_RE = re.compile(rb"<v\b[^>]*>.*?</v>|<v\b[^>]*/>", re.S)
T_RE = re.compile(rb'\st="[^"]*"')


def strip_cached(src: Path, dst: Path) -> int:
    """Copy `src`, deleting the cached <v> of every formula cell. Returns cells stripped.

    An engine given this copy cannot hand back stored results: it has to compute.
    """
    n = 0

    def fix(m):
        nonlocal n
        attrs, body = m.group(1), m.group(3)
        if body is None or b"<f" not in body:
            return m.group(0)
        n += 1
        return b"<c" + T_RE.sub(b"", attrs) + b">" + V_RE.sub(b"", body) + b"</c>"

    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet"):
                data = C_RE.sub(fix, data)
            if item.filename == "xl/calcChain.xml":
                pass
            zo.writestr(item, data)
    return n


def ground_truth(path: Path):
    """{(sheet, a1): (formula_text, stored_value, is_array)} for every formula cell."""
    wf = openpyxl.load_workbook(path, keep_vba=False)
    wv = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for ws in wf.worksheets:
        wsv = wv[ws.title]
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                arr = isinstance(v, ArrayFormula)
                if arr:
                    f = v.text
                elif isinstance(v, str) and v.startswith("="):
                    f = v
                else:
                    continue
                out[(ws.title, c.coordinate)] = (f, wsv[c.coordinate].value, arr)
    return out


ERR = {
    "#N/A",
    "#VALUE!",
    "#REF!",
    "#DIV/0!",
    "#NUM!",
    "#NAME?",
    "#NULL!",
    "#N/IMPL!",
    "#ERROR!",
    "#CIRC!",
    "#SPILL!",
    "#CALC!",
}


def norm(v):
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return float(v)
    try:
        import numpy as np

        if isinstance(v, np.generic):
            return norm(v.item())
    except ImportError:
        pass
    if hasattr(v, "isoformat"):
        return ("date", v.isoformat())
    return str(v)


def same(exp, got, rel=1e-9, ab=1e-9):
    e, g = norm(exp), norm(got)
    if e is None and g is None:
        return True
    if isinstance(e, bool) or isinstance(g, bool):
        return (
            e == g
            or (isinstance(e, float) and isinstance(g, bool) and e == float(g))
            or (isinstance(g, float) and isinstance(e, bool) and g == float(e))
        )
    if isinstance(e, float) and isinstance(g, float):
        if math.isnan(e) or math.isnan(g):
            return False
        return abs(e - g) <= max(ab, rel * max(abs(e), abs(g)))
    if e is None and g == 0.0 or g is None and e == 0.0:
        return True  # empty vs zero: reference to a blank cell
    return e == g


FN_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\(")


def functions(formula: str):
    return sorted({m.upper() for m in FN_RE.findall(re.sub(r'"[^"]*"', '""', formula))})
