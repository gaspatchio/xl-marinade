"""Strip Excel's absolute-path stamp from saved workbooks.

Whenever real Excel saves a workbook it embeds the save location in
xl/workbook.xml as <x15ac:absPath url="/Users/..."> (inside an
mc:AlternateContent block). Headless openpyxl builds are clean, but the
committed artifacts are Excel-saved (batch-macro evidence + cached
values), so they carry the author's local filesystem path — invisible to
text grep because it sits inside a deflated zip member.

Run this as the LAST step after all Excel saves and BEFORE IR extraction
/ commit, so committed workbooks, their sha256 hashes and the IR
databases stay consistent. Rewrites only xl/workbook.xml; all other zip
members (vbaProject.bin, cached values, macro results) are copied
byte-for-byte. Usage: strip_workbook_abspath.py [paths…] (default: every
committed workbooks/*/<Model>.xlsx|.xlsm).
"""

import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]

ABSPATH_BLOCK = re.compile(
    rb"<mc:AlternateContent[^>]*>(?:(?!</mc:AlternateContent>).)*?"
    rb"x15ac:absPath.*?</mc:AlternateContent>",
    re.DOTALL,
)


def strip(path: Path) -> bool:
    """Remove the absPath block from one workbook. Returns True if changed."""
    with zipfile.ZipFile(path) as zin:
        xml = zin.read("xl/workbook.xml")
        new_xml = ABSPATH_BLOCK.sub(b"", xml)
        if new_xml == xml:
            return False
        with tempfile.NamedTemporaryFile(dir=path.parent, suffix=path.suffix,
                                         delete=False) as tmp:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = new_xml if item.filename == "xl/workbook.xml" \
                        else zin.read(item.filename)
                    zout.writestr(item, data)
    shutil.move(tmp.name, path)
    return True


def main() -> None:
    if len(sys.argv) > 1:
        targets = [Path(a) for a in sys.argv[1:]]
    else:
        targets = sorted(p for d in (REPO / "workbooks").iterdir() if d.is_dir()
                         for p in d.glob("*.xls[xm]")
                         if not p.name.startswith("~$"))
    for path in targets:
        changed = strip(path)
        print(f"{'stripped' if changed else 'clean   '}  {path.relative_to(REPO)}")


if __name__ == "__main__":
    main()
