"""One-time bootstrap: produce vbaProject.bin for the SmithWilson workbook.

macOS Excel has no scriptable VBE, so this drives the UI once:
open the built .xlsx, import BatchRunner.bas via the VBE (GUI automation),
save as .xlsm, then extract xl/vbaProject.bin into this folder. After that,
build_workbook.py injects the bin on every build with no GUI involved.

Requires: Excel installed, Accessibility permission for the terminal.
Re-run only when BatchRunner.bas changes.
"""

import subprocess
import sys
import time
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1] / "verification"))
from excel_lock import hold  # noqa: E402
XLSX = HERE / "SmithWilson.xlsx"
BAS = HERE / "BatchRunner.bas"
TMP_XLSM = HERE / "_bootstrap_tmp.xlsm"
BIN = HERE / "vbaProject.bin"


def _foreign_workbooks(our_name):
    """Names of open workbooks that are not ours - another agent's session."""
    res = subprocess.run(
        ["osascript", "-e",
         'tell application "Microsoft Excel" to get name of every workbook'],
        capture_output=True, text=True)
    if res.returncode != 0:
        return []
    names = [n.strip() for n in res.stdout.strip().split(",") if n.strip()]
    return [n for n in names if n not in (our_name, "Book1", "", "missing value")]


def osa(script: str) -> str:
    res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"osascript failed: {res.stderr.strip()}\n--- script ---\n{script}")
    return res.stdout.strip()


def harvest(xlsm_path: Path):
    """Extract xl/vbaProject.bin from a manually-saved .xlsm."""
    with zipfile.ZipFile(xlsm_path) as z:
        if "xl/vbaProject.bin" not in z.namelist():
            sys.exit(f"no vbaProject.bin inside {xlsm_path}")
        BIN.write_bytes(z.read("xl/vbaProject.bin"))
    print(f"wrote {BIN} ({BIN.stat().st_size} bytes)")


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--harvest":
        harvest(Path(sys.argv[2]))
        return
    if not XLSX.exists():
        sys.exit(f"{XLSX} missing - run build_workbook.py first")
    with hold("VBA bootstrap (VBE import)"):
        _bootstrap()


def _bootstrap():
    TMP_XLSM.unlink(missing_ok=True)

    # NEVER quit Excel while another agent's workbooks are open - wait for
    # them to finish instead (the Excel lock is cooperative; this guard
    # protects agents that do not take it)
    for _ in range(60):
        foreign = _foreign_workbooks(XLSX.name)
        if not foreign:
            break
        print(f"[bootstrap] waiting - another session has {foreign} open...",
              file=sys.stderr)
        time.sleep(10)
    else:
        raise RuntimeError(f"gave up: foreign workbooks still open: {foreign}")

    # Fresh Excel with only our workbook open. Quit is asynchronous: wait
    # for the process to actually exit, or the dying instance tears down
    # the relaunched one mid-bootstrap.
    subprocess.run(["osascript", "-e", 'tell application "Microsoft Excel" to quit'],
                   capture_output=True)
    for _ in range(30):
        if subprocess.run(["pgrep", "-x", "Microsoft Excel"],
                          capture_output=True).returncode != 0:
            break
        time.sleep(1)
    time.sleep(1)
    osa(f'''
    tell application "Microsoft Excel"
        activate
        open POSIX file "{XLSX}"
    end tell''')
    # Wait until Excel is fully up (menu bar reachable) - large workbooks
    # take longer than a fixed sleep
    for _ in range(30):
        time.sleep(2)
        try:
            osa('''
            tell application "System Events" to tell process "Microsoft Excel"
                get name of menu bar 1
            end tell''')
            break
        except RuntimeError:
            continue
    else:
        raise RuntimeError("Excel menu bar never became available")

    # Open the VBE (Tools > Macro > Visual Basic Editor)
    osa('''
    tell application "System Events" to tell process "Microsoft Excel"
        set frontmost to true
        click menu item "Visual Basic Editor" of menu "Macro" of menu item "Macro" of menu "Tools" of menu bar 1
    end tell''')
    # The VBE replaces the menu bar; poll until its File menu is up
    for _ in range(15):
        time.sleep(2)
        try:
            osa('''
            tell application "System Events" to tell process "Microsoft Excel"
                get menu item "Import File..." of menu "File" of menu bar 1
            end tell''')
            break
        except RuntimeError:
            continue
    else:
        raise RuntimeError("VBE menu bar never became available")

    # VBE: File > Import File..., then type the path into the open panel
    osa('''
    tell application "System Events" to tell process "Microsoft Excel"
        click menu item "Import File..." of menu "File" of menu bar 1
    end tell''')
    # Wait for the open panel (its own window named "Import File") before typing
    for _ in range(20):
        time.sleep(1)
        out = osa('''
        tell application "System Events" to tell process "Microsoft Excel"
            if exists window "Import File" then
                return "panel"
            else
                return "no"
            end if
        end tell''')
        if out == "panel":
            break
    else:
        raise RuntimeError("Import File open panel never appeared")
    osa(f'''
    tell application "System Events" to tell process "Microsoft Excel"
        keystroke "g" using {{command down, shift down}}
        delay 1.5
        keystroke "{BAS}"
        delay 1
        keystroke return
        delay 1.5
        keystroke return
    end tell''')
    time.sleep(3)

    # Save as macro-enabled workbook (scriptable, no GUI needed)
    osa(f'''
    tell application "Microsoft Excel"
        save workbook as active workbook filename POSIX file "{TMP_XLSM}" file format macro enabled XML file format with overwrite
    end tell''')
    time.sleep(2)
    osa('tell application "Microsoft Excel" to close active workbook saving no')
    subprocess.run(["osascript", "-e", 'tell application "Microsoft Excel" to quit'],
                   capture_output=True)

    if not TMP_XLSM.exists():
        sys.exit("save-as-xlsm did not produce a file")
    with zipfile.ZipFile(TMP_XLSM) as z:
        names = z.namelist()
        if "xl/vbaProject.bin" not in names:
            sys.exit(f"no vbaProject.bin in saved file (module import failed?). Parts: {names}")
        BIN.write_bytes(z.read("xl/vbaProject.bin"))
    TMP_XLSM.unlink()
    print(f"wrote {BIN} ({BIN.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
