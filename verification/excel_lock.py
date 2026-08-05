"""Machine-wide Excel lock for coordinating agents.

macOS Excel is a single shared instance, so any two agents driving it
concurrently corrupt each other's runs (severed Apple-event connections,
stale add-in state). Every agent must hold this lock for the duration of
any Excel work — xlwings, AppleScript/osascript, or opening workbooks —
and must NEVER `pkill` Excel (that kills the other agent's session too).

Stdlib only; works with any python3. The lockfile is a fixed, shared path
so every repo on the machine coordinates through the same lock.

Usage:
    # preferred: run a command under the lock (acquire -> run -> release)
    python3 excel_lock.py run --purpose "batch reconciliation" -- osascript job.scpt

    # or manage explicitly (e.g. around a multi-command session)
    python3 excel_lock.py acquire --purpose "VBA bootstrap" [--timeout 600]
    python3 excel_lock.py release
    python3 excel_lock.py status

    # from Python
    from excel_lock import hold
    with hold("value proof"):
        ...drive Excel...

Queueing: `acquire`/`run` block (poll every 2s) until the current holder
releases, up to --timeout seconds. A lock whose owning process is dead is
stale and is broken automatically.
"""

import argparse
import json
import os
import subprocess
import sys
import time
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

LOCK_PATH = Path("/tmp/excel-agent.lock")
POLL_SECONDS = 2.0
DEFAULT_TIMEOUT = 900


def _pid_alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True


def _read_lock():
    try:
        return json.loads(LOCK_PATH.read_text())
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return None


def _is_stale(info) -> bool:
    return not (info and isinstance(info.get("pid"), int) and _pid_alive(info["pid"]))


def acquire(purpose: str, timeout: float = DEFAULT_TIMEOUT, owner_pid=None) -> None:
    """Block until the lock is ours. Raises TimeoutError."""
    deadline = time.monotonic() + timeout
    waited = False
    while True:
        try:
            fd = os.open(LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o666)
            with os.fdopen(fd, "w") as f:
                json.dump({
                    "pid": owner_pid or os.getpid(),
                    "purpose": purpose,
                    "cwd": os.getcwd(),
                    "acquired_at": datetime.now(timezone.utc).isoformat(),
                }, f)
            if waited:
                print(f"[excel-lock] acquired after waiting ({purpose})", file=sys.stderr)
            return
        except FileExistsError:
            info = _read_lock()
            if _is_stale(info):
                try:
                    LOCK_PATH.unlink()
                    print("[excel-lock] broke stale lock "
                          f"(holder pid {info.get('pid') if info else '?'} is gone)",
                          file=sys.stderr)
                except FileNotFoundError:
                    pass
                continue
            if time.monotonic() >= deadline:
                raise TimeoutError(
                    f"Excel lock held by pid {info['pid']} "
                    f"({info.get('purpose', '?')}, since {info.get('acquired_at', '?')}); "
                    f"gave up after {timeout:.0f}s. Check `excel_lock.py status`."
                )
            if not waited:
                print(f"[excel-lock] busy - held by pid {info['pid']} "
                      f"({info.get('purpose', '?')}); waiting...", file=sys.stderr)
                waited = True
            time.sleep(POLL_SECONDS)


def release() -> bool:
    """Release regardless of holder (single-user machine; keep it simple)."""
    try:
        LOCK_PATH.unlink()
        return True
    except FileNotFoundError:
        return False


@contextmanager
def hold(purpose: str, timeout: float = DEFAULT_TIMEOUT):
    acquire(purpose, timeout)
    try:
        yield
    finally:
        release()


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_acq = sub.add_parser("acquire", help="take the lock (blocks while busy)")
    p_acq.add_argument("--purpose", default="unspecified")
    p_acq.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT)
    p_acq.add_argument("--owner-pid", type=int, default=None,
                       help="record this pid as holder (default: this process's parent shell)")

    sub.add_parser("release", help="release the lock")
    sub.add_parser("status", help="show holder, or 'free'")

    p_run = sub.add_parser("run", help="run a command while holding the lock")
    p_run.add_argument("--purpose", default="unspecified")
    p_run.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT)
    p_run.add_argument("command", nargs=argparse.REMAINDER,
                       help="command after `--`")

    args = ap.parse_args()

    if args.cmd == "acquire":
        # default holder = parent (the shell/agent), so the lock survives
        # this short-lived CLI call and staleness tracks the real owner
        acquire(args.purpose, args.timeout, args.owner_pid or os.getppid())
        print("acquired")
        return 0

    if args.cmd == "release":
        print("released" if release() else "was not locked")
        return 0

    if args.cmd == "status":
        info = _read_lock()
        if info is None:
            print("free")
        elif _is_stale(info):
            print(f"stale (holder pid {info.get('pid')} is gone): {info}")
        else:
            print(f"held: {info}")
        return 0

    if args.cmd == "run":
        cmd = args.command
        if cmd and cmd[0] == "--":
            cmd = cmd[1:]
        if not cmd:
            ap.error("no command given after `--`")
        with hold(args.purpose, args.timeout):
            return subprocess.run(cmd).returncode

    return 2


if __name__ == "__main__":
    sys.exit(main())
