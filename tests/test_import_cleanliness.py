"""The core imports open no sockets and have no import-time side effects.

This mechanically enforces the promise the whole product rests on: a bare
install of the deterministic core never touches the network on import.
"""

import os
import subprocess
import sys
import textwrap


def test_importing_xl_marinade_opens_no_socket():
    code = textwrap.dedent(
        """
        import socket

        def _deny(*args, **kwargs):
            raise AssertionError("network access at import time")

        socket.socket = _deny

        import xl_marinade
        import xl_marinade.core
        import xl_marinade.core.api
        import xl_marinade.errors

        print("clean")
        """
    )
    env = {**os.environ, "PYTHONPATH": "src"}
    result = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, env=env)
    assert result.returncode == 0, result.stderr
    assert "clean" in result.stdout
