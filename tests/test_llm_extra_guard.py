"""Importing the [llm] add-on without its deps gives an actionable error;
the core still imports and works."""

import os
import subprocess
import sys
import textwrap


def test_llm_guard_message_without_openai():
    code = textwrap.dedent(
        """
        import builtins
        _real = builtins.__import__

        def _guard(name, *a, **k):
            if name == "openai" or name.startswith("openai."):
                raise ImportError("no openai")
            return _real(name, *a, **k)

        builtins.__import__ = _guard

        import xl_marinade          # the core must still import with openai absent
        import xl_marinade.docs     # the free docs tier must still import with openai absent

        try:
            import xl_marinade.llm  # noqa: F401
        except ImportError as e:
            assert "xl-marinade[llm]" in str(e), str(e)
            print("guard-ok")
        else:
            raise SystemExit("expected ImportError from the [llm] guard")
        """
    )
    env = {**os.environ, "PYTHONPATH": "src"}
    result = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, env=env)
    assert result.returncode == 0, result.stderr
    assert "guard-ok" in result.stdout
