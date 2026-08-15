"""Every workflow must pin the uv version, and they must all pin the same one.

uv decides the core metadata version written into the wheel. `release.yml`
tracked `latest`, uv began emitting `Metadata-Version: 2.5`, and the pinned
publisher rejected every artifact — v0.3.0 built cleanly and then failed to
publish. The build toolchain moving under a pinned publisher is a standing
hazard, not a one-off.

Two things are checked, and the second is the one that actually bit:

* No step may pin `latest`.
* Every step must pin *explicitly*. `setup-uv` defaults to the newest release
  when `version:` is omitted, so an omitted input is `latest` with nothing to
  grep for — which is exactly the shape `release.yml` was in.

Also asserts the pins agree, so CI cannot pass on one uv while releases publish
from another.

Parsed by line rather than with a YAML library on purpose: this is the only
YAML introspection in the suite, and it does not justify a dependency the
package does not otherwise need.
"""

import re
from pathlib import Path

WORKFLOWS = Path(__file__).resolve().parent.parent / ".github" / "workflows"

SETUP_UV = re.compile(r"^(?P<indent>\s*)-\s+uses:\s*astral-sh/setup-uv@")
VERSION = re.compile(r"^\s*version:\s*[\"']?(?P<value>[^\"'\s#]+)")
# A step ends where the next list item at the same indent begins.
NEXT_STEP = re.compile(r"^\s*-\s")


def _setup_uv_pins() -> list[tuple[str, int, str | None]]:
    """Return (workflow, line number, pinned version or None) per setup-uv step."""
    found: list[tuple[str, int, str | None]] = []
    # Both extensions: GitHub Actions runs `.yaml` too, so globbing only `.yml`
    # would let a new workflow float its uv version with the guard still green.
    for path in sorted([*WORKFLOWS.glob("*.yml"), *WORKFLOWS.glob("*.yaml")]):
        lines = path.read_text(encoding="utf-8").splitlines()
        for i, line in enumerate(lines):
            match = SETUP_UV.match(line)
            if not match:
                continue
            step_indent = len(match.group("indent"))
            pinned: str | None = None
            for following in lines[i + 1 :]:
                if not following.strip() or following.lstrip().startswith("#"):
                    continue
                # Dedent to the step's own level, or a new list item: step over.
                if len(following) - len(following.lstrip()) <= step_indent and NEXT_STEP.match(
                    following
                ):
                    break
                version = VERSION.match(following)
                if version:
                    pinned = version.group("value")
                    break
            found.append((path.name, i + 1, pinned))
    return found


def test_every_setup_uv_step_pins_a_version():
    pins = _setup_uv_pins()
    assert pins, (
        "no setup-uv steps found — this guard has stopped guarding anything. "
        f"Check the layout of {WORKFLOWS}."
    )
    offenders = [
        f"{name}:{line} -> {value or 'omitted (defaults to latest)'}"
        for name, line, value in pins
        if value is None or value == "latest"
    ]
    assert not offenders, (
        "setup-uv must pin an explicit uv version — uv sets the wheel's core "
        "metadata version, and tracking latest broke the v0.3.0 publish:\n  "
        + "\n  ".join(offenders)
    )


def test_all_workflows_pin_the_same_uv():
    versions = {value for _, _, value in _setup_uv_pins() if value}
    assert len(versions) == 1, (
        f"workflows disagree on the uv version {sorted(versions)} — CI would "
        "then validate on a different toolchain from the one that publishes"
    )
