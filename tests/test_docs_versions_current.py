"""Version claims written into docs prose must match the code.

Three separate instances of this bug shipped in two days: the bundled skill
claimed `schema_version "2.0"` after the 3.0 bump, `docs/reference/schema.md`
claimed package `0.1.0` after the 0.2.0 bump, and the same file then hardcoded
`0.2.0`. Prose has no compiler, so nothing failed when it drifted.

Two versions, two different treatments:

- The **package version** is not named in the docs at all. It cannot be kept
  current, and a reader who needs it can ask the installed distribution. The
  test enforces its absence, so it cannot creep back in.
- The **output-schema version** is named, because a reader genuinely needs to
  know which contract a database speaks. It is pinned to `SCHEMA_VERSION`, and
  the test asserts it matched something before checking it — rewording the prose
  fails loudly rather than silently reducing the test to a no-op. A vacuous
  version test is worse than none, because it reads as coverage.
"""

import importlib.metadata
import re
from pathlib import Path

DOCS = Path(__file__).resolve().parent.parent / "docs"

# The docs deliberately name no package version — it cannot be kept current, and
# a reader who needs it can ask the installed distribution. So this is a negative
# assertion: no three-part version may sit near the package name. Scoped by
# proximity so an unrelated version (a Python release, another tool) is fine,
# while `pip install xl-marinade==0.3.0` or "__version__ is 0.3.0" is not.
SEMVER = re.compile(r"[0-9]+\.[0-9]+\.[0-9]+")
PACKAGE_TERMS = ("xl-marinade", "xl_marinade", "package version")
PACKAGE_PROXIMITY = 80

# A quoted two-part version in backticks — `"3.0"`.
QUOTED_VERSION = re.compile(r'`"([0-9]+\.[0-9]+)"`')

# ...but only counts as an output-schema claim when `schema_version` appears in
# the run-up. docs/guide/diffing.md documents a *changelist format* version
# ("currently `"1.0"`, not a workbook version") which is a different number
# entirely; anchoring on the surrounding term keeps it out.
SCHEMA_CONTEXT = "schema_version"
SCHEMA_LOOKBACK = 250


def _markdown_sources() -> list[tuple[Path, str]]:
    return [(p, p.read_text(encoding="utf-8")) for p in sorted(DOCS.rglob("*.md"))]


def test_docs_never_hardcode_the_package_version():
    offenders: list[str] = []
    for path, text in _markdown_sources():
        for match in SEMVER.finditer(text):
            window = text[
                max(0, match.start() - PACKAGE_PROXIMITY) : match.end() + PACKAGE_PROXIMITY
            ]
            if any(term in window for term in PACKAGE_TERMS):
                line = text.count("\n", 0, match.start()) + 1
                offenders.append(f"{path.name}:{line} -> {match.group(0)}")

    assert not offenders, (
        "docs name a package version, which cannot be kept current — installed is "
        f"{importlib.metadata.version('xl-marinade')} today and will not be tomorrow: "
        f"{offenders}. Point the reader at importlib.metadata.version('xl-marinade') "
        "instead, or drop the version entirely."
    )


def test_docs_schema_version_claims_match_the_pipeline_constant():
    from xl_marinade.core.new_arch.fast_extraction_pipeline import SCHEMA_VERSION

    found: list[tuple[Path, str]] = []
    for path, text in _markdown_sources():
        for match in QUOTED_VERSION.finditer(text):
            run_up = text[max(0, match.start() - SCHEMA_LOOKBACK) : match.start()]
            if SCHEMA_CONTEXT in run_up:
                found.append((path, match.group(1)))

    assert found, (
        "no schema-version claim found in docs/ — the prose was reworded and this "
        f"test silently stopped checking. Update SCHEMA_CONTEXT in {__file__}."
    )
    wrong = [(str(p), v) for p, v in found if v != SCHEMA_VERSION]
    assert not wrong, (
        f"docs claim a stale output-schema version (pipeline says {SCHEMA_VERSION}): {wrong}"
    )
