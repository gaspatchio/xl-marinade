"""Version claims written into docs prose must match the code.

Three separate instances of this bug shipped in two days: the bundled skill
claimed `schema_version "2.0"` after the 3.0 bump, `docs/reference/schema.md`
claimed package `0.1.0` after the 0.2.0 bump, and the same file then hardcoded
`0.2.0`. Prose has no compiler, so nothing failed when it drifted.

These tests pin the two claims that exist today. Each asserts it found at least
one claim before checking it, so rewording the prose fails loudly rather than
silently reducing the test to a no-op — a vacuous version test is worse than no
version test, because it reads as coverage.
"""

import importlib.metadata
import re
from pathlib import Path

DOCS = Path(__file__).resolve().parent.parent / "docs"

# "`xl_marinade.__version__` (currently\n`0.2.0`)" — the line break is real, the
# prose is wrapped, so DOTALL and a bounded gap rather than a same-line match.
PACKAGE_CLAIM = re.compile(r"__version__`\s*\(currently\s*`?\s*`([0-9]+\.[0-9]+\.[0-9]+)`")

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


def test_docs_package_version_claims_match_the_installed_distribution():
    expected = importlib.metadata.version("xl-marinade")
    found: list[tuple[Path, str]] = []
    for path, text in _markdown_sources():
        found += [(path, m) for m in PACKAGE_CLAIM.findall(text)]

    assert found, (
        "no package-version claim found in docs/ — the prose was reworded and this "
        f"test silently stopped checking. Update PACKAGE_CLAIM in {__file__}."
    )
    wrong = [(str(p), v) for p, v in found if v != expected]
    assert not wrong, f"docs claim a stale package version (installed {expected}): {wrong}"


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
