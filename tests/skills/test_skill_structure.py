# SPDX-FileCopyrightText: 2026 Opio Inc.
#
# SPDX-License-Identifier: Apache-2.0

"""Verify skill files exist and satisfy the Open Agent Skills structural rubric.

Deterministic, dependency-free, runs in milliseconds. Catches missing files,
broken frontmatter, name/dir mismatch, oversize SKILL.md, and reference files
that link onward more than one level deep.
"""

import re
import tomllib
from pathlib import Path

import pytest

SKILLS_DIR = Path(__file__).resolve().parents[2] / "skills"

with (SKILLS_DIR / "skills.toml").open("rb") as _fh:
    EXPECTED_SKILLS = list(tomllib.load(_fh)["order"])

NAME_RE = re.compile(r"^[a-z0-9]+(-[a-z0-9]+)*$")
REF_LINK_RE = re.compile(r"\]\(([^)]+)\)")


def _frontmatter(skill_name: str) -> dict[str, str]:
    """Parse a skill's single-line-value YAML frontmatter (no yaml dependency)."""
    content = (SKILLS_DIR / skill_name / "SKILL.md").read_text(encoding="utf-8")
    block = content.split("---", 2)[1]
    fm: dict[str, str] = {}
    for line in block.splitlines():
        key, sep, value = line.partition(":")
        if sep and not key.startswith(" "):
            fm[key.strip()] = value.strip()
    return fm


@pytest.mark.parametrize("skill_name", EXPECTED_SKILLS)
def test_skill_exists(skill_name: str) -> None:
    """Every registered skill has a SKILL.md file."""
    assert (SKILLS_DIR / skill_name / "SKILL.md").exists()


@pytest.mark.parametrize("skill_name", EXPECTED_SKILLS)
def test_skill_has_valid_frontmatter(skill_name: str) -> None:
    """Every skill has frontmatter with a name and a description."""
    content = (SKILLS_DIR / skill_name / "SKILL.md").read_text(encoding="utf-8")
    assert content.startswith("---"), f"{skill_name} missing frontmatter"
    assert len(content.split("---", 2)) >= 3, f"{skill_name} frontmatter not closed"
    fm = _frontmatter(skill_name)
    assert "name" in fm, f"{skill_name} missing 'name'"
    assert "description" in fm, f"{skill_name} missing 'description'"


@pytest.mark.parametrize("skill_name", EXPECTED_SKILLS)
def test_frontmatter_name_matches_dir(skill_name: str) -> None:
    """Open Agent Skills spec: frontmatter name must equal the parent directory."""
    assert _frontmatter(skill_name)["name"] == skill_name


@pytest.mark.parametrize("skill_name", EXPECTED_SKILLS)
def test_skill_name_is_valid_kebab(skill_name: str) -> None:
    """The frontmatter name is lowercase kebab-case and <= 64 chars."""
    name = _frontmatter(skill_name)["name"]
    assert NAME_RE.match(name), f"{skill_name}: name '{name}' is not kebab-case"
    assert len(name) <= 64, f"{skill_name}: name exceeds 64 chars"


@pytest.mark.parametrize("skill_name", EXPECTED_SKILLS)
def test_skill_description_within_1024(skill_name: str) -> None:
    """The frontmatter description is within Anthropic's 1024-char limit."""
    desc = _frontmatter(skill_name)["description"]
    assert len(desc) <= 1024, f"{skill_name}: description is {len(desc)} chars (max 1024)"


@pytest.mark.parametrize("skill_name", EXPECTED_SKILLS)
def test_skill_md_under_600_lines(skill_name: str) -> None:
    """Anthropic guidance: keep SKILL.md under 600 lines (split into references)."""
    n = len((SKILLS_DIR / skill_name / "SKILL.md").read_text(encoding="utf-8").splitlines())
    assert n <= 600, f"{skill_name} SKILL.md is {n} lines (max 600)"


@pytest.mark.parametrize("skill_name", EXPECTED_SKILLS)
def test_references_one_level_deep(skill_name: str) -> None:
    """A reference file must not link onward to another reference (one level deep)."""
    refs = SKILLS_DIR / skill_name / "references"
    if not refs.is_dir():
        return
    for md in refs.glob("*.md"):
        for target in REF_LINK_RE.findall(md.read_text(encoding="utf-8")):
            assert "references/" not in target, (
                f"{md.relative_to(SKILLS_DIR)} links onward to a reference: {target}"
            )


def test_expected_skills_cover_all_directories() -> None:
    """EXPECTED_SKILLS (from the registry) matches the skill directories on disk."""
    on_disk = {p.parent.name for p in SKILLS_DIR.glob("*/SKILL.md")}
    assert set(EXPECTED_SKILLS) == on_disk, (
        f"registry-only={set(EXPECTED_SKILLS) - on_disk} disk-only={on_disk - set(EXPECTED_SKILLS)}"
    )


def test_name_rule_rejects_bad_names() -> None:
    """The kebab-case rule rejects caps/underscores and namespace prefixes."""
    assert not NAME_RE.match("Using_XL_Marinade")
    assert not NAME_RE.match("my/skill")
    assert NAME_RE.match("using-xl-marinade")
