#!/usr/bin/env python
# SPDX-FileCopyrightText: 2026 Opio Inc.
#
# SPDX-License-Identifier: Apache-2.0

"""Generate the Claude plugin + self-hosted marketplace from the canonical registry.

skills/skills.toml is the single source of truth for the skill set, its order,
and the plugin metadata. This generates the Claude Code plugin manifest
(.claude-plugin/plugin.json, which globs ./skills/), the self-hosted marketplace
listing (.claude-plugin/marketplace.json), the Cursor manifest, and the Copilot
agent-plugin marketplace + instructions — all from that one SSOT.

    uv run python scripts/gen_skill_manifests.py           # write
    uv run python scripts/gen_skill_manifests.py --check    # verify (exit 1 on drift)

The AGENTS.md install count/list is verified by tests/skills/test_skill_manifests.py,
not rewritten here.
"""

from __future__ import annotations

import argparse
import json
import sys
import tomllib
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
REGISTRY = REPO_ROOT / "skills" / "skills.toml"


def load_order() -> list[str]:
    """Return the ordered skill list from the canonical registry."""
    with REGISTRY.open("rb") as fh:
        return list(tomllib.load(fh)["order"])


def skill_meta(name: str) -> dict[str, str]:
    """Return a skill's frontmatter (name, description) from its SKILL.md.

    Minimal single-line-value parser — no YAML dependency, so the generator
    stays stdlib-only and runs in any CI venv.
    """
    text = (REPO_ROOT / "skills" / name / "SKILL.md").read_text(encoding="utf-8")
    block = text.split("---", 2)[1]
    meta: dict[str, str] = {}
    for line in block.splitlines():
        key, sep, value = line.partition(":")
        if sep and not key.startswith(" "):
            meta[key.strip()] = value.strip()
    return meta


def load_plugin() -> dict:
    """Return the [plugin] metadata table from the canonical registry."""
    with REGISTRY.open("rb") as fh:
        return dict(tomllib.load(fh)["plugin"])


def _dumps(data: dict) -> str:
    return json.dumps(data, indent=2, ensure_ascii=False) + "\n"


def render_claude_plugin() -> str:
    """The Claude Code plugin manifest (globs ./skills/; self-maintaining)."""
    m = load_plugin()
    return _dumps(
        {
            "name": m["name"],
            "version": m["version"],
            "description": m["description"],
            "author": {"name": m["author_name"], "url": m["author_url"]},
            "homepage": m["homepage"],
            "repository": m["repository"],
            "license": m["license"],
            "keywords": m["keywords"],
            "skills": "./skills/",
        }
    )


def render_marketplace() -> str:
    """Self-hosted marketplace listing this repo as a single plugin (source ./)."""
    m = load_plugin()
    return _dumps(
        {
            "name": m["name"],
            "owner": {"name": m["author_name"]},  # Gate 1: owner takes name/email, not url
            "description": m["description"],  # required by `claude plugin validate --strict`
            "plugins": [
                {
                    "name": m["name"],
                    "source": "./",
                    "description": m["description"],
                    "version": m["version"],
                }
            ],
        }
    )


def render_cursor_plugin() -> str:
    """Thin Cursor manifest pointing at the one canonical ./skills tree."""
    m = load_plugin()
    return _dumps(
        {
            "name": m["name"],
            "version": m["version"],
            "description": m["description"],
            "author": {"name": m["author_name"], "url": m["author_url"]},
            "license": m["license"],
            "skills": "./skills/",  # Gate 3: dir-pointer, no ../ traversal
        }
    )


def render_copilot_marketplace() -> str:
    """Copilot agent-plugin marketplace (.github/plugin/marketplace.json).

    Gate 3: same self-hosted shape as the Claude marketplace.
    """
    return render_marketplace()


def render_copilot_instructions() -> str:
    """Copilot always-loaded instructions that point at the bundled skill(s).

    Rather than dumping the contributor-facing AGENTS.md, this routes Copilot
    to the usage skill(s) — the actual how-to-use guidance — derived from the
    skills registry so it stays in sync as skills are added.
    """
    m = load_plugin()
    lines = [
        "<!-- GENERATED from skills/skills.toml by scripts/gen_skill_manifests.py. "
        "Do not edit. -->",
        "",
        f"# {m['display_name']} — Copilot instructions",
        "",
        m["description"],
        "",
        "When a task involves extracting, querying, auditing, diffing, or verifying "
        "an Excel workbook, load the bundled skill below and follow it.",
        "",
        "## Skills",
    ]
    for name in load_order():
        meta = skill_meta(name)
        lines.append(f"- **{name}** — {meta.get('description', '')}")
        lines.append(f"  See `skills/{name}/SKILL.md`.")
    return "\n".join(lines) + "\n"


ARTIFACTS = {
    REPO_ROOT / ".claude-plugin" / "plugin.json": render_claude_plugin,
    REPO_ROOT / ".claude-plugin" / "marketplace.json": render_marketplace,
    REPO_ROOT / ".cursor-plugin" / "plugin.json": render_cursor_plugin,
    REPO_ROOT / ".github" / "plugin" / "marketplace.json": render_copilot_marketplace,
    REPO_ROOT / ".github" / "copilot-instructions.md": render_copilot_instructions,
}


def main() -> int:
    """Write (or, with --check, verify) the generated plugin artifacts."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check",
        action="store_true",
        help="exit 1 if any generated file is out of date",
    )
    args = parser.parse_args()
    drift: list[Path] = []
    for path, renderer in ARTIFACTS.items():
        expected = renderer()
        if args.check:
            current = path.read_text(encoding="utf-8") if path.exists() else ""
            if current != expected:
                drift.append(path)
        else:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(expected, encoding="utf-8")
    if args.check and drift:
        names = ", ".join(str(p.relative_to(REPO_ROOT)) for p in drift)
        print(
            f"Out-of-date generated file(s): {names}\n"
            f"Run: uv run python scripts/gen_skill_manifests.py",
            file=sys.stderr,
        )
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
