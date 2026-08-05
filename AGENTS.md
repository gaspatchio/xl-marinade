# XL Marinade — shared agent rules

## Project Overview

XL Marinade turns an Excel workbook into an auditable **formula graph** in SQLite: every
cell's formula, its cross-sheet dependencies, canonical formula families, and (optionally)
VBA edges. It is **two-tier**:

- a **free, deterministic** core + documentation generator (network-free, no AI), and
- an **optional bring-your-own-key (BYOK) LLM** layer that enriches labels/classification.

The deterministic SQLite output is the product's stable, SemVer'd public contract.

## AI Plugin Installation

XL Marinade ships as a plugin for AI coding agents — install once to get the
`using-xl-marinade` skill (how to drive `marinade extract` / `diff` / `document`
and query the SQLite IR) plus always-loaded usage instructions. Every editor
manifest is generated from one `skills/` tree by
`scripts/gen_skill_manifests.py` (SSOT: `skills/skills.toml`).

### Claude Code
```
/plugin marketplace add gaspatchio/xl-marinade
/plugin install xl-marinade@xl-marinade
```

### VS Code / GitHub Copilot
Add to user settings:
```json
{ "chat.plugins.marketplaces": ["gaspatchio/xl-marinade"] }
```
Requires `"chat.plugins.enabled": true` (Agent Plugins is in preview).

### Cursor
Open the project — the `.cursor-plugin/` directory is auto-detected.

### Any Agent (universal)
```bash
npx skills add gaspatchio/xl-marinade
```

### What You Get
- **1 skill**: using-xl-marinade
- **Always-loaded usage instructions** (Copilot) generated from the skill set

## Architecture — packages and the tier boundary

| Package | Role | Ships in | Network |
|---|---|---|---|
| `xl_marinade.core` | Extraction → SQLite (`extract()`), IR diff, labelling primitives | base install (Tier-0) | **never** |
| `xl_marinade.docs` | Deterministic documentation (`document()` → `documentation.md` + `model_spec.json`) | base install (**free**) | **never** |
| `xl_marinade.llm` | BYOK enrichment (`document(..., provider=None)`, sprint7 pipeline) | `xl-marinade[llm]` extra | only on explicit opt-in |
| `xl_marinade.cli` | Thin Typer CLI, command `marinade` (`extract` / `document` / `diff`) | base install | — |

**The tier boundary is a hard, one-way invariant: `llm → docs → core`, never the reverse.**
- `xl_marinade.docs` imports **no** `openai` and **no** `xl_marinade.llm`; `import xl_marinade.docs` must succeed with the `[llm]` extra absent.
- `xl_marinade.core` stays Tier-0: network-free, no `llm`/`openai` (the one lazy, `[llm]`-guarded VBA-enrichment seam is the sole exception and is opt-in).
- `import xl_marinade.llm` raises an actionable "install `xl-marinade[llm]`" message when `openai` is absent.

These invariants are enforced by tests (`tests/test_docs_no_llm.py`, `tests/test_core_no_llm_dep.py`, `tests/test_core_no_egress.py`, `tests/test_import_cleanliness.py`) — do not weaken them.

## BYOK / egress rules

- No network call without **explicit opt-in** (a provider/key) **and** first-run egress disclosure.
- The LLM client is configured in exactly one seam — `xl_marinade.llm.factory.make_llm_client()`. Env: `LLM_PROVIDER` (`openai` | `azure` | `openai_compatible`), `LLM_API_KEY` (falls back to `OPENAI_API_KEY`), `LLM_BASE_URL` (Azure / vLLM / Ollama / LiteLLM). Product-specific env vars use the `MARINADE_` prefix (e.g. `MARINADE_ONTOLOGY_PATH`).
- **No key ⇒ deterministic-only, never raise.** Every LLM entry point probes availability and falls back to the deterministic path.

## Development commands

```bash
uv sync                                        # install (add [llm] for the BYOK extra)
uv run pytest -q                               # the test suite is the gate
uv run pytest -v                               # verbose
uv run ruff format . && uv run ruff check .     # format + lint (curated select; see pyproject)
uv run mypy src/xl_marinade && uv run pyright src/xl_marinade   # both type checkers must pass
marinade extract book.xlsx -o ir.db             # CLI: deterministic extraction
marinade document ir.db -o out/                 # CLI: deterministic docs (add [llm] to enrich)
```

The system Python has no dependencies — always `uv run`. The package is imported from `src/` via `pyproject.toml`'s `pythonpath`, so a plain `uv run pytest` works from a clean checkout.

## Python standards

Mirror the sibling **gaspatchio** Python rules:

- **Ruff**: a curated lint set (`select = ["E","F","W","I","UP","B","A"]` — see `pyproject.toml`); code must lint clean except the listed ignores, and must already be `ruff format`-clean.
- **Every function, method, and class has explicit type hints** — no implicit `Any`, a non-`Any` return type, and `| None` over `Optional[T]`. (This is a hard rule: an untyped public parameter is a review defect.)
- **Never `print()` in production code** — use the `loguru` logger. (The only `print` allowed is inside a CLI `__main__`/Typer command that is the user-facing output.)
- **f-strings** for interpolation; **`pathlib.Path`** over `os.path`; **`with`** context managers for resources.
- `@dataclass(slots=True)` or `pydantic.BaseModel` for structured data — not ad-hoc dicts. Prefer the **RORO** pattern (receive an object, return an object) for complex I/O.
- Handle errors/edge cases first with **early returns**; keep the happy path last; raise **typed exceptions** (the `errors.py` hierarchy) that the CLI maps to exit codes. Avoid bare `except`.
- `snake_case` for files and calculated names; lowercase-with-underscores directories.
- **Linter strips unused imports on save** — add a new import **and its first use in the same edit**, or the on-save `ruff` removes it before you use it.

## Design documents

Specs and implementation plans live in `ref/<topic>/` (specs under `specs/`, plans under `plans/`), following the numbered-prefix convention.

## Commit conventions

- **Sign your commits** (SSH or GPG; `git config commit.gpgsign true`).
- **Conventional commits** (`feat:`, `fix:`, `docs:`, `chore:`, `refactor:`, `test:`, `build:`); explain the *why*. Keep commits focused and atomic. Reference issue numbers where applicable.
- **Never** add an AI-assistant signature or `Co-Authored-By: <assistant>` trailer.
- Authorship is canonicalized via `.mailmap` (Klaas Stijnen as the extractor's original author); the packaging/refactor commits authored as Matt Wright stay as-is.

## GitHub workflow

**[`CONTRIBUTING.md`](CONTRIBUTING.md) is canonical** — branches vs forks, signing setup, commit format, issue-title conventions, PR scope, and the merge policy all live there and apply to you exactly as they apply to a human. Read it before opening an issue or a PR. What follows is only what differs when there is no human in the loop.

- **`main` is protected by ruleset** (signed commits, no force-push, no deletion, no bypass actors). These are server-side rules — a push that violates one is refused, not warned about. Never commit to `main`; branch, then open a PR.
- **Verify claims; do not relay them.** "Byte-identical output" and "all tests pass" are hypotheses until you have run them yourself. State what you ran, and on which SHA. Reporting a PR author's claim as though you had checked it is the failure mode that matters here.
- **Bind every test claim to a SHA.** Re-check `headRefOid` before asserting results — a branch can move between your fetch and your review, and a stale "all green" misleads the author more than saying nothing would.
- **Anchor findings to lines** via a formal review with inline comments, not a wall of prose. Separate blockers from risks from nits and say which is which; reserve `--request-changes` for genuine ship-blockers.
- **Never self-apply `confirmed` or `pending-release`** — those record a maintainer's judgement, not yours.
- **Check the bundled skill and docs against schema changes.** `skills/using-xl-marinade/` ships queries against named views; a rename that misses them breaks every agent consuming the tool — including you.
