<!--
Title this PR in conventional-commit form, e.g.
  fix(extract): range collapse spills unbounded SQLite temp on rolling-window models
See CONTRIBUTING.md for the full guide.
-->

## What and why

<!-- What changes, and the reason. Link the issue: "Fixes #N." -->

## Scope

- [ ] This PR addresses **one concern**. Unrelated work has been split out.

## Evidence

<!-- Say what you ran, and on which commit. "CI is green" is not evidence for a
     behaviour claim — name the test. -->

- [ ] `uv run pytest -q` passes
- [ ] `uv tool run reuse lint` passes

**If this changes extraction output:**

- [ ] Output equivalence demonstrated (`tests/test_hotspot_characterization.py`, goldens
      captured from the *pre-change* code and still green) — or the changed edges are
      listed below with the reason they changed.
- [ ] `SCHEMA_VERSION` bumped if a documented view changed (renamed/removed view = major).
- [ ] `skills/using-xl-marinade/` and `docs/reference/schema.md` updated to match.

**If this changes the public API or CLI:**

- [ ] `CHANGELOG.md` updated.

## Notes for the reviewer

<!-- Anything you want looked at hardest, and anything you're unsure about. -->
