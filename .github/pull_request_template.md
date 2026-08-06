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

## Changelog

- [ ] User-visible change (behaviour, API, CLI, output schema, performance) → an
      entry under `## [Unreleased]`, or under the pending release section if one is
      open. Internal refactors and repo hygiene don't need one.
- [ ] **If anything merged to `main` while this PR was open, re-check that your entry
      is still there.** Concurrent PRs edit the same block, so the later merge
      silently wins and no test catches it — this is how #8's speedup disappeared
      from the 0.2.0 notes.

## Notes for the reviewer

<!-- Anything you want looked at hardest, and anything you're unsure about. -->
