# Contributing to XL Marinade

Thanks for helping. This document is the canonical contribution guide — it applies to
humans and to AI coding agents equally. Agents should also read [`AGENTS.md`](AGENTS.md),
which adds rules specific to working without a human in the loop.

## Reporting a bug

Open an issue using the bug template. The one rule that matters most:

**Never attach a client workbook.** If the problem only reproduces on a confidential
model, email the report to **security@opioinc.com** instead — that channel accepts
ordinary bugs, not just vulnerabilities, and we open an anonymised public issue with a
synthetic reproduction. See [`SECURITY.md`](SECURITY.md).

### Writing a good issue title

The title names **the surface, the wrong behaviour, and the mechanism**. Long and specific
beats short and vague — the title is what a future reader greps.

    extract: "database or disk is full" on rolling-window models (unbounded
      SQLite temp spill in range collapse)
    extract: lookup-dense sheets (INDEX/MATCH ledgers) run ~3-30x slower than
      sibling sheets

- **Lead with what is *silently* wrong.** Wrong edges that look right are the worst failure
  mode this tool has — the output is a provenance claim, so a plausible-but-wrong graph is
  worse than a crash. Say "silently" in the title when it applies.
- **Carry the contrast case** when one exists ("this sheet vs its sibling"). It is the
  fastest route to the mechanism.
- Include a reproduction. A generator snippet (openpyxl, or `test_workbook_generator/`)
  is ideal — it is shareable where a real workbook is not.

New issues land `needs-triage`. `confirmed`, `needs-repro` and `pending-release` record a
maintainer's judgement; please don't self-apply them.

## Setting up

```bash
uv sync --locked        # exact pinned versions, matching CI
uv run pytest -q        # full suite
uv tool run reuse lint  # licence headers (CI enforces this)
```

## Branch or fork?

Both work. Which you use depends on whether you have commit access:

- **Commit access** → push a branch to this repository and open the PR from it. Simpler,
  and CI behaves identically to `main`.
- **No commit access** → fork, push to your fork, open the PR from there. This is the
  standard GitHub flow for outside contributors and nothing about it is second-class.

If you find yourself contributing regularly from a fork, ask for commit access — once
trust is established a fork adds friction and buys nothing.

**On CI:** fork-based PRs cannot read repository secrets. This repo's CI needs nothing
beyond the automatic `GITHUB_TOKEN`, so the full suite runs on fork PRs exactly as it does
on branches. Leave **"Allow edits by maintainers"** ticked so a maintainer can push a
rebase or a small fix rather than bouncing the PR back to you.

## Commits

- **Sign your commits.** `main` is protected by a ruleset requiring signed commits.
  Register your key on GitHub **as a signing key** — that is a separate entry from an auth
  key, even when the key material is identical:

  ```bash
  ssh-keygen -t ed25519 -C "signing" -f ~/.ssh/id_ed25519_signing
  git config --global gpg.format ssh
  git config --global user.signingkey ~/.ssh/id_ed25519_signing.pub
  git config --global commit.gpgsign true
  ```

  Set this up **before** you start a branch. `required_signatures` is evaluated against
  every commit in the PR, and **a squash merge does not sidestep it** — GitHub refuses the
  merge before it creates the squash commit. Unsigned commits therefore have to be
  re-signed, which means a rebase. If you're already stuck with some, say so on the PR and
  a maintainer will rebase it for you (your authorship is preserved); it isn't something
  you need to untangle alone.

- **Conventional commits** (`feat:`, `fix:`, `docs:`, `chore:`, `refactor:`, `test:`,
  `build:`) with a scope where one applies. Explain the *why*, not just the *what*. Keep
  commits focused and atomic.
- **Never** add an AI-assistant signature or a `Co-Authored-By: <assistant>` trailer.
- Authorship is canonicalised via `.mailmap`.

## Pull requests

- Title in conventional-commit form, matching the commit style.
- Body closes its issues explicitly: `Fixes #N.`
- **One PR, one concern.** Don't append unrelated commits to an open PR. Performance work
  reviewed as a footnote to a docs sync is how correctness regressions get through — split
  it out, and re-request review when the head moves.
- **A change to extraction output needs equivalence evidence, not an assertion.**
  `tests/test_hotspot_characterization.py` is the mechanism: capture a golden from the
  *pre-change* code, then prove it stays green afterwards. That is a different claim from
  "the suite passes" — extend `tests/golden/` when a change touches a surface it doesn't
  yet pin.
- **Bump `SCHEMA_VERSION` when a documented view changes.** A renamed or removed view is a
  major bump, and the bundled skill in `skills/using-xl-marinade/` ships queries against
  those view names — update it in the same PR.

`main` rejects force-pushes and deletions, and we **squash merge**.

## Reviewing

Reviewers, and anyone commenting on someone else's PR:

- **Verify claims; don't relay them.** "Byte-identical output" and "all tests pass" are
  hypotheses until you have run them. State what you ran, and on which SHA.
- **Bind test claims to a SHA.** A branch can move between your fetch and your review; a
  stale "all green" misleads the author more than saying nothing would.
- **Anchor findings to lines** with inline comments rather than a wall of prose. Reserve
  "request changes" for genuine ship-blockers.
- **Separate blockers from risks from nits, and say which is which.** A nit presented at
  the same volume as a blocker costs the author time.

## Code of conduct

Participation is governed by [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md).
