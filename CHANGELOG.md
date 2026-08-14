# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres
to [Semantic Versioning](https://semver.org/spec/v2.0.0.html). The SQLite output
schema is a versioned public contract.

## [Unreleased]

### Added
- `marinade --version` (and `-V`) report the installed version. Previously the
  only way to learn it was package metadata.

### Changed
- Extraction telemetry is written to `<database>.telemetry.json` beside the
  output rather than a fixed `telemetry.json`. The old fixed name silently
  overwrote any same-named file in the chosen directory, and two extractions
  sharing an output directory clobbered each other's telemetry.

### Fixed
- `marinade extract -o <dir>/<file>` now creates a missing output directory
  instead of failing with SQLite's bare `unable to open database file`, which
  named neither the cause nor the path and arrived after the CLI had already
  printed the output path as though it were fine. If the directory cannot be
  created, the error now names it.
- A database path containing `#`, `%HH` or `'` no longer opens the wrong file.
  Read-only opens built a SQLite URI by interpolating the path into
  `file:{path}?mode=ro`, so SQLite parsed filename characters as URI syntax: a
  `#` truncated the path at a fragment marker and took `?mode=ro` with it,
  silently opening — and creating — a different, empty database with no
  read-only guarantee, which surfaced as "no bindings" rather than an error.
  Paths now go through `Path.as_uri()`, and ATTACH binds the filename as a
  parameter.
- Generated text artifacts are now byte-identical across platforms.
  `documentation.md`, `model_spec.json`, the `marinade diff` changelist,
  `telemetry.json` and the mutation/usage logs were written in text mode
  without `newline=`, so Python translated every `\n` to `\r\n` on Windows —
  spurious diffs for anything hashing or comparing two runs, against a tool
  whose stated contract is determinism. All text writes now pin `newline="\n"`.
- `apply_actuarial_classification` no longer reports failure for a run that
  succeeded. Its completion banner was the package's only non-ASCII write to
  stdout, which Windows encodes with the ANSI code page under `errors='strict'`
  when redirected — and the banner fires after the overlay database is
  committed, so the `UnicodeEncodeError` turned a finished run into a non-zero
  exit. The banner is now ASCII and goes to stderr.
- A failed or interrupted `marinade extract` no longer destroys the previous
  output database. The final VACUUM wrote to the target after unlinking it, so
  a full disk or a Ctrl-C left the user with neither the old database nor a new
  one; the write now goes to a sibling temp file moved into place atomically.
  Note the tradeoff: the previous database is retained until the new one is
  complete, so extraction now needs room for both at once and fails cleanly on
  a nearly-full disk where it previously succeeded destructively.
- `marinade extract` crashed on ANY workbook under a non-UTF-8 locale
  (`'charmap' codec can't decode byte 0x90` on Windows): text files were
  read and written with the platform's default encoding. All text I/O now
  passes `encoding="utf-8"` explicitly, enforced by ruff's PLW1514.

## [0.2.0] - 2026-08-06

### Changed
- **Breaking (output schema 2.0 → 3.0):** the unified node view `atlas_nodes`
  is renamed to `marinade_nodes`. Same columns and semantics; only the view
  name changes. Databases stamped `schema_version` 3.0 no longer contain an
  `atlas_nodes` view. The bundled `using-xl-marinade` skill queries the new
  name.

### Fixed
- Extraction of rolling-window models no longer spills unbounded SQLite temp
  space ("database or disk is full") in the range-collapse step; replaced by
  chunked breadth aggregation + disjoint-box union counting with identical
  output.
- Lookup-dense sheets (INDEX/MATCH ledgers) no longer re-scan the lookup array
  once per formula, and no longer pay a `runtime_checkable` Protocol check per
  cell read. Bounded-range MATCH scans are memoized per immutable snapshot —
  never from a live `Workbook` source, whose mutations must stay visible — and
  the value source is classified once at construction. A 15k-formula
  INDEX/MATCH probe goes 111s → 5s and a ledger-style benchmark 231s → 17s,
  with byte-identical extraction verified on 8 real workbooks. The memo costs
  ~194 MB peak RSS on a 2.3M-formula model.

## [0.1.0] - 2026-08-02

### Added
- Deterministic Excel formula-graph extraction to SQLite (`marinade extract`).
- Deterministic documentation generation (`marinade document`); optional
  bring-your-own-key LLM enrichment via the `xl-marinade[llm]` extra.
- IR diff (`marinade diff`) emitting a JSON changelist.
- Importable library API: `xl_marinade.extract`, `xl_marinade.diff`,
  `xl_marinade.docs.document`, `xl_marinade.llm.document`.
