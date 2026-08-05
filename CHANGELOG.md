# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres
to [Semantic Versioning](https://semver.org/spec/v2.0.0.html). The SQLite output
schema is a versioned public contract.

## [Unreleased]

### Changed
- **Breaking (output schema 2.0 → 3.0):** the unified node view `atlas_nodes`
  is renamed to `marinade_nodes`. Same columns and semantics; only the view
  name changes. Databases stamped `schema_version` 3.0 no longer contain an
  `atlas_nodes` view.

## [0.1.0] - 2026-08-02

### Added
- Deterministic Excel formula-graph extraction to SQLite (`marinade extract`).
- Deterministic documentation generation (`marinade document`); optional
  bring-your-own-key LLM enrichment via the `xl-marinade[llm]` extra.
- IR diff (`marinade diff`) emitting a JSON changelist.
- Importable library API: `xl_marinade.extract`, `xl_marinade.diff`,
  `xl_marinade.docs.document`, `xl_marinade.llm.document`.
