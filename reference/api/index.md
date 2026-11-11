# API reference

The public Python API. All functions raise the typed `xl_marinade.errors` hierarchy rather than leaking internal exceptions.

## Top-level

## `xl_marinade`

XL Marinade — deterministic Excel formula-graph extraction.

The public API is re-exported here. Importing this package has no side effects and makes no network calls (see tests/test_import_cleanliness.py).

### `diff(db_a, db_b)`

Compare two IR databases and return a replay-complete changelist.

Deterministic and network-free. Thin typed wrapper over :func:`xl_marinade.core.ir_diff.pipeline.diff_ir` so callers (and the CLI) see a typed error instead of a raw `sqlite3.Error` when an input database is unreadable, corrupt, or not an IR database.

Raises:

| Type            | Description                                                                                                          |
| --------------- | -------------------------------------------------------------------------------------------------------------------- |
| `DiffError`     | if the diff fails for an unexpected reason (e.g. an unreadable or corrupt IR database).                              |
| `MarinadeError` | for typed diff failures such as xl_marinade.core.ir_diff.model.DiffVerificationError, which are re-raised unchanged. |

### `extract(workbook, out, *, max_memory_mb=1800, enrich=False)`

Extract an Excel workbook's formula graph to a SQLite database.

Deterministic and network-free by default. Returns the path to the written database.

Parameters:

| Name     | Type   | Description                                                                                                                                                                       | Default |
| -------- | ------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------- |
| `enrich` | `bool` | opt-in LLM VBA enrichment (makes network calls; requires the xl-marinade[llm] extra). Defaults to False — a bare install never contacts a network, even if OPENAI_API_KEY is set. | `False` |

Raises:

| Type              | Description          |
| ----------------- | -------------------- |
| `ExtractionError` | if extraction fails. |

## Extraction & diff

## `xl_marinade.core.api`

Public library API for XL Marinade.

The heavy extraction pipeline is imported lazily inside :func:`extract` so that `import xl_marinade` stays cheap and side-effect-free.

### `diff(db_a, db_b)`

Compare two IR databases and return a replay-complete changelist.

Deterministic and network-free. Thin typed wrapper over :func:`xl_marinade.core.ir_diff.pipeline.diff_ir` so callers (and the CLI) see a typed error instead of a raw `sqlite3.Error` when an input database is unreadable, corrupt, or not an IR database.

Raises:

| Type            | Description                                                                                                          |
| --------------- | -------------------------------------------------------------------------------------------------------------------- |
| `DiffError`     | if the diff fails for an unexpected reason (e.g. an unreadable or corrupt IR database).                              |
| `MarinadeError` | for typed diff failures such as xl_marinade.core.ir_diff.model.DiffVerificationError, which are re-raised unchanged. |

### `extract(workbook, out, *, max_memory_mb=1800, enrich=False)`

Extract an Excel workbook's formula graph to a SQLite database.

Deterministic and network-free by default. Returns the path to the written database.

Parameters:

| Name     | Type   | Description                                                                                                                                                                       | Default |
| -------- | ------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------- |
| `enrich` | `bool` | opt-in LLM VBA enrichment (makes network calls; requires the xl-marinade[llm] extra). Defaults to False — a bare install never contacts a network, even if OPENAI_API_KEY is set. | `False` |

Raises:

| Type              | Description          |
| ----------------- | -------------------- |
| `ExtractionError` | if extraction fails. |

## Deterministic documentation

## `xl_marinade.docs`

xl_marinade.docs — deterministic documentation generation (free tier).

Network-free; imports no `openai`. See :func:`xl_marinade.docs.pipeline.document`.

### `document(ir_db, out_dir)`

Generate deterministic documentation for an extracted IR database.

Writes `documentation.md` and `model_spec.json` under `out_dir`. Returns the path to `documentation.md`. No network, no LLM.

## Errors

## `xl_marinade.errors`

Typed error hierarchy for XL Marinade.

Library code raises these; the CLI (`xl_marinade.cli`) is the only layer that maps them to process exit codes.

### `DiffError`

Bases: `MarinadeError`

IR diff failed for a reason not covered by a more specific error.

Wraps unexpected failures from the diff pipeline (e.g. an unreadable or corrupt IR database) so callers see a typed error instead of a raw `sqlite3.Error`.

### `ExtractionError`

Bases: `MarinadeError`

Extraction failed for a reason not covered by a more specific error.

### `LLMUnavailable`

Bases: `MarinadeError`

An LLM feature was requested but no provider is configured/installed.

### `MarinadeError`

Bases: `Exception`

Base class for all XL Marinade errors.

### `MemoryBudgetExceeded`

Bases: `MarinadeError`

Extraction exceeded the configured memory budget.

### `UnsupportedInput`

Bases: `MarinadeError`

The input is not a supported or readable Excel workbook.
