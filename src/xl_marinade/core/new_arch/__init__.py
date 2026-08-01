# ABOUTME: New memory-efficient extraction architecture namespace
# ABOUTME: Contains streaming XML parser, PRAGMA policy, and related components

"""
Memory-efficient IR extraction architecture.

This module provides the core components for streaming Excel workbook extraction
with bounded memory usage and deterministic output.

Key modules:
- sqlite_pragmas: PRAGMA policy enforcement for bounded memory
- fast_parser: XML streaming parser for Excel workbooks
- cell_identity: Integer-based cell identity encoding
- workbook_catalog: Sheet catalog and metadata extraction
"""

from xl_marinade.core.new_arch.sqlite_pragmas import (
    MANDATORY_PRAGMAS,
    apply_pragmas,
    get_pragma_config,
    verify_pragmas,
)

__all__ = [
    "apply_pragmas",
    "verify_pragmas",
    "get_pragma_config",
    "MANDATORY_PRAGMAS",
]
