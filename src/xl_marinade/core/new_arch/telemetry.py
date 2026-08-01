# ABOUTME: Telemetry collection for memory-efficient extraction pipeline
# ABOUTME: Records performance metrics (wall_time_s, peak_rss_mb, row_count, parse_time_ms) for SLA validation

"""
Telemetry Collection

Records performance metrics during extraction for SLA validation and debugging.

Design reference: §11.2 of memory_efficient_extraction_architecture.md
"""

import time
from dataclasses import dataclass


@dataclass
class TelemetryData:
    """Telemetry data collected during extraction."""

    # Timing metrics
    wall_time_s: float = 0.0
    parse_time_ms: float = 0.0
    formula_parse_time_ms: float = 0.0
    reference_extract_time_ms: float = 0.0
    dedupe_time_ms: float = 0.0
    finalize_time_ms: float = 0.0
    vacuum_time_ms: float = 0.0

    # Memory metrics
    peak_rss_mb: float = 0.0

    # Row counts
    row_count: int = 0
    raw_cells_count: int = 0
    raw_formulas_count: int = 0
    raw_edges_internal_count: int = 0
    raw_edges_range_count: int = 0
    raw_edges_external_count: int = 0

    final_cells_count: int = 0
    final_formulas_count: int = 0
    final_edges_internal_count: int = 0
    final_edges_range_count: int = 0
    final_edges_external_count: int = 0

    # Metadata
    sqlite_version: str = ""
    schema_version: str = ""
    build_mode: str = ""
    extractor_git_sha: str = ""
    workbook_sha256: str = ""
    ir_db_path: str = ""

    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            "wall_time_s": round(self.wall_time_s, 3),
            "parse_time_ms": round(self.parse_time_ms, 3),
            "formula_parse_time_ms": round(self.formula_parse_time_ms, 3),
            "reference_extract_time_ms": round(self.reference_extract_time_ms, 3),
            "dedupe_time_ms": round(self.dedupe_time_ms, 3),
            "finalize_time_ms": round(self.finalize_time_ms, 3),
            "vacuum_time_ms": round(self.vacuum_time_ms, 3),
            "peak_rss_mb": round(self.peak_rss_mb, 2),
            "row_count": self.row_count,
            "raw_cells_count": self.raw_cells_count,
            "raw_formulas_count": self.raw_formulas_count,
            "raw_edges_internal_count": self.raw_edges_internal_count,
            "raw_edges_range_count": self.raw_edges_range_count,
            "raw_edges_external_count": self.raw_edges_external_count,
            "final_cells_count": self.final_cells_count,
            "final_formulas_count": self.final_formulas_count,
            "final_edges_internal_count": self.final_edges_internal_count,
            "final_edges_range_count": self.final_edges_range_count,
            "final_edges_external_count": self.final_edges_external_count,
            "sqlite_version": self.sqlite_version,
            "schema_version": self.schema_version,
            "build_mode": self.build_mode,
            "extractor_git_sha": self.extractor_git_sha,
            "workbook_sha256": self.workbook_sha256,
            "ir_db_path": self.ir_db_path,
        }


class TelemetryCollector:
    """
    Telemetry collector for extraction pipeline.

    Tracks timing, memory, and row counts during extraction.
    """

    def __init__(self):
        """Initialize telemetry collector."""
        self.data = TelemetryData()
        self._start_time: float | None = None
        self._phase_start_time: float | None = None

    def start(self) -> None:
        """Start overall timing."""
        self._start_time = time.time()

    def stop(self) -> None:
        """Stop overall timing."""
        if self._start_time is not None:
            self.data.wall_time_s = time.time() - self._start_time

    def start_phase(self) -> None:
        """Start timing for a phase."""
        self._phase_start_time = time.time()

    def stop_phase(self, phase_name: str) -> float:
        """
        Stop timing for a phase and record elapsed time.

        Args:
            phase_name: Name of the phase (parse, formula_parse, etc.)

        Returns:
            Elapsed time in milliseconds
        """
        if self._phase_start_time is None:
            return 0.0

        elapsed_ms = (time.time() - self._phase_start_time) * 1000

        # Record in appropriate field
        if phase_name == "parse":
            self.data.parse_time_ms = elapsed_ms
        elif phase_name == "formula_parse":
            self.data.formula_parse_time_ms = elapsed_ms
        elif phase_name == "reference_extract":
            self.data.reference_extract_time_ms = elapsed_ms
        elif phase_name == "dedupe":
            self.data.dedupe_time_ms = elapsed_ms
        elif phase_name == "finalize":
            self.data.finalize_time_ms = elapsed_ms
        elif phase_name == "vacuum":
            self.data.vacuum_time_ms = elapsed_ms

        self._phase_start_time = None
        return elapsed_ms

    def record_memory(self, peak_rss_mb: float) -> None:
        """Record peak RSS."""
        self.data.peak_rss_mb = peak_rss_mb

    def record_row_count(self, count: int) -> None:
        """Record total row count."""
        self.data.row_count = count

    def record_raw_counts(
        self, cells: int, formulas: int, edges_internal: int, edges_range: int, edges_external: int
    ) -> None:
        """Record raw table counts."""
        self.data.raw_cells_count = cells
        self.data.raw_formulas_count = formulas
        self.data.raw_edges_internal_count = edges_internal
        self.data.raw_edges_range_count = edges_range
        self.data.raw_edges_external_count = edges_external

    def record_final_counts(
        self, cells: int, formulas: int, edges_internal: int, edges_range: int, edges_external: int
    ) -> None:
        """Record final table counts."""
        self.data.final_cells_count = cells
        self.data.final_formulas_count = formulas
        self.data.final_edges_internal_count = edges_internal
        self.data.final_edges_range_count = edges_range
        self.data.final_edges_external_count = edges_external

    def record_metadata(
        self,
        sqlite_version: str,
        schema_version: str,
        build_mode: str,
        extractor_git_sha: str,
        workbook_sha256: str,
        ir_db_path: str,
    ) -> None:
        """Record metadata."""
        self.data.sqlite_version = sqlite_version
        self.data.schema_version = schema_version
        self.data.build_mode = build_mode
        self.data.extractor_git_sha = extractor_git_sha
        self.data.workbook_sha256 = workbook_sha256
        self.data.ir_db_path = ir_db_path

    def get_data(self) -> TelemetryData:
        """Get telemetry data."""
        return self.data

    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return self.data.to_dict()
