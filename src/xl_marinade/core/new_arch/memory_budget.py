# ABOUTME: RSS-based memory budget controller with configurable warn/abort thresholds
# ABOUTME: Monitors process RSS using psutil and enforces memory limits during extraction

"""
Memory Budget Controller

Monitors process RSS (Resident Set Size) using psutil and enforces memory limits.
Implements warn-at-80% and abort-on-breach policy.

Design reference: §8 Phase 8, §11.2, §13 of memory_efficient_extraction_architecture.md
"""

import sys
from dataclasses import dataclass

from xl_marinade.errors import MemoryBudgetExceeded

try:
    import psutil

    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False


# Exit codes (reserved by sub-sprint plan)
EXIT_SLA_FAILURE = 42


@dataclass
class MemoryBudgetConfig:
    """Configuration for memory budget controller."""

    max_memory_mb: int = (
        3000  # Raised from 1800 to handle large workbooks (a large workbook: 770k rows)
    )
    warn_threshold_pct: float = 0.8  # Warn at 80% of max
    check_interval_rows: int = 10_000  # Check every N rows (design §8)
    enabled: bool = True  # Can be disabled for testing


class MemoryBudgetController:
    """
    Memory budget controller using psutil for OS-level RSS monitoring.

    Implements:
    - Periodic RSS sampling (every N rows)
    - Warn at 80% threshold
    - Abort at 100% threshold with exit code 42
    - Peak RSS tracking for telemetry
    """

    def __init__(self, config: MemoryBudgetConfig | None = None):
        """
        Initialize memory budget controller.

        Args:
            config: Configuration (uses defaults if None)
        """
        self.config = config or MemoryBudgetConfig()

        if not HAS_PSUTIL:
            print("⚠️  WARNING: psutil not available - memory monitoring disabled", file=sys.stderr)
            print(
                "   Install psutil for memory budget enforcement: pip install psutil",
                file=sys.stderr,
            )
            self.config.enabled = False
            self.process = None
        else:
            # Verify psutil is available
            try:
                self.process = psutil.Process()
            except Exception as e:
                print(f"⚠️  WARNING: Failed to initialize psutil Process: {e}", file=sys.stderr)
                print("   Memory monitoring disabled.", file=sys.stderr)
                self.config.enabled = False
                self.process = None

        # State
        self.peak_rss_mb: float = 0.0
        self.last_check_row_count: int = 0
        self.warned: bool = False
        self.check_count: int = 0

    def get_current_rss_mb(self) -> float:
        """
        Get current process RSS in MB.

        Returns:
            RSS in megabytes (OS-level measurement)
        """
        if not self.process:
            return 0.0

        try:
            # Get memory info from psutil
            mem_info = self.process.memory_info()
            rss_bytes = mem_info.rss
            rss_mb = rss_bytes / (1024 * 1024)
            return rss_mb
        except Exception as e:
            # Non-fatal: log warning and return 0
            print(f"WARNING: Failed to get RSS: {e}", file=sys.stderr)
            return 0.0

    def check(self, row_count: int) -> None:
        """
        Check memory budget at periodic intervals.

        Args:
            row_count: Current row count (used for sampling frequency)

        Raises:
            SystemExit: If RSS exceeds max_memory_mb (exit code 42)
        """
        if not self.config.enabled:
            return

        # Check if we should sample (every N rows)
        rows_since_last_check = row_count - self.last_check_row_count
        if rows_since_last_check < self.config.check_interval_rows:
            return

        # Sample RSS
        current_rss_mb = self.get_current_rss_mb()
        self.check_count += 1
        self.last_check_row_count = row_count

        # Update peak
        if current_rss_mb > self.peak_rss_mb:
            self.peak_rss_mb = current_rss_mb

        # Check thresholds
        warn_threshold_mb = self.config.max_memory_mb * self.config.warn_threshold_pct

        if current_rss_mb >= self.config.max_memory_mb:
            # ABORT: Budget exceeded
            print(
                f"\n❌ MEMORY BUDGET EXCEEDED\n"
                f"   Current RSS: {current_rss_mb:.2f} MB\n"
                f"   Max allowed: {self.config.max_memory_mb} MB\n"
                f"   Row count: {row_count:,}\n"
                f"   Check count: {self.check_count}\n"
                f"   Peak RSS: {self.peak_rss_mb:.2f} MB\n"
                f"\n"
                f"Aborting extraction to prevent OOM.\n",
                file=sys.stderr,
            )
            raise MemoryBudgetExceeded(
                f"Memory budget exceeded: RSS {current_rss_mb:.0f} MB >= max {self.config.max_memory_mb} MB"
            )

        elif current_rss_mb >= warn_threshold_mb and not self.warned:
            # WARN: Approaching limit (only once)
            print(
                f"\n⚠️  MEMORY WARNING\n"
                f"   Current RSS: {current_rss_mb:.2f} MB\n"
                f"   Threshold: {warn_threshold_mb:.2f} MB (80%)\n"
                f"   Max allowed: {self.config.max_memory_mb} MB\n"
                f"   Row count: {row_count:,}\n",
                file=sys.stderr,
            )
            self.warned = True

    def get_peak_rss_mb(self) -> float:
        """
        Get peak RSS observed during extraction.

        Returns:
            Peak RSS in megabytes
        """
        return self.peak_rss_mb

    def get_telemetry(self) -> dict:
        """
        Get telemetry data for reporting.

        Returns:
            Dictionary with peak_rss_mb, check_count, and config
        """
        return {
            "peak_rss_mb": round(self.peak_rss_mb, 2),
            "check_count": self.check_count,
            "max_memory_mb": self.config.max_memory_mb,
            "warn_threshold_mb": round(
                self.config.max_memory_mb * self.config.warn_threshold_pct, 2
            ),
            "check_interval_rows": self.config.check_interval_rows,
        }
