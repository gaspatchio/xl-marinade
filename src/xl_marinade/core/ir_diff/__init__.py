# ABOUTME: Deterministic IR version diff algorithm.
# ABOUTME: Compares two production IR databases and produces a replay-complete changelist.

from xl_marinade.core.ir_diff.pipeline import IR_DIFF_CACHE_VERSION, diff_ir

__all__ = ["IR_DIFF_CACHE_VERSION", "diff_ir"]
