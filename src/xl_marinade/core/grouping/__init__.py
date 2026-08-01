# ABOUTME: Public interface for constant binding grouping module
# ABOUTME: Exports core grouping functions and classes for merging constant ranges

from xl_marinade.core.grouping.geometry import BoundingBox, compute_bounding_box
from xl_marinade.core.grouping.overlap_reporter import detect_overlaps, generate_overlap_report
from xl_marinade.core.grouping.rules import merge_constant_ranges

__all__ = [
    "BoundingBox",
    "compute_bounding_box",
    "merge_constant_ranges",
    "detect_overlaps",
    "generate_overlap_report",
]
