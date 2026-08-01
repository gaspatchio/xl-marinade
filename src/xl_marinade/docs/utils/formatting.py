# ABOUTME: Formatting utilities for documentation display
# ABOUTME: Handles truncation of large arrays and matrices

from typing import Any


def format_value_for_display(value: Any, max_items: int = 5) -> str:
    """
    Format a value for display in documentation, truncating large structures.

    Args:
        value: The value to format (can be scalar, list, or nested list)
        max_items: Maximum number of items to show before truncating

    Returns:
        Formatted string representation
    """
    if value is None:
        return "null"

    # Handle list/array
    if isinstance(value, list):
        if not value:
            return "[]"

        # Check if it's a matrix (list of lists)
        if isinstance(value[0], list):
            rows = len(value)
            cols = len(value[0]) if rows > 0 else 0

            # For small matrices, showing content might be okay, but generally
            # we want to show dimensions for tables
            if rows <= max_items and cols <= max_items:
                # Small enough to show roughly
                return str(value)

            return f"[Table {rows}x{cols}]"

        # 1D Array
        length = len(value)
        if length <= max_items:
            return str(value)

        # Truncate
        shown_items = value[:max_items]
        shown_str = ", ".join(repr(x) for x in shown_items)
        remaining = length - max_items
        return f"[{shown_str}, ... ({remaining} more)]"

    # Handle float rounding for cleanliness
    if isinstance(value, float):
        # If it's effectively an integer, show as int
        if value.is_integer():
            return str(int(value))
        # Otherwise round to reasonable precision
        return f"{value:.4g}"

    return str(value)


def format_table_value(value: Any, max_items: int = 3) -> str:
    """
    Format a value specifically for table cells in Markdown.
    More aggressive truncation than format_value_for_display to prevent table layout issues.

    Args:
        value: The value to format (can be scalar, list, dict, or nested structures)
        max_items: Maximum number of items to show before truncating (default: 3)

    Returns:
        Formatted string safe for Markdown table cells
    """
    # Handle None explicitly
    if value is None:
        return "-"

    # Handle dict/object
    if isinstance(value, dict):
        if not value:
            return "{}"
        keys_count = len(value)
        if keys_count <= max_items:
            # Small dict - show as JSON-like
            items = [f"{k}: {v}" for k, v in list(value.items())[:max_items]]
            return "{" + ", ".join(items) + "}"
        # Large dict - just show count
        return f"{{...{keys_count} keys}}"

    # Handle list/array
    if isinstance(value, list):
        if not value:
            return "[]"

        # Check if it's a matrix (list of lists)
        if isinstance(value[0], list):
            rows = len(value)
            cols = len(value[0]) if rows > 0 else 0
            return f"[Table {rows}×{cols}]"

        # 1D Array - truncate aggressively for tables
        length = len(value)
        if length <= max_items:
            # Format items cleanly
            items = []
            for item in value:
                if isinstance(item, float):
                    if item.is_integer():
                        items.append(str(int(item)))
                    else:
                        items.append(f"{item:.4g}")
                else:
                    items.append(str(item))
            return "[" + ", ".join(items) + "]"

        # Truncate - show first max_items only
        shown_items = []
        for item in value[:max_items]:
            if isinstance(item, float):
                if item.is_integer():
                    shown_items.append(str(int(item)))
                else:
                    shown_items.append(f"{item:.4g}")
            else:
                shown_items.append(str(item))

        remaining = length - max_items
        return f"[{', '.join(shown_items)}, ... ({remaining} more)]"

    # Handle bool (must come before int since bool is subclass of int)
    if isinstance(value, bool):
        return str(value)

    # Handle float
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.4f}"

    # Handle int
    if isinstance(value, int):
        return str(value)

    # Handle string - truncate if too long
    if isinstance(value, str):
        if len(value) > 50:
            return value[:47] + "..."
        return value

    # Fallback for other types
    str_repr = str(value)
    if len(str_repr) > 50:
        return str_repr[:47] + "..."
    return str_repr


def _format_scalar(item: Any) -> str:
    if isinstance(item, bool):
        return str(item)
    if isinstance(item, float):
        if item.is_integer():
            return str(int(item))
        return f"{item:.4g}"
    return str(item)


def format_time_series(value: Any, period_labels: Any, max_items: int = 6) -> str:
    """
    Render a time-linked series as a labelled t-indexed sequence (Phase B1).

    Pairs each value with its real period label (the actual axis cell), e.g.
    "t=1: 100.0, t=2: 95.0, ...". This replaces the shape-blind raw-array
    truncation for genuine time series. Falls back to plain value formatting if
    the value is not a 1D list.

    Args:
        value: The series values (a 1D list for a time series).
        period_labels: The axis cell values used as t labels (may be None/short).
        max_items: How many (label, value) pairs to show before truncating.

    Returns:
        A compact "t=<label>: <value>" sequence string for a Markdown cell.
    """
    if not isinstance(value, list) or not value or isinstance(value[0], list):
        # Not a 1D series - defer to the standard table formatter.
        return format_table_value(value, max_items=max_items)

    labels = period_labels if isinstance(period_labels, list) else []

    # Only pair per-t when the values line up 1:1 with the period labels. A
    # flattened 2D grid (len(value) != len(labels)) would mis-align labels, so
    # defer to the standard formatter rather than fabricate a wrong t mapping.
    if labels and len(value) != len(labels):
        return format_table_value(value, max_items=max_items)
    length = len(value)
    shown = min(length, max_items)
    pairs = []
    for i in range(shown):
        label = labels[i] if i < len(labels) else (i + 1)
        pairs.append(f"t={_format_scalar(label)}: {_format_scalar(value[i])}")

    rendered = ", ".join(pairs)
    if length > shown:
        rendered += f", ... ({length - shown} more)"
    return rendered
