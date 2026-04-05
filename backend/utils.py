"""Utility functions shared across backend modules."""

from typing import Any


def to_number(value: Any) -> float | int | None:
    """Convert value to float or int, return None if not convertible."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_timestamp(value: Any) -> int | None:
    """Convert value to Unix timestamp (int), return None if not convertible."""
    from datetime import datetime
    
    if value is None:
        return None
    if isinstance(value, datetime):
        return int(value.timestamp())
    if isinstance(value, (int, float)):
        return int(value)
    return None


def avg(values: list[float | int]) -> float:
    """Calculate average of a list of numbers."""
    return (sum(values) / len(values)) if values else 0.0


def safe_avg(values: list) -> float:
    """Calculate average of a list, filtering out None values."""
    filtered = [v for v in values if v is not None]
    return sum(filtered) / len(filtered) if filtered else 0.0


def fmt_time(seconds: float) -> str:
    """Format seconds to HH:MM:SS or MM:SS format."""
    total = int(round(seconds))
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    if h > 0:
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"
