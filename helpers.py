# helpers.py
import re
import hashlib
from datetime import date, datetime
import pandas as pd


def format_day_month(dates):
    """
    Return short day.month labels for x-ticks.
    Example: [01.07, 02.07, ...]
    """
    idx = pd.to_datetime(dates)
    return [d.strftime("%d.%m") if not pd.isna(d) else "" for d in idx]


def sanitize_filename(title: str) -> str:
    """
    Create a safe filename from a plot title.
    - Lowercase
    - Spaces -> underscores
    - Remove non-alphanumeric/underscore/dash
    - Limit length to avoid filesystem issues
    """
    safe = title.lower().strip()
    safe = safe.replace(" ", "_")
    safe = re.sub(r"[^a-z0-9_\-]+", "", safe)
    if not safe:
        safe = "chart"
    return safe[:80] + ".png"  # cap length, always end with .png


def file_content_hash(file) -> str:
    """Return a short stable hash for an uploaded file-like object."""
    if file is None:
        return ""

    try:
        pos = file.tell()
        sha1 = hashlib.sha1()
        for chunk in iter(lambda: file.read(8192), b""):
            sha1.update(chunk)
        file.seek(pos)
        return sha1.hexdigest()[:10]
    except Exception:
        return "nofile"


def make_context_key(file, mode: str, date_value: datetime) -> str:
    """
    Create a stable short hash for the current analysis context.
    Context is defined by (file content, mode, date).
    Used to persist checkbox selections across reruns.
    """
    file_hash = file_content_hash(file)

    if isinstance(date_value, datetime):
        date_str = date_value.strftime("%Y-%m-%d")
    elif isinstance(date_value, date):
        date_str = date_value.strftime("%Y-%m-%d")
    elif date_value:
        date_str = str(date_value)
    else:
        date_str = ""

    raw = f"{file_hash}|{mode}|{date_str}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]
