# helpers.py
import re
import hashlib
from datetime import datetime
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


def make_context_key(file, mode: str, date_value: datetime) -> str:
    """
    Create a stable short hash for the current analysis context.
    Context is defined by (file content, mode, date).
    Used to persist checkbox selections across reruns.
    """
    # file may be None or an uploaded file-like object
    file_hash = ""
    if file is not None:
        try:
            pos = file.tell()
            content = file.read()
            file.seek(pos)
            file_hash = hashlib.sha1(content).hexdigest()[:10]
        except Exception:
            file_hash = "nofile"

    date_str = ""
    if isinstance(date_value, datetime):
        date_str = date_value.strftime("%Y-%m-%d")
    elif date_value:
        date_str = str(date_value)

    raw = f"{file_hash}|{mode}|{date_str}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]