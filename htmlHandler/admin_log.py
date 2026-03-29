"""Append-only trace log under python_files/adminLog/ (gitignored)."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

_LOG_DIR = Path(__file__).resolve().parent.parent / "adminLog"
_LOG_FILE = _LOG_DIR / "htmlHandler_trace.txt"
_MAX_INLINE = 500


def _timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def trace(source: str, message: str, sample: str | None = None) -> None:
    """Write one timestamped line to stdout AND adminLog/htmlHandler_trace.txt."""
    line = f"[{_timestamp()}] [{source}] {message}"
    if sample is not None:
        if len(sample) > _MAX_INLINE:
            sample = sample[:_MAX_INLINE] + "…"
        line += f"\n    sample: {sample}"
    print(line, flush=True)
    _LOG_DIR.mkdir(parents=True, exist_ok=True)
    try:
        with open(_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass
