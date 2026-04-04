"""Append-only trace log under {BASE_DIR}/adminLog/ (gitignored).

Loads python_files/.env from this package's parent so BASE_DIR is available.
If BASE_DIR is unset, logs fall back to python_files/adminLog/ (next to this package).
"""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

_PYTHON_FILES_DIR = Path(__file__).resolve().parent.parent
load_dotenv(_PYTHON_FILES_DIR / ".env")

_base = os.getenv("BASE_DIR")
if _base:
    _LOG_DIR = Path(_base).expanduser().resolve() / "adminLog"
else:
    _LOG_DIR = _PYTHON_FILES_DIR / "adminLog"
_LOG_FILE = _LOG_DIR / "htmlHandler_trace.txt"
_MAX_INLINE = 500
_DEBUG_MODE = os.getenv("DEBUG_MODE", "0").strip().lower() in ("1", "true", "yes")
# Set TRACE_INCLUDE_SAMPLES=1 in .env to append HTML/text snippets in trace entries.
_INCLUDE_SAMPLES = os.getenv("TRACE_INCLUDE_SAMPLES", "").strip().lower() in (
    "1",
    "true",
    "yes",
)


def _timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def trace(source: str, message: str, sample: str | None = None) -> None:
    """Write one timestamped line to adminLog/htmlHandler_trace.txt (FILE ONLY).

    Does NOT print to stdout — the clean ``» Step`` console format handles user-facing
    output. Keeps programFileOutput.txt readable.

    Optional *sample* is included in the file only when TRACE_INCLUDE_SAMPLES=1.
    """
    line = f"[{_timestamp()}] [{source}] {message}"
    if sample is not None and _INCLUDE_SAMPLES:
        if len(sample) > _MAX_INLINE:
            sample = sample[:_MAX_INLINE] + "…"
        line += f"\n    sample: {sample}"
    _LOG_DIR.mkdir(parents=True, exist_ok=True)
    try:
        with open(_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass
