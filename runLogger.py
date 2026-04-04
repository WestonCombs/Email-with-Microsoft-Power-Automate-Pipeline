"""Centralized run logger — accumulating segment files in BASE_DIR/logs/.

All log files ACCUMULATE across runs (opened in append mode).
Each run that calls write_run_header() adds a clear divider so you can tell runs apart.

Segment files written to BASE_DIR/logs/:
  timing.txt                        — per-run timing breakdown (written by mainRunner)
  tracking_hrefs.txt                — per-email href/tracking report
  openai_extraction.txt             — per-email OpenAI field extraction results
  emailFetching.txt                 — email fetching events (written by mainRunner)
  grabbingImportantEmailContent.txt — per-email pipeline summary
  htmlHandler.txt                   — HTML processing events
  sortJSONByOrderNumber.txt         — sort events
  EnvironmentInitialization.txt     — initialization events

When DEBUG_MODE=1, additional detail-only files are also written:
  debug_tracking_hrefs.txt
  debug_grabbingImportantEmailContent.txt
  debug_htmlHandler.txt
  debug_emailFetching.txt
  (etc.)

Debug files contain the same entries as regular files PLUS extra field/URL detail
that would be noisy in the normal view. The console always stays clean regardless
of DEBUG_MODE — debug detail only goes to files.
"""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path


# Evaluated once at import time — subprocess inherits env vars from parent.
_DEBUG_MODE: bool = os.getenv("DEBUG_MODE", "0").strip().lower() in ("1", "true", "yes")


def _logs_dir() -> Path:
    base = (os.getenv("BASE_DIR") or "").strip()
    root = Path(base).expanduser().resolve() if base else Path(__file__).resolve().parent
    d = root / "logs"
    try:
        d.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    return d


def is_debug() -> bool:
    """Return True when DEBUG_MODE=1 is set in the environment."""
    return _DEBUG_MODE


def ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _append(path: Path, text: str) -> None:
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write(text)
    except OSError:
        pass


def write_run_header(segment: str, label: str = "") -> None:
    """Write a run-start divider to segment log and its debug counterpart."""
    now = ts()
    header = (
        f"\n{'=' * 72}\n"
        f"  {segment}{('  —  ' + label) if label else ''}  |  {now}\n"
        f"{'=' * 72}\n"
    )
    _append(_logs_dir() / f"{segment}.txt", header)
    if _DEBUG_MODE:
        _append(_logs_dir() / f"debug_{segment}.txt", header)


def log(segment: str, text: str, *, nl: bool = True) -> None:
    """Append *text* to BASE_DIR/logs/<segment>.txt."""
    _append(_logs_dir() / f"{segment}.txt", text + ("\n" if nl else ""))


def debug(segment: str, text: str, *, nl: bool = True) -> None:
    """Append *text* to BASE_DIR/logs/debug_<segment>.txt — only when DEBUG_MODE=1."""
    if not _DEBUG_MODE:
        return
    _append(_logs_dir() / f"debug_{segment}.txt", text + ("\n" if nl else ""))


def write_timing_entry(buffer_path: Path, data: dict) -> None:
    """Append one JSON-lines timing entry to *buffer_path*."""
    _append(buffer_path, json.dumps(data, default=str) + "\n")


def read_timing_buffer(buffer_path: Path) -> list[dict]:
    """Read all JSON-lines timing entries from *buffer_path*."""
    if not buffer_path.exists():
        return []
    entries: list[dict] = []
    try:
        for raw in buffer_path.read_text(encoding="utf-8").splitlines():
            raw = raw.strip()
            if raw:
                try:
                    entries.append(json.loads(raw))
                except json.JSONDecodeError:
                    pass
    except OSError:
        pass
    return entries
