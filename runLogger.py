"""Centralized run logger — accumulating segment files in BASE_DIR/logs/.

All log files ACCUMULATE across runs (opened in append mode).
Each run that calls write_run_header() adds a clear divider so you can tell runs apart.

Segment files written to BASE_DIR/logs/:
  master.txt                        — interleaved copy of every segment + debug lines
  timing.txt                        — per-run timing breakdown (written by mainRunner)
  tracking_hrefs.txt                — per-email href/tracking report
  openai_extraction.txt             — per-email OpenAI field extraction results
  emailFetching.txt                 — email fetching events (written by mainRunner)
  grabbingImportantEmailContent.txt — per-email pipeline summary
  htmlHandler.txt                   — HTML processing events (includes trace() lines)
  sortJSONByOrderNumber.txt         — sort events
  environmentInitialization.txt     — initialization events

When DEBUG_MODE=1, additional detail-only files are also written:
  debug_tracking_hrefs.txt
  debug_grabbingImportantEmailContent.txt
  debug_htmlHandler.txt
  debug_emailFetching.txt
  (etc.)

Debug lines are also appended to master.txt with a [DEBUG][segment] prefix.

Debug files contain the same entries as regular files PLUS extra field/URL detail
that would be noisy in the normal view. The console always stays clean regardless
of DEBUG_MODE — debug detail only goes to files.
"""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path

try:
    from dotenv import load_dotenv

    load_dotenv(Path(__file__).resolve().parent / ".env")
except ImportError:
    pass


# Evaluated once at import time — subprocess inherits env vars from parent.
_DEBUG_MODE: bool = os.getenv("DEBUG_MODE", "0").strip().lower() in ("1", "true", "yes")
_MASTER_FILE = "master.txt"
_MAX_TRACE_SAMPLE = 500
_INCLUDE_TRACE_SAMPLES = os.getenv("TRACE_INCLUDE_SAMPLES", "").strip().lower() in (
    "1",
    "true",
    "yes",
)


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


def _append_master(segment: str, text: str, *, debug_line: bool = False, nl: bool = True) -> None:
    """Mirror one log write into master.txt with a segment prefix on each line."""
    if segment == Path(_MASTER_FILE).stem:
        return
    payload = text + ("\n" if nl else "")
    if not payload:
        return
    prefix = f"[DEBUG][{segment}] " if debug_line else f"[{segment}] "
    path = _logs_dir() / _MASTER_FILE
    try:
        with open(path, "a", encoding="utf-8") as f:
            for line in payload.splitlines(True):
                f.write(prefix + line)
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
    _append_master(segment, header, nl=True)
    if _DEBUG_MODE:
        _append(_logs_dir() / f"debug_{segment}.txt", header)


def log(segment: str, text: str, *, nl: bool = True) -> None:
    """Append *text* to BASE_DIR/logs/<segment>.txt."""
    line = text + ("\n" if nl else "")
    _append(_logs_dir() / f"{segment}.txt", line)
    _append_master(segment, text, nl=nl)


def debug(segment: str, text: str, *, nl: bool = True) -> None:
    """Append *text* to BASE_DIR/logs/debug_<segment>.txt — only when DEBUG_MODE=1."""
    if not _DEBUG_MODE:
        return
    line = text + ("\n" if nl else "")
    _append(_logs_dir() / f"debug_{segment}.txt", line)
    _append_master(segment, text, debug_line=True, nl=nl)


def trace(source: str, message: str, sample: str | None = None) -> None:
    """Append one timestamped trace line to logs/htmlHandler.txt (file only, no console).

    Formerly adminLog/htmlHandler_trace.txt; merged into the htmlHandler segment.
    Optional *sample* is included when TRACE_INCLUDE_SAMPLES=1 in the environment.
    """
    line = f"[{ts()}] [{source}] {message}"
    if sample is not None and _INCLUDE_TRACE_SAMPLES:
        frag = sample
        if len(frag) > _MAX_TRACE_SAMPLE:
            frag = frag[:_MAX_TRACE_SAMPLE] + "…"
        line += f"\n    sample: {frag}"
    log("htmlHandler", line)


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
