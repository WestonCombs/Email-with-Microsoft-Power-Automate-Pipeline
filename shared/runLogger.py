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
  program_errors.txt                — fatal exits and subprocess failures (always; not gated on DEBUG_MODE)

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
    from shared.settings_store import apply_runtime_settings_from_json

    apply_runtime_settings_from_json()
except Exception:
    from .project_paths import ensure_base_dir_in_environ

    ensure_base_dir_in_environ()


_TRUTHY = ("1", "true", "yes")
_MASTER_FILE = "master.txt"
_MAX_TRACE_SAMPLE = 500


def _is_truthy(raw: str | None) -> bool:
    return (raw or "").strip().lower() in _TRUTHY


_INCLUDE_TRACE_SAMPLES = _is_truthy(os.getenv("TRACE_INCLUDE_SAMPLES"))


def _logs_dir() -> Path:
    base = (os.getenv("BASE_DIR") or "").strip()
    if not base:
        raise ValueError(
            "BASE_DIR is unset — project root could not be inferred "
            '(expected scripts under a folder named "python_files").'
        )
    root = Path(base).expanduser().resolve()
    d = root / "logs"
    try:
        d.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    return d


def _email_contents_log_dir() -> Path:
    base = (os.getenv("BASE_DIR") or "").strip()
    if not base:
        raise ValueError(
            "BASE_DIR is unset — project root could not be inferred "
            '(expected scripts under a folder named "python_files").'
        )
    root = Path(base).expanduser().resolve()
    d = root / "email_contents" / "log"
    try:
        d.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    return d


def is_debug() -> bool:
    """Return True when DEBUG_MODE=1 is set in the environment (read at call time).

    Uses the current ``DEBUG_MODE`` in the environment (from ``email_sorter_settings.json`` via Settings).
    """
    return _is_truthy(os.getenv("DEBUG_MODE", "0"))


def ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _mirror_log_path(path: Path) -> Path | None:
    try:
        logs_root = _logs_dir().resolve()
        rel = path.resolve().relative_to(logs_root)
    except Exception:
        return None
    try:
        return _email_contents_log_dir() / rel
    except ValueError:
        return None


def _append_single(path: Path, text: str) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "a", encoding="utf-8") as f:
            f.write(text)
    except OSError:
        pass


def _append(path: Path, text: str) -> None:
    _append_single(path, text)
    mirror_path = _mirror_log_path(path)
    if mirror_path is None:
        return
    try:
        if mirror_path.resolve() == path.resolve():
            return
    except Exception:
        pass
    _append_single(mirror_path, text)


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
        merged = "".join(prefix + line for line in payload.splitlines(True))
        _append(path, merged)
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
    if is_debug():
        _append(_logs_dir() / f"debug_{segment}.txt", header)


def log(segment: str, text: str, *, nl: bool = True) -> None:
    """Append *text* to BASE_DIR/logs/<segment>.txt."""
    line = text + ("\n" if nl else "")
    _append(_logs_dir() / f"{segment}.txt", line)
    _append_master(segment, text, nl=nl)


def debug(segment: str, text: str, *, nl: bool = True) -> None:
    """Append *text* to BASE_DIR/logs/debug_<segment>.txt — only when DEBUG_MODE=1."""
    if not is_debug():
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


def record_program_error_exit(
    *,
    exit_code: int,
    summary: str,
    detail: str | None = None,
    source: str = "unknown",
) -> None:
    """Append one error-exit record to ``program_errors.txt`` under ``BASE_DIR/logs/``.

    Used for non-zero exits, launcher-reported subprocess failures, and similar cases.
    Always writes when ``BASE_DIR`` is set and the path is writable — independent of
    ``DEBUG_MODE``. Does nothing if ``BASE_DIR`` is missing or the write fails.
    """
    try:
        logs = _logs_dir()
    except ValueError:
        return
    sep = "─" * 72
    sum_one_line = summary.replace("\r\n", "\n").replace("\r", "\n")
    lines = [
        f"\n{sep}",
        f"{ts()}  exit_code={exit_code}  source={source}",
        f"  summary: {sum_one_line}",
    ]
    if detail:
        d = detail.replace("\r\n", "\n").replace("\r", "\n").strip()
        if d:
            lines.append("  detail:")
            lines.extend("    " + ln for ln in d.split("\n"))
    lines.append(sep)
    block = "\n".join(lines) + "\n"
    _append(logs / "program_errors.txt", block)
    _append_master("program_errors", block, nl=True)


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
