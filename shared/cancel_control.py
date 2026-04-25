"""Shared cancellation helpers for launcher + pipeline subprocesses."""

from __future__ import annotations

import os
import subprocess
import time
from pathlib import Path

_LAUNCHER_CANCEL_FILE = ".email_sorter_cancel"


class CancelRequestedError(RuntimeError):
    """Raised when a cooperative stop request is detected."""


def cancel_file_path(base_dir: Path) -> Path:
    return base_dir / "logs" / _LAUNCHER_CANCEL_FILE


def clear_cancel_request(base_dir: Path) -> None:
    p = cancel_file_path(base_dir)
    try:
        if p.is_file():
            p.unlink()
    except OSError:
        pass


def request_cancel(base_dir: Path) -> None:
    p = cancel_file_path(base_dir)
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text("1\n", encoding="utf-8")
    except OSError:
        pass


def is_cancel_requested(base_dir: Path) -> bool:
    return cancel_file_path(base_dir).is_file()


def ensure_not_cancelled(base_dir: Path, *, context: str = "") -> None:
    if is_cancel_requested(base_dir):
        detail = f" ({context})" if context else ""
        raise CancelRequestedError(f"Pipeline stop requested{detail}")


def run_subprocess_cancellable(
    cmd: list[str],
    *,
    cwd: str,
    env: dict[str, str] | None,
    base_dir: Path,
    poll_seconds: float = 0.20,
    timeout_seconds: float | None = None,
) -> int:
    """Run a subprocess while polling cancellation and terminate on stop."""
    proc = subprocess.Popen(
        cmd,
        cwd=cwd,
        env=env if env is not None else os.environ.copy(),
    )
    try:
        started = time.monotonic()
        while True:
            rc = proc.poll()
            if rc is not None:
                return int(rc)
            if timeout_seconds and (time.monotonic() - started) > timeout_seconds:
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.wait(timeout=5)
                raise subprocess.TimeoutExpired(cmd=cmd, timeout=timeout_seconds)
            if is_cancel_requested(base_dir):
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.wait(timeout=5)
                raise CancelRequestedError("Pipeline stop requested while subprocess was running")
            time.sleep(max(0.05, poll_seconds))
    finally:
        if proc.poll() is None:
            try:
                proc.terminate()
            except OSError:
                pass
