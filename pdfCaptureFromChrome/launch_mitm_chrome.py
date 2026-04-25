"""
Launch Google Chrome in an isolated profile for mitmproxy capture.

Uses ``--user-data-dir`` under ``<BASE_DIR>/logs/pdfCaptureFromChrome`` and
``--proxy-server`` for this process only.

Typical use is via ``BASE_DIR/mitm_pdf_capture/run_pdf_capture.py`` (starts mitmdump first).

    python run_pdf_capture.py
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

try:
    from .paths import (  # type: ignore[import-not-found]
        CHROME_USER_DATA_MITM,
        DEFAULT_START_URL,
        PDF_CAPTURE_SESSION_LOG,
        PDF_CAPTURE_ROOT,
        ensure_import_path,
        normalize_start_url,
    )
except ImportError:
    from paths import (  # type: ignore[no-redef]  # noqa: E402
        CHROME_USER_DATA_MITM,
        DEFAULT_START_URL,
        PDF_CAPTURE_SESSION_LOG,
        PDF_CAPTURE_ROOT,
        ensure_import_path,
        normalize_start_url,
    )

ensure_import_path()

_CAPTURE_LOG = PDF_CAPTURE_SESSION_LOG


def _append_capture_log(message: str) -> None:
    """Same file as ``mitm_pdf_capture/run_pdf_capture`` session log — stderr is invisible when the viewer detaches the console."""
    try:
        with open(_CAPTURE_LOG, "a", encoding="utf-8", newline="\n") as f:
            f.write(f"{datetime.now().isoformat(timespec='seconds')} [launch_chrome] {message}\n")
    except OSError:
        pass


def _find_chrome_exe() -> Path | None:
    program_files = Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
    program_files_x86 = Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
    local_app = Path(os.environ.get("LOCALAPPDATA", ""))
    candidates = [
        program_files / "Google" / "Chrome" / "Application" / "chrome.exe",
        program_files_x86 / "Google" / "Chrome" / "Application" / "chrome.exe",
        local_app / "Google" / "Chrome" / "Application" / "chrome.exe",
    ]
    for p in candidates:
        if p.is_file():
            return p
    return None


def find_chrome_executable() -> Path | None:
    """Path to ``chrome.exe`` for readiness checks and optional overrides."""
    return _find_chrome_exe()


def launch_isolated_chrome(
    port: int,
    chrome_path: Path | None = None,
    *,
    start_url: str = DEFAULT_START_URL,
    remote_debugging_port: int | None = None,
    verbose: bool = True,
) -> subprocess.Popen | None:
    """Start Chrome; return ``Popen`` so the caller can terminate the browser tree."""
    chrome = chrome_path or _find_chrome_exe()
    if not chrome:
        msg = "Could not find chrome.exe. Install Google Chrome or pass --chrome-path."
        _append_capture_log(msg)
        print(msg, file=sys.stderr)
        return None

    CHROME_USER_DATA_MITM.mkdir(parents=True, exist_ok=True)

    args = [
        str(chrome),
        f"--user-data-dir={CHROME_USER_DATA_MITM.resolve()}",
        f"--proxy-server=127.0.0.1:{port}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-sync",
        "--disable-background-networking",
        "--disable-background-mode",
        start_url,
    ]
    if remote_debugging_port is not None:
        args.insert(4, f"--remote-debugging-port={remote_debugging_port}")
    if verbose:
        print("Starting isolated Chrome for mitmproxy:")
        print(" ", subprocess.list2cmdline(args))
        print(f"Profile dir (delete to reset): {CHROME_USER_DATA_MITM}")
        if remote_debugging_port is not None:
            print(f"Chrome DevTools port: {remote_debugging_port}")
        print("Chrome will close automatically after a PDF is saved.\n")

    try:
        return subprocess.Popen(args, cwd=PDF_CAPTURE_ROOT)
    except OSError as e:
        msg = f"Failed to start Chrome: {e}"
        _append_capture_log(msg)
        print(msg, file=sys.stderr)
        return None


def launch_isolated_chrome_no_proxy(
    chrome_path: Path | None = None,
    *,
    start_url: str = "about:blank",
    remote_debugging_port: int,
    verbose: bool = False,
) -> subprocess.Popen | None:
    """
    Isolated profile Chrome **without** mitmproxy. Used by HtmlCaptureController (Ctrl+Enter PDF snapshot).
    Reuses the same user-data dir as the MITM launcher for shared CA / cookies.
    """
    chrome = chrome_path or _find_chrome_exe()
    if not chrome:
        msg = "Could not find chrome.exe. Install Google Chrome or pass a chrome path."
        _append_capture_log(msg)
        if verbose:
            print(msg, file=sys.stderr)
        return None

    CHROME_USER_DATA_MITM.mkdir(parents=True, exist_ok=True)

    args: list[object] = [
        str(chrome),
        f"--user-data-dir={CHROME_USER_DATA_MITM.resolve()}",
        f"--remote-debugging-port={remote_debugging_port}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-sync",
        "--disable-background-networking",
        "--disable-background-mode",
        start_url,
    ]
    if verbose:
        _append_capture_log(
            f"no_proxy: remote-debugging-port={remote_debugging_port} start_url={start_url!r}"
        )
    try:
        return subprocess.Popen([str(x) for x in args], cwd=PDF_CAPTURE_ROOT)
    except OSError as e:
        msg = f"Failed to start Chrome (no proxy): {e}"
        _append_capture_log(msg)
        if verbose:
            print(msg, file=sys.stderr)
        return None


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Isolated Chrome proxied to local mitm (run mitmdump separately)")
    p.add_argument(
        "start_url",
        nargs="?",
        default=None,
        help=f"First page to open (default: {DEFAULT_START_URL})",
    )
    p.add_argument("--port", type=int, default=8080, help="mitmproxy port (default 8080)")
    p.add_argument(
        "--chrome-path",
        type=Path,
        default=None,
        help="Path to chrome.exe if not found automatically",
    )
    ns = p.parse_args(argv)
    url = normalize_start_url(ns.start_url)
    proc = launch_isolated_chrome(ns.port, chrome_path=ns.chrome_path, start_url=url)
    return 0 if proc is not None else 1


if __name__ == "__main__":
    raise SystemExit(main())
