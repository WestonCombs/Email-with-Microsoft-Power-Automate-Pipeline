"""Paths and shared constants for the PDF capture package."""

from __future__ import annotations

import os
import sys
from pathlib import Path

# Directory containing this file; still used as ``cwd`` for code-side resources.
PDF_CAPTURE_ROOT = Path(__file__).resolve().parent

try:
    from dotenv import load_dotenv

    load_dotenv(PDF_CAPTURE_ROOT.parent / ".env", override=False)
except ImportError:
    pass

DEFAULT_START_URL = "http://mitm.it"

# First-page preview in the success dialog (about 80% of 560x720).
PREVIEW_MAX_WIDTH = 448
PREVIEW_MAX_HEIGHT = 576


def _project_root() -> Path | None:
    base_raw = (os.getenv("BASE_DIR") or "").strip()
    if not base_raw:
        return None
    return Path(base_raw).expanduser().resolve()


def _ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def pdf_capture_runtime_dir() -> Path:
    project_root = _project_root()
    if project_root is None:
        return _ensure_dir(PDF_CAPTURE_ROOT)
    return _ensure_dir(project_root / "logs" / "pdfCaptureFromChrome")


def default_pdf_output_dir() -> Path:
    project_root = _project_root()
    if project_root is None:
        return _ensure_dir(pdf_capture_runtime_dir() / "captured_pdfs")
    return _ensure_dir(project_root / "email_contents" / "pdf")


PDF_CAPTURE_RUNTIME_DIR = pdf_capture_runtime_dir()

# Written by the mitm addon when the first PDF is saved; ``run_pdf_capture.py`` watches this file.
PDF_CAPTURE_DONE_FILE = PDF_CAPTURE_RUNTIME_DIR / ".pdfCaptureFromChrome_done.json"
PDF_CAPTURE_SESSION_LOG = PDF_CAPTURE_RUNTIME_DIR / "pdf_capture_session.log"
PDF_CAPTURE_STDOUT_LOG = PDF_CAPTURE_RUNTIME_DIR / "mitmdump.stdout.log"
PDF_CAPTURE_STDERR_LOG = PDF_CAPTURE_RUNTIME_DIR / "mitmdump.stderr.log"
CHROME_USER_DATA_MITM = PDF_CAPTURE_RUNTIME_DIR / "chrome_user_data_mitm"


def ensure_import_path() -> None:
    """Allow ``import paths`` / sibling imports when running scripts from any cwd."""
    p = str(PDF_CAPTURE_ROOT)
    if p not in sys.path:
        sys.path.insert(0, p)


def normalize_start_url(raw: str | None) -> str:
    """Default to mitm.it; ensure http(s) scheme."""
    s = (raw or "").strip()
    if not s:
        return DEFAULT_START_URL
    if not s.lower().startswith(("http://", "https://")):
        return "https://" + s
    return s


def is_mitm_it_install_url(url: str) -> bool:
    """True when Chrome should open the mitmproxy CA install page."""
    return url.rstrip("/").lower() in ("http://mitm.it", "https://mitm.it")


def split_debug_positional(args: list[str]) -> tuple[list[str], bool]:
    """If the last token is ``0`` or ``1``, treat it as quiet (0) / debug (1) and strip it.

    Otherwise debug defaults to ``True`` (verbose logging, same as before).
    """
    if not args:
        return [], True
    last = args[-1].strip()
    if last in ("0", "1"):
        return args[:-1], last == "1"
    return args, True
