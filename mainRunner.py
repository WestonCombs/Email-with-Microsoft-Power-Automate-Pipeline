"""Main runner — replaces the Microsoft Power Automate flow.

Usage:
    python mainRunner.py
    python mainRunner.py --skip-email-fetch
    python mainRunner.py --custom-import-html

Reads Microsoft Graph / OpenAI settings from ``python_files/.env`` (set ``BASE_DIR`` in Email Sorter → Settings — it is saved there), then:
  1. Runs environment initialization (folder/file verification)
  2. Fetches all emails from the configured mailbox folder (Microsoft Graph, OAuth 2.0),
     unless ``--skip-email-fetch`` (reprocess HTML already under ``email_contents/html/``
     and ``email_contents/pdf/`` — see :func:`_discover_local_email_html_files`)
     or ``--custom-import-html`` (debug: ``BASE_DIR/custom_import_html_files/*.html`` only;
     Graph metadata is filled with placeholders — see :func:`_custom_import_placeholders`).
  3. For each email: writes HTML body, runs the extraction pipeline
  4. Sorts the accumulated JSON results by order number
  5. Creates the Excel workbook
"""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

_PYTHON_FILES_DIR = Path(__file__).resolve().parent
_MIN_PYTHON = (3, 9)
_REQUIREMENTS_FILE = _PYTHON_FILES_DIR / "requirements.txt"


def _parse_requirement_lines(req_path: Path) -> list[str]:
    lines: list[str] = []
    if not req_path.is_file():
        return lines
    for raw in req_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        lines.append(line)
    return lines


def _requirements_satisfied_fallback() -> bool:
    """If packaging is unavailable, verify the usual project imports exist."""
    for mod in ("bs4", "dotenv", "openpyxl", "openai", "msal"):
        try:
            __import__(mod)
        except ImportError:
            return False
    return True


def _requirements_satisfied(req_path: Path) -> bool:
    """Return True if installed distributions satisfy every line in requirements.txt."""
    try:
        from importlib.metadata import PackageNotFoundError, version
        from packaging.requirements import Requirement
        from packaging.version import Version
    except ImportError:
        return _requirements_satisfied_fallback()

    for line in _parse_requirement_lines(req_path):
        try:
            req = Requirement(line)
        except Exception:
            continue
        try:
            installed = version(req.name)
        except PackageNotFoundError:
            return False
        try:
            if not req.specifier.contains(Version(installed), prereleases=True):
                return False
        except Exception:
            return False
    return True


def _pip_available() -> bool:
    try:
        r = subprocess.run(
            [sys.executable, "-m", "pip", "--version"],
            capture_output=True,
            text=True,
            timeout=30,
        )
        return r.returncode == 0
    except (OSError, subprocess.TimeoutExpired):
        return False


def _ensure_runtime_ready() -> None:
    """Verify Python version, pip, and requirements before importing project dependencies."""
    if sys.version_info < _MIN_PYTHON:
        ver = ".".join(str(x) for x in sys.version_info[:3])
        need = ".".join(str(x) for x in _MIN_PYTHON)
        print(
            f"ERROR: This project needs Python {need} or newer.\n"
            f"  Current interpreter: {sys.executable}\n"
            f"  Version reported: {ver}\n"
            "  Install Python from https://www.python.org/downloads/ "
            "and ensure `python` points to it (or use the py launcher on Windows).",
            file=sys.stderr,
        )
        sys.exit(1)

    if not _pip_available():
        print(
            "ERROR: pip is not available for this Python installation.\n"
            f"  Interpreter: {sys.executable}\n"
            "  Reinstall Python from https://www.python.org/downloads/ "
            "and enable 'pip' / 'Add Python to PATH' in the installer.",
            file=sys.stderr,
        )
        sys.exit(1)

    if not _REQUIREMENTS_FILE.is_file():
        print(
            f"WARNING: {_REQUIREMENTS_FILE.name} not found; skipping dependency install.",
            file=sys.stderr,
        )
    elif not _requirements_satisfied(_REQUIREMENTS_FILE):
        _run_pip_install()

    _env_path = _PYTHON_FILES_DIR / ".env"
    if not _env_path.is_file():
        print(
            f"WARNING: {_env_path.name} not found. Copy .env.example to .env, then set the project folder in Email Sorter → Settings and other keys as needed.",
            file=sys.stderr,
        )


def _run_pip_install() -> None:
    """Install packages from requirements.txt (caller verified file exists and deps missing)."""

    print(
        "Some packages from requirements.txt are missing or outdated.\n"
        "Installing dependencies with: "
        f"{sys.executable} -m pip install -r {_REQUIREMENTS_FILE.name}\n"
    )
    try:
        r = subprocess.run(
            [sys.executable, "-m", "pip", "install", "-r", str(_REQUIREMENTS_FILE)],
            cwd=str(_PYTHON_FILES_DIR),
        )
    except OSError as e:
        print(f"ERROR: Could not run pip: {e}", file=sys.stderr)
        sys.exit(1)
    if r.returncode != 0:
        print(
            "\nERROR: pip install failed. Check your network connection and try:\n"
            f"  {sys.executable} -m pip install -r {_REQUIREMENTS_FILE}",
            file=sys.stderr,
        )
        sys.exit(1)

    if not _requirements_satisfied(_REQUIREMENTS_FILE):
        print(
            "ERROR: Dependencies still do not match requirements.txt after install.\n"
            f"  Try manually: {sys.executable} -m pip install -r {_REQUIREMENTS_FILE}",
            file=sys.stderr,
        )
        sys.exit(1)


if __name__ == "__main__":
    _ensure_runtime_ready()

from dotenv import load_dotenv

load_dotenv(_PYTHON_FILES_DIR / ".env")

from shared import runLogger as RL
from emailFetching.emailFetcher import (
    extract_email,
    extract_sender_name,
    fetch_emails,
)

BASE_DIR_ENV = "BASE_DIR"
OPENAI_USAGE_REL = Path("logs") / "openai usage"

# Debug ``--custom-import-html``: substitute Graph envelope fields (not read from each file).
_CUSTOM_IMPORT_LABEL = "customImportHTML"
_CUSTOM_IMPORT_EMAIL = "customImportHTML@local.invalid"


def _custom_import_outlook_env() -> dict[str, str]:
    from_raw = f"{_CUSTOM_IMPORT_LABEL} <{_CUSTOM_IMPORT_EMAIL}>"
    return {
        "OUTLOOK_FROM_RAW": from_raw,
        "OUTLOOK_SENT_LINE": _CUSTOM_IMPORT_LABEL,
        "OUTLOOK_TO_LINE": _CUSTOM_IMPORT_LABEL,
        "OUTLOOK_HEADER_TITLE": _CUSTOM_IMPORT_LABEL,
    }


# ──────────────────────────────────────────────
# OpenAI usage file management
# ──────────────────────────────────────────────
def _next_usage_index(usage_dir: Path) -> int:
    """Return N where usageN.txt is the next available filename."""
    if not usage_dir.exists():
        return 1
    max_n = 0
    pattern = re.compile(r"^usage(\d+)\.txt$", re.IGNORECASE)
    for p in usage_dir.iterdir():
        if p.is_file():
            m = pattern.match(p.name)
            if m:
                max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def create_usage_log(base_dir: Path, flow_started_at: datetime) -> Path:
    """Create usageN.txt with a header line. Returns the file path."""
    usage_dir = base_dir / OPENAI_USAGE_REL
    usage_dir.mkdir(parents=True, exist_ok=True)
    n = _next_usage_index(usage_dir)
    path = usage_dir / f"usage{n}.txt"
    header = f"Flow started: {flow_started_at.strftime('%Y-%m-%d %H:%M:%S')}\n\n"
    path.write_text(header, encoding="utf-8")
    return path


def print_usage_summary(usage_log_path: Path) -> None:
    """Read the usage log and print a summary block to stdout."""
    if not usage_log_path.exists():
        return
    text = usage_log_path.read_text(encoding="utf-8")
    lines = [l for l in text.splitlines() if "prompt_tokens=" in l]
    if not lines:
        print(f"\n  OpenAI usage: no API calls were made this run.")
        print(f"  Usage log: {usage_log_path}")
        return

    total_prompt = 0
    total_completion = 0
    total_tokens = 0
    total_cost = 0.0
    elapsed_values: list[float] = []

    for line in lines:
        m_p = re.search(r"prompt_tokens=(\d+)", line)
        m_c = re.search(r"completion_tokens=(\d+)", line)
        m_t = re.search(r"total_tokens=(\d+)", line)
        m_cost = re.search(r"cost=\$([0-9.]+)", line)
        m_elapsed = re.search(r"elapsed_secs=([0-9.]+)", line)
        if m_p:
            total_prompt += int(m_p.group(1))
        if m_c:
            total_completion += int(m_c.group(1))
        if m_t:
            total_tokens += int(m_t.group(1))
        if m_cost:
            total_cost += float(m_cost.group(1))
        if m_elapsed:
            elapsed_values.append(float(m_elapsed.group(1)))

    avg_time = sum(elapsed_values) / len(elapsed_values) if elapsed_values else 0.0

    print(f"\n  OpenAI usage summary ({len(lines)} API call(s)):")
    print(f"    prompt_tokens:     {total_prompt:,}")
    print(f"    completion_tokens: {total_completion:,}")
    print(f"    total_tokens:      {total_tokens:,}")
    print(f"    total cost:        ${total_cost:.4f}")
    print(f"    avg time/email:    {avg_time:.2f}s")
    print(f"  Usage log: {usage_log_path}")


# Must match grabbingImportantEmailContent.EXIT_OPENAI_RATE_LIMIT_FATAL
_OPENAI_FATAL_EXIT = 3

_LAUNCHER_PROGRESS_ENV = "EMAIL_SORTER_LAUNCHER_PROGRESS"
_LAUNCHER_CANCEL_FILE = ".email_sorter_cancel"


def _launcher_cancel_path(base_dir: Path) -> Path:
    return base_dir / "logs" / _LAUNCHER_CANCEL_FILE


def _clear_launcher_cancel(base_dir: Path) -> None:
    p = _launcher_cancel_path(base_dir)
    try:
        if p.is_file():
            p.unlink()
    except OSError:
        pass


def _launcher_cancel_requested(base_dir: Path) -> bool:
    return _launcher_cancel_path(base_dir).is_file()


def _emit_run_launcher_progress(pct: int, msg: str = "") -> None:
    """Machine-readable line for email_sorter_launcher (stdout, line-buffered)."""
    if os.getenv(_LAUNCHER_PROGRESS_ENV) != "1":
        return
    pct = max(0, min(100, int(pct)))
    line = f"EMAIL_SORTER_RUN_PROGRESS pct={pct}"
    if msg:
        line += " msg=" + msg.replace("\n", " ").replace("\r", "")[:160]
    print(line, flush=True)


def _run_pct_after_email(i: int, n_emails: int) -> int:
    """Map completed email index i (1..n) into 13..85 (last 15% reserved for sort + Excel)."""
    if n_emails <= 0:
        return 85
    return 13 + int(72 * i / n_emails)


def _prompt_openai_moderator_action() -> None:
    """Tell the user (console + dialog) that OpenAI quota/rate must be fixed by a moderator."""
    msg = (
        "OpenAI failed after all automatic retries (rate limit / quota exhausted).\n\n"
        "A moderator must fix the OPENAI_API_KEY and the OpenAI account it is tied to "
        "(billing, limits, and key scope at platform.openai.com).\n\n"
        "This run has been stopped. Remaining emails were not processed."
    )
    print("\n" + "=" * 60, file=sys.stderr)
    print("FATAL: OpenAI API — action required (moderator)", file=sys.stderr)
    print("=" * 60, file=sys.stderr)
    print(msg, file=sys.stderr)
    print("=" * 60 + "\n", file=sys.stderr)
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showerror("OpenAI — moderator action required", msg, parent=root)
        root.destroy()
    except Exception:
        pass


def _rebuild_email_html_archive_folder(html_dir: Path) -> None:
    """Ensure archived HTML folder exists; do not clear it (same idea as the PDF folder)."""
    html_dir.mkdir(parents=True, exist_ok=True)


# ──────────────────────────────────────────────
# Step runners (each invokes its script as a subprocess so cwd and __main__
# match standalone execution)
# ──────────────────────────────────────────────
def run_environment_init() -> float:
    """Returns elapsed seconds."""
    script = _PYTHON_FILES_DIR / "EnvironmentInitialization" / "runner.py"
    print("[Step 1] Environment initialization ...")
    t = time.perf_counter()
    result = subprocess.run([sys.executable, str(script)], cwd=str(script.parent))
    elapsed = time.perf_counter() - t
    if result.returncode != 0:
        raise RuntimeError(f"Environment initialization failed (exit {result.returncode})")
    print(f"[Step 1] Done  ({elapsed:.2f}s)\n")
    return elapsed


def run_grabbing_important_content(
    html_file: str,
    subject: str,
    sender_email: str,
    sender_name: str,
    *,
    usage_log_path: Path | None = None,
    timing_buffer_path: Path | None = None,
    outlook_pdf_header_env: dict[str, str] | None = None,
) -> None:
    script = (
        _PYTHON_FILES_DIR
        / "grabbingImportantEmailContent"
        / "grabbingImportantEmailContent.py"
    )
    cmd = [
        sys.executable,
        str(script),
        "--file", html_file,
        "--subject", subject,
        "--email", sender_email,
        "--sender-name", sender_name,
    ]
    env = os.environ.copy()
    if usage_log_path:
        env["OPENAI_USAGE_LOG_PATH"] = str(usage_log_path)
    if timing_buffer_path:
        env["TIMING_BUFFER_PATH"] = str(timing_buffer_path)
    if outlook_pdf_header_env:
        env["OUTLOOK_PREPEND_PDF_HEADER"] = "1"
        for key, value in outlook_pdf_header_env.items():
            env[key] = value
    result = subprocess.run(cmd, cwd=str(_PYTHON_FILES_DIR), env=env)
    if result.returncode == _OPENAI_FATAL_EXIT:
        _prompt_openai_moderator_action()
        sys.exit(_OPENAI_FATAL_EXIT)
    if result.returncode != 0:
        print(f"  WARNING: extractor exited with code {result.returncode}")


def run_sort_json() -> float:
    """Returns elapsed seconds."""
    script = _PYTHON_FILES_DIR / "sortJSONByOrderNumber" / "sortJSONByOrderNumber.py"
    t = time.perf_counter()
    result = subprocess.run([sys.executable, str(script)], cwd=str(_PYTHON_FILES_DIR))
    elapsed = time.perf_counter() - t
    if result.returncode != 0:
        print(f"  WARNING: sortJSONByOrderNumber exited with code {result.returncode}")
    return elapsed


def run_create_excel() -> float:
    """Returns elapsed seconds."""
    script = _PYTHON_FILES_DIR / "createExcelDocument" / "createExcelDocument.py"
    t = time.perf_counter()
    result = subprocess.run([sys.executable, str(script)], cwd=str(_PYTHON_FILES_DIR))
    elapsed = time.perf_counter() - t
    if result.returncode != 0:
        print(f"  WARNING: createExcelDocument exited with code {result.returncode}")
    return elapsed


def _fmt(seconds: float) -> str:
    """Human-readable seconds with at most 2 decimal places."""
    return f"{seconds:.2f}s"


def _discover_local_email_html_files(base_dir: Path) -> list[Path]:
    """HTML sources to reprocess when Graph fetch is skipped.

    After a normal run, bodies are often **renamed** out of ``file1.html`` form: PDFs and
    archived copies live under ``email_contents/pdf/`` and especially ``email_contents/html/``
    with descriptive names. We gather every ``*.html`` in those two folders (excluding temp
    rename stubs), sorted for a stable order.
    """
    html_dir = base_dir / "email_contents" / "html"
    pdf_dir = base_dir / "email_contents" / "pdf"
    found: list[Path] = []
    for folder in (html_dir, pdf_dir):
        if not folder.is_dir():
            continue
        for p in folder.glob("*.html"):
            if not p.is_file():
                continue
            if p.name.startswith("__tmp_rename_"):
                continue
            found.append(p.resolve())

    seen: set[str] = set()
    uniq: list[Path] = []
    for p in sorted(found, key=lambda x: (str(x.parent).lower(), x.name.lower())):
        key = str(p)
        if key not in seen:
            seen.add(key)
            uniq.append(p)
    return uniq


def _discover_custom_import_html_files(base_dir: Path) -> list[Path]:
    """``*.html`` under ``BASE_DIR/custom_import_html_files/`` (debug import), stable order."""
    folder = base_dir / "custom_import_html_files"
    if not folder.is_dir():
        return []
    found: list[Path] = []
    for p in folder.glob("*.html"):
        if p.is_file():
            found.append(p.resolve())
    found.sort(key=lambda x: x.name.lower())
    return found


def _parse_saved_email_html_metadata(html_text: str) -> tuple[str, str, str]:
    """Subject, sender email, sender name from saved outlook-style ``fileN.html``."""

    def _cell_after_label(label: str) -> str:
        pat = rf">\s*{re.escape(label)}\s*</td>\s*<td[^>]*>(.*?)</td>"
        m = re.search(pat, html_text, re.IGNORECASE | re.DOTALL)
        if not m:
            return ""
        inner = m.group(1)
        inner = re.sub(r"<[^>]+>", " ", inner)
        return " ".join(html.unescape(inner).split())

    subj = _cell_after_label("Subject")
    from_raw = _cell_after_label("From")
    if not from_raw.strip():
        from_raw = "Unknown <unknown@local.invalid>"
    em = extract_email(from_raw)
    sn = extract_sender_name(from_raw)
    return subj, em, sn


def _print_and_log_timing_summary(
    *,
    flow_started_at: datetime,
    total_s: float,
    init_s: float,
    fetch_s: float,
    email_count: int,
    process_s: float,
    sort_s: float,
    excel_s: float,
    timing_entries: list[dict],
    usage_log_path: Path | None,
    fetch_skipped: bool = False,
) -> None:
    """Print the full timing summary to console and append it to logs/timing.txt."""
    processed = [e for e in timing_entries if not e.get("is_duplicate")]
    duplicates = [e for e in timing_entries if e.get("is_duplicate")]

    def _avg(key: str, entries: list[dict]) -> float:
        vals = [e[key] for e in entries if key in e and isinstance(e[key], (int, float))]
        return sum(vals) / len(vals) if vals else 0.0

    def _count_ran(key: str, entries: list[dict]) -> int:
        return sum(1 for e in entries if e.get(key))

    avg_total   = _avg("total_s",    processed)
    avg_step1   = _avg("step1_s",    processed)
    avg_step2   = _avg("step2_s",    processed)
    avg_step3   = _avg("step3_s",    processed)
    avg_step4   = _avg("step4_s",    processed)
    avg_step5   = _avg("step5_s",    processed)
    avg_step5b  = _avg("step5b_s",   processed)
    n_step5     = _count_ran("step5_ran",  processed)
    n_step5b    = _count_ran("step5b_ran", processed)
    n_proc      = len(processed)
    n_dup       = len(duplicates)

    tracking_single   = sum(1 for e in processed if e.get("tracking_result") == "single")
    tracking_multiple = sum(1 for e in processed if e.get("tracking_result") == "multiple")
    tracking_none     = sum(1 for e in processed if e.get("tracking_result") == "none")

    # OpenAI cost from usage log
    total_tokens = 0
    total_cost   = 0.0
    n_api_calls  = 0
    if usage_log_path and usage_log_path.exists():
        for line in usage_log_path.read_text(encoding="utf-8").splitlines():
            if "total_tokens=" in line:
                m = re.search(r"total_tokens=(\d+)", line)
                mc = re.search(r"cost=\$([0-9.]+)", line)
                if m:
                    total_tokens += int(m.group(1))
                    n_api_calls += 1
                if mc:
                    total_cost += float(mc.group(1))

    W = 60
    sep = "=" * W
    if fetch_skipped:
        fetch_line = (
            f"  Email fetching:        {_fmt(fetch_s)}  "
            f"(skipped Graph — {email_count} local file(s))"
        )
    else:
        fetch_line = (
            f"  Email fetching:        {_fmt(fetch_s)}  ({email_count} email(s) retrieved)"
        )
    lines: list[str] = [
        "",
        sep,
        f"  Run complete: {flow_started_at.strftime('%Y-%m-%d %H:%M:%S')}",
        sep,
        "",
        f"  Total Program Duration: {_fmt(total_s)}",
        "",
        f"  Environment init:      {_fmt(init_s)}",
        fetch_line,
        f"  Email processing:      {_fmt(process_s)}  "
        f"({n_proc} processed, {n_dup} duplicate(s) skipped)",
    ]
    if n_proc:
        lines += [
            f"    avg per email:         {_fmt(avg_total)}",
            f"    -- HTML read:          {_fmt(avg_step1)}/email",
            f"    -- plaintext conv:     {_fmt(avg_step2)}/email",
            f"    -- href extraction:    {_fmt(avg_step3)}/email",
            f"    -- redirect resolve:   {_fmt(avg_step4)}/email  "
            f"[{sum(1 for e in processed if e.get('hrefs_fetchable',0)>0)}/{n_proc} had http links]",
            f"    -- OpenAI call:        {_fmt(avg_step5)}/email  [{n_step5}/{n_proc} made API call]",
            f"    -- gift card check:    {_fmt(avg_step5b)}/email  [{n_step5b}/{n_proc} ran check]",
        ]
    lines += [
        f"  Sort JSON:             {_fmt(sort_s)}",
        f"  Create Excel:          {_fmt(excel_s)}",
        "",
    ]
    if n_api_calls:
        lines.append(f"  OpenAI: {n_api_calls} call(s)  |  {total_tokens:,} tokens  |  ${total_cost:.4f}")
    if n_proc:
        lines.append(
            f"  Tracking links: {tracking_single} single  |  "
            f"{tracking_multiple} multiple  |  {tracking_none} none"
        )
    lines += ["", sep, ""]

    block = "\n".join(lines)
    print(block)

    # Append to logs/timing.txt
    RL.log("timing", block)


# ──────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────
def main() -> None:
    t_main_start = time.perf_counter()
    flow_started_at = datetime.now()

    argp = argparse.ArgumentParser(description="Email Sorter main runner")
    argp.add_argument(
        "--skip-email-fetch",
        action="store_true",
        help=(
            "Do not connect to Microsoft Graph. Reprocess *.html under "
            "email_contents/html/ and email_contents/pdf/ (see _discover_local_email_html_files)."
        ),
    )
    argp.add_argument(
        "--custom-import-html",
        action="store_true",
        help=(
            "Do not connect to Microsoft Graph. Process *.html under "
            "BASE_DIR/custom_import_html_files/ with placeholder sender/subject (debug)."
        ),
    )
    args, _unknown = argp.parse_known_args()
    skip_email_fetch = bool(args.skip_email_fetch)
    custom_import_html = bool(args.custom_import_html)
    if skip_email_fetch and custom_import_html:
        print(
            "ERROR: use only one of --skip-email-fetch or --custom-import-html.",
            file=sys.stderr,
        )
        sys.exit(1)

    base_dir_raw = os.getenv(BASE_DIR_ENV)
    if not base_dir_raw:
        print(
            f'ERROR: {BASE_DIR_ENV} is not set. Set it in Email Sorter → Settings ("Project folder on disk") and Save.',
            file=sys.stderr,
        )
        sys.exit(1)
    base_dir = Path(base_dir_raw).expanduser().resolve()
    _clear_launcher_cancel(base_dir)
    _emit_run_launcher_progress(1, "Starting…")

    azure_client_id = (os.getenv("AZURE_CLIENT_ID") or "").strip()
    azure_tenant_id = (os.getenv("AZURE_TENANT_ID") or "common").strip()
    auth_flow = (os.getenv("GRAPH_AUTH_FLOW") or "interactive").strip()
    mail_folder = (
        os.getenv("GRAPH_MAIL_FOLDER")
        or os.getenv("IMAP_MAIL_FOLDER")
        or "INBOX"
    ).strip()
    token_cache_raw = os.getenv("GRAPH_TOKEN_CACHE_PATH")
    token_cache_path = (
        Path(token_cache_raw).expanduser().resolve()
        if token_cache_raw
        else (_PYTHON_FILES_DIR / ".graph_token_cache.bin")
    )
    debug_mode = RL.is_debug()

    if not skip_email_fetch and not custom_import_html and not azure_client_id:
        print(
            "ERROR: AZURE_CLIENT_ID must be set in .env (Azure app registration client ID).",
            file=sys.stderr,
        )
        sys.exit(1)

    demo_mode = os.getenv("DEMO_MODE", "0").strip().lower() in ("1", "true", "yes")

    W = 60
    print(f"\n{'=' * W}")
    print(f"  Email Sorter  |  {flow_started_at.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Folder: {mail_folder!r}  |  DEBUG: {'ON' if debug_mode else 'off'}")
    if skip_email_fetch:
        print(
            "  Email fetch: SKIPPED (--skip-email-fetch) — reusing saved *.html under "
            "email_contents/html/ and email_contents/pdf/"
        )
    elif custom_import_html:
        print(
            "  Email fetch: SKIPPED (--custom-import-html) — using *.html under "
            "BASE_DIR/custom_import_html_files/ (placeholder From/Subject for debug)"
        )
    print(f"{'=' * W}\n")

    # ── Prepare timing buffer (temp JSONL) ───────────────────────
    timing_buffer_path = base_dir / "logs" / ".timing_buffer.jsonl"
    (base_dir / "logs").mkdir(parents=True, exist_ok=True)
    try:
        timing_buffer_path.write_text("", encoding="utf-8")
    except OSError:
        timing_buffer_path = None

    # ── Run headers in segment logs ──────────────────────────────
    run_label = flow_started_at.strftime("%Y-%m-%d %H:%M:%S")
    RL.write_run_header("timing",          run_label)
    RL.write_run_header("emailFetching",   run_label)
    RL.write_run_header("grabbingImportantEmailContent", run_label)
    RL.write_run_header("tracking_hrefs",  run_label)
    RL.write_run_header("openai_extraction", run_label)
    RL.write_run_header("htmlHandler",     run_label)

    # ── Create the single OpenAI usage log for this run ─────────
    usage_log_path: Path | None = None
    try:
        usage_log_path = create_usage_log(base_dir, flow_started_at)
    except OSError as e:
        print(f"WARNING: Could not create OpenAI usage log: {e}\n")

    # ── Step 1: Environment initialization ──────────────────────
    init_s = run_environment_init()
    _emit_run_launcher_progress(5, "Environment ready")

    # ── Step 2: Fetch emails ─────────────────────────────────────
    attachments_dir = base_dir / "email_contents" / "attachments"
    pdf_dir = base_dir / "email_contents" / "pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)

    t_fetch = time.perf_counter()
    fetch_skipped = False
    if skip_email_fetch:
        fetch_skipped = True
        html_archive = base_dir / "email_contents" / "html"
        print(
            "[Step 2] Skipping Microsoft Graph — reprocessing saved HTML from:\n"
            f"         {html_archive}\n"
            f"         {pdf_dir}\n"
            "         (each source is copied to pdf/fileN.html for the extractor; originals are unchanged.)\n"
        )
        sources = _discover_local_email_html_files(base_dir)
        fetch_s = time.perf_counter() - t_fetch
        if not sources:
            print(
                "  No *.html files found in email_contents/html or email_contents/pdf.\n"
                "  Run once with mailbox fetch to download mail, or add HTML files there."
            )
            _emit_run_launcher_progress(100, "No saved HTML to process")
            return
        print(f"[Step 2] Found {len(sources)} HTML file(s) to reprocess  ({fetch_s:.2f}s)\n")
        RL.log(
            "emailFetching",
            f"{RL.ts()}  skip_graph=1  local_files={len(sources)}  time={fetch_s:.2f}s",
        )
        n_emails = len(sources)
        print(f"[Step 3] Processing {n_emails} email(s) ...\n")
        _emit_run_launcher_progress(13, f"Found {n_emails} saved email(s)")
        _rebuild_email_html_archive_folder(html_archive)
        t_process_start = time.perf_counter()
        for i, src_path in enumerate(sources, start=1):
            if i > 1 and _launcher_cancel_requested(base_dir):
                break
            dst_path = pdf_dir / f"file{i}.html"
            try:
                shutil.copy2(src_path, dst_path)
            except OSError as e:
                print(f"  ERROR: could not copy {src_path.name} -> {dst_path.name}: {e}")
                continue
            try:
                html_text = dst_path.read_text(encoding="utf-8-sig")
            except OSError as e:
                print(f"  ERROR: could not read {dst_path}: {e}")
                continue
            subj, sender_email, sender_name = _parse_saved_email_html_metadata(html_text)
            subj_display = (subj or "(no subject)")[:60]
            print(
                f"[{i}/{n_emails}] \"{subj_display}\" — {sender_name} <{sender_email}>\n"
                f"         (from {src_path.name})"
            )
            t_email = time.perf_counter()
            run_grabbing_important_content(
                html_file=f"file{i}.html",
                subject=subj,
                sender_email=sender_email,
                sender_name=sender_name,
                usage_log_path=usage_log_path,
                timing_buffer_path=timing_buffer_path,
            )
            email_elapsed = time.perf_counter() - t_email
            print(f"  done  ({email_elapsed:.2f}s)\n")
            _emit_run_launcher_progress(
                _run_pct_after_email(i, n_emails),
                f"Email {i}/{n_emails}",
            )
        process_s = time.perf_counter() - t_process_start
    elif custom_import_html:
        fetch_skipped = True
        html_archive = base_dir / "email_contents" / "html"
        custom_dir = base_dir / "custom_import_html_files"
        print(
            "[Step 2] Skipping Microsoft Graph — custom debug import from:\n"
            f"         {custom_dir}\n"
            "         (each file is copied to email_contents/pdf/fileN.html; originals are unchanged.)\n"
        )
        sources = _discover_custom_import_html_files(base_dir)
        fetch_s = time.perf_counter() - t_fetch
        if not sources:
            print(
                "  No *.html files found (folder missing or empty).\n"
                f"  Add HTML files under:\n  {custom_dir}"
            )
            _emit_run_launcher_progress(100, "No custom import HTML")
            return
        print(f"[Step 2] Found {len(sources)} HTML file(s) for custom import  ({fetch_s:.2f}s)\n")
        RL.log(
            "emailFetching",
            f"{RL.ts()}  skip_graph=1  custom_import_html_files={len(sources)}  time={fetch_s:.2f}s",
        )
        n_emails = len(sources)
        print(f"[Step 3] Processing {n_emails} file(s) (placeholder metadata: {_CUSTOM_IMPORT_LABEL}) ...\n")
        _emit_run_launcher_progress(13, f"Found {n_emails} custom import file(s)")
        _rebuild_email_html_archive_folder(html_archive)
        t_process_start = time.perf_counter()
        subj = _CUSTOM_IMPORT_LABEL
        sender_email = _CUSTOM_IMPORT_EMAIL
        sender_name = _CUSTOM_IMPORT_LABEL
        outlook_env = _custom_import_outlook_env()
        for i, src_path in enumerate(sources, start=1):
            if i > 1 and _launcher_cancel_requested(base_dir):
                break
            dst_path = pdf_dir / f"file{i}.html"
            try:
                shutil.copy2(src_path, dst_path)
            except OSError as e:
                print(f"  ERROR: could not copy {src_path.name} -> {dst_path.name}: {e}")
                continue
            print(
                f"[{i}/{n_emails}] \"{subj}\" — {sender_name} <{sender_email}>\n"
                f"         (from {src_path.name})"
            )
            t_email = time.perf_counter()
            run_grabbing_important_content(
                html_file=f"file{i}.html",
                subject=subj,
                sender_email=sender_email,
                sender_name=sender_name,
                usage_log_path=usage_log_path,
                timing_buffer_path=timing_buffer_path,
                outlook_pdf_header_env=outlook_env,
            )
            email_elapsed = time.perf_counter() - t_email
            print(f"  done  ({email_elapsed:.2f}s)\n")
            _emit_run_launcher_progress(
                _run_pct_after_email(i, n_emails),
                f"Email {i}/{n_emails}",
            )
        process_s = time.perf_counter() - t_process_start
    else:
        print(f"[Step 2] Fetching emails from folder {mail_folder!r} ...")
        if demo_mode:
            print(
                f"  [DEMO_MODE] {'device_code' if auth_flow == 'device_code' else 'interactive'} "
                "login forced.\n"
            )
        emails = fetch_emails(
            mail_folder=mail_folder,
            attachments_dir=attachments_dir,
            azure_client_id=azure_client_id,
            azure_tenant_id=azure_tenant_id,
            auth_flow=auth_flow,
            token_cache_path=token_cache_path,
            force_full_graph_auth=demo_mode,
        )
        fetch_s = time.perf_counter() - t_fetch

        if not emails:
            print("  No emails found. Nothing to process.")
            _emit_run_launcher_progress(100, "Mailbox folder empty")
            return

        print(f"[Step 2] Fetched {len(emails)} email(s)  ({fetch_s:.2f}s)\n")
        RL.log(
            "emailFetching",
            f"{RL.ts()}  folder={mail_folder!r}  tenant={azure_tenant_id}  "
            f"auth={auth_flow}  fetched={len(emails)}  time={fetch_s:.2f}s",
        )

        # ── Step 3: Process each email ───────────────────────────────
        n_emails = len(emails)
        print(f"[Step 3] Processing {n_emails} email(s) ...\n")
        _emit_run_launcher_progress(13, f"Fetched {n_emails} email(s)")
        _rebuild_email_html_archive_folder(base_dir / "email_contents" / "html")
        t_process_start = time.perf_counter()

        for i, msg in enumerate(emails, start=1):
            if i > 1 and _launcher_cancel_requested(base_dir):
                break
            subj_display = (msg.subject or "(no subject)")[:60]
            print(f"[{i}/{n_emails}] \"{subj_display}\" — {msg.sender_name} <{msg.sender_email}>")

            email_html = pdf_dir / f"file{i}.html"
            email_html.write_text(msg.body_html, encoding="utf-8")

            t_email = time.perf_counter()
            run_grabbing_important_content(
                html_file=f"file{i}.html",
                subject=msg.subject,
                sender_email=msg.sender_email,
                sender_name=msg.sender_name,
                usage_log_path=usage_log_path,
                timing_buffer_path=timing_buffer_path,
                outlook_pdf_header_env={
                    "OUTLOOK_FROM_RAW": msg.from_raw,
                    "OUTLOOK_SENT_LINE": msg.sent_line,
                    "OUTLOOK_TO_LINE": msg.to_line,
                    "OUTLOOK_HEADER_TITLE": msg.header_title,
                },
            )
            email_elapsed = time.perf_counter() - t_email
            print(f"  done  ({email_elapsed:.2f}s)\n")
            _emit_run_launcher_progress(
                _run_pct_after_email(i, n_emails),
                f"Email {i}/{n_emails}",
            )

        process_s = time.perf_counter() - t_process_start

    _emit_run_launcher_progress(85, "Sorting results…")
    # ── Step 4: Sort JSON ────────────────────────────────────────
    print("[Step 4] Sorting JSON by order number ...")
    sort_s = run_sort_json()
    print(f"[Step 4] Done  ({sort_s:.2f}s)\n")
    _emit_run_launcher_progress(92, "Creating Excel…")

    # ── Step 5: Create Excel ─────────────────────────────────────
    print("[Step 5] Creating Excel document ...")
    excel_s = run_create_excel()
    print(f"[Step 5] Done  ({excel_s:.2f}s)\n")
    _emit_run_launcher_progress(100, "Complete")

    total_s = time.perf_counter() - t_main_start

    # ── Timing summary ───────────────────────────────────────────
    timing_entries = RL.read_timing_buffer(timing_buffer_path) if timing_buffer_path else []
    try:
        if timing_buffer_path and timing_buffer_path.exists():
            timing_buffer_path.unlink()
    except OSError:
        pass

    _print_and_log_timing_summary(
        flow_started_at=flow_started_at,
        total_s=total_s,
        init_s=init_s,
        fetch_s=fetch_s,
        email_count=n_emails,
        process_s=process_s,
        sort_s=sort_s,
        excel_s=excel_s,
        timing_entries=timing_entries,
        usage_log_path=usage_log_path,
        fetch_skipped=fetch_skipped,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)
