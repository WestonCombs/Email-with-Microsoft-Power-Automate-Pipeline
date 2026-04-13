"""Main runner — replaces the Microsoft Power Automate flow.

Usage:
    python mainRunner.py

Reads Microsoft Graph / OpenAI settings and BASE_DIR from python_files/.env, then:
  1. Runs environment initialization (folder/file verification)
  2. Fetches all emails from the configured mailbox folder (Microsoft Graph, OAuth 2.0)
  3. For each email: writes HTML body, runs the extraction pipeline
  4. Sorts the accumulated JSON results by order number
  5. Creates the Excel workbook
"""

from __future__ import annotations

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
            f"WARNING: {_env_path.name} not found. Copy .env.example to .env and set BASE_DIR and other values.",
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

import runLogger as RL
from emailFetching.emailFetcher import fetch_emails, prepend_outlook_style_header

BASE_DIR_ENV = "BASE_DIR"
OPENAI_USAGE_REL = Path("logs") / "openai usage"


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
    """Empty archived HTML folder before a run that will create PDFs (non-duplicates only)."""
    html_dir.mkdir(parents=True, exist_ok=True)
    for child in list(html_dir.iterdir()):
        try:
            if child.is_file():
                child.unlink()
            elif child.is_dir():
                shutil.rmtree(child)
        except OSError as e:
            print(f"WARNING: could not remove {child}: {e}")


# ──────────────────────────────────────────────
# Step runners (each invokes its script as a subprocess so cwd and __main__
# match standalone execution)
# ──────────────────────────────────────────────
def run_environment_init() -> float:
    """Returns elapsed seconds."""
    script = _PYTHON_FILES_DIR / "environmentInitialization" / "runner.py"
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
    lines: list[str] = [
        "",
        sep,
        f"  Run complete: {flow_started_at.strftime('%Y-%m-%d %H:%M:%S')}",
        sep,
        "",
        f"  Total Program Duration: {_fmt(total_s)}",
        "",
        f"  Environment init:      {_fmt(init_s)}",
        f"  Email fetching:        {_fmt(fetch_s)}  ({email_count} email(s) retrieved)",
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

    base_dir_raw = os.getenv(BASE_DIR_ENV)
    if not base_dir_raw:
        print(f"ERROR: {BASE_DIR_ENV} is not set in .env", file=sys.stderr)
        sys.exit(1)
    base_dir = Path(base_dir_raw).expanduser().resolve()

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

    if not azure_client_id:
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

    # ── Step 2: Fetch emails ─────────────────────────────────────
    attachments_dir = base_dir / "email_contents" / "attachments"
    pdf_dir = base_dir / "email_contents" / "pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)

    print(f"[Step 2] Fetching emails from folder {mail_folder!r} ...")
    if demo_mode:
        print(f"  [DEMO_MODE] {'device_code' if auth_flow=='device_code' else 'interactive'} login forced.\n")
    t_fetch = time.perf_counter()
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
        return

    print(f"[Step 2] Fetched {len(emails)} email(s)  ({fetch_s:.2f}s)\n")
    RL.log("emailFetching",
        f"{RL.ts()}  folder={mail_folder!r}  tenant={azure_tenant_id}  "
        f"auth={auth_flow}  fetched={len(emails)}  time={fetch_s:.2f}s"
    )

    # ── Step 3: Process each email ───────────────────────────────
    n_emails = len(emails)
    print(f"[Step 3] Processing {n_emails} email(s) ...\n")
    _rebuild_email_html_archive_folder(base_dir / "email_contents" / "html")
    t_process_start = time.perf_counter()

    for i, msg in enumerate(emails, start=1):
        subj_display = (msg.subject or "(no subject)")[:60]
        print(f"[{i}/{n_emails}] \"{subj_display}\" — {msg.sender_name} <{msg.sender_email}>")

        email_html = pdf_dir / f"file{i}.html"
        email_html.write_text(
            prepend_outlook_style_header(msg.body_html, msg),
            encoding="utf-8",
        )

        t_email = time.perf_counter()
        run_grabbing_important_content(
            html_file=f"file{i}.html",
            subject=msg.subject,
            sender_email=msg.sender_email,
            sender_name=msg.sender_name,
            usage_log_path=usage_log_path,
            timing_buffer_path=timing_buffer_path,
        )
        email_elapsed = time.perf_counter() - t_email
        print(f"  done  ({email_elapsed:.2f}s)\n")

    process_s = time.perf_counter() - t_process_start

    # ── Step 4: Sort JSON ────────────────────────────────────────
    print("[Step 4] Sorting JSON by order number ...")
    sort_s = run_sort_json()
    print(f"[Step 4] Done  ({sort_s:.2f}s)\n")

    # ── Step 5: Create Excel ─────────────────────────────────────
    print("[Step 5] Creating Excel document ...")
    excel_s = run_create_excel()
    print(f"[Step 5] Done  ({excel_s:.2f}s)\n")

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
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)
