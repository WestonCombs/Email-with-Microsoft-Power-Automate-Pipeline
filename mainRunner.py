"""Main runner — replaces the Microsoft Power Automate flow.

Usage:
    python mainRunner.py

Reads IMAP credentials and BASE_DIR from python_files/.env, then:
  1. Runs environment initialization (folder/file verification)
  2. Fetches all emails from the configured IMAP folder
  3. For each email: writes HTML body, runs the extraction pipeline
  4. Sorts the accumulated JSON results by order number
  5. Creates the Excel workbook
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

_PYTHON_FILES_DIR = Path(__file__).resolve().parent

from dotenv import load_dotenv

load_dotenv(_PYTHON_FILES_DIR / ".env")

from emailFetching.emailFetcher import fetch_emails

BASE_DIR_ENV = "BASE_DIR"
OPENAI_USAGE_REL = Path("email_contents") / "openai usage"


class _Tee:
    """Writes to both an original stream and a log file simultaneously."""

    def __init__(self, log_path: Path, original_stream):
        self._file = open(log_path, "a", encoding="utf-8")
        self._original = original_stream

    def write(self, msg):
        self._original.write(msg)
        self._file.write(msg.replace("\ufeff", "") if isinstance(msg, str) else msg)

    def flush(self):
        self._original.flush()
        self._file.flush()

    def close(self):
        self._file.close()


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


# ──────────────────────────────────────────────
# Step runners (each invokes its script as a subprocess so their own
# _Tee / sys.path setup works identically to standalone execution)
# ──────────────────────────────────────────────
def run_environment_init() -> None:
    script = _PYTHON_FILES_DIR / "EnvironmentInitialization" / "runner.py"
    print(f"[Step 1] Environment initialization …")
    result = subprocess.run(
        [sys.executable, str(script)],
        cwd=str(script.parent),
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"Environment initialization failed (exit code {result.returncode})"
        )
    print("[Step 1] Environment initialization complete.\n")


def run_grabbing_important_content(
    html_file: str,
    subject: str,
    sender_email: str,
    sender_name: str,
    *,
    usage_log_path: Path | None = None,
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
    result = subprocess.run(cmd, cwd=str(_PYTHON_FILES_DIR), env=env)
    if result.returncode != 0:
        print(
            f"  WARNING: grabbingImportantEmailContent exited with code {result.returncode}"
        )


def run_sort_json() -> None:
    script = _PYTHON_FILES_DIR / "sortJSONByOrderNumber" / "sortJSONByOrderNumber.py"
    print(f"\n[Step 4] Sorting JSON by order number …")
    result = subprocess.run([sys.executable, str(script)], cwd=str(_PYTHON_FILES_DIR))
    if result.returncode != 0:
        print(f"  WARNING: sortJSONByOrderNumber exited with code {result.returncode}")


def run_create_excel() -> None:
    script = _PYTHON_FILES_DIR / "createExcelDocument" / "createExcelDocument.py"
    print(f"\n[Step 5] Creating Excel document …")
    result = subprocess.run([sys.executable, str(script)], cwd=str(_PYTHON_FILES_DIR))
    if result.returncode != 0:
        print(f"  WARNING: createExcelDocument exited with code {result.returncode}")


# ──────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────
def main() -> None:
    start_time = time.time()
    flow_started_at = datetime.now()

    base_dir_raw = os.getenv(BASE_DIR_ENV)
    if not base_dir_raw:
        print(f"ERROR: {BASE_DIR_ENV} is not set in .env", file=sys.stderr)
        sys.exit(1)
    base_dir = Path(base_dir_raw).expanduser().resolve()

    imap_server = os.getenv("IMAP_SERVER")
    imap_port = int(os.getenv("IMAP_PORT", "993"))
    imap_username = os.getenv("IMAP_USERNAME")
    imap_password = os.getenv("IMAP_PASSWORD")
    imap_folder = os.getenv("IMAP_MAIL_FOLDER", "INBOX")
    imap_ssl = os.getenv("IMAP_USE_SSL", "1") != "0"

    if not all([imap_server, imap_username, imap_password]):
        print(
            "ERROR: IMAP_SERVER, IMAP_USERNAME, and IMAP_PASSWORD must be set in .env",
            file=sys.stderr,
        )
        sys.exit(1)

    print(f"{'=' * 60}")
    print(f"Main runner started: {flow_started_at.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"BASE_DIR: {base_dir}")
    print(f"Mail folder: {imap_folder}")
    print(f"{'=' * 60}\n")

    # ── Create the single OpenAI usage log for this run ─────────
    usage_log_path: Path | None = None
    try:
        usage_log_path = create_usage_log(base_dir, flow_started_at)
        print(f"OpenAI usage log: {usage_log_path}\n")
    except OSError as e:
        print(f"WARNING: Could not create OpenAI usage log: {e}\n")

    # ── Step 1: Environment initialization ──────────────────────
    run_environment_init()

    # ── Step 2: Fetch emails via IMAP ───────────────────────────
    attachments_dir = base_dir / "email_contents" / "attachments"
    pdf_dir = base_dir / "email_contents" / "pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)

    print(f"[Step 2] Fetching emails from {imap_server}:{imap_port} as {imap_username} …")
    emails = fetch_emails(
        imap_server=imap_server,
        port=imap_port,
        username=imap_username,
        password=imap_password,
        mail_folder=imap_folder,
        use_ssl=imap_ssl,
        attachments_dir=attachments_dir,
    )

    if not emails:
        print("No emails found. Nothing to process.")
        return

    print(f"[Step 2] Fetched {len(emails)} email(s).\n")

    # ── Step 3: Process each email through the extraction pipeline
    print(f"[Step 3] Processing {len(emails)} email(s) …\n")
    for i, msg in enumerate(emails, start=1):
        print(f"--- Email {i}/{len(emails)} ---")
        print(f"  From:    {msg.from_raw}")
        print(f"  Subject: {msg.subject}")
        print(f"  Sender:  {msg.sender_name} <{msg.sender_email}>")

        incoming_html = pdf_dir / "incoming.html"
        incoming_html.write_text(msg.body_html, encoding="utf-8")

        run_grabbing_important_content(
            html_file="incoming.html",
            subject=msg.subject,
            sender_email=msg.sender_email,
            sender_name=msg.sender_name,
            usage_log_path=usage_log_path,
        )
        print()

    # ── Step 4: Sort JSON ───────────────────────────────────────
    run_sort_json()

    # ── Step 5: Create Excel ────────────────────────────────────
    run_create_excel()

    elapsed = time.time() - start_time
    print(f"\n{'=' * 60}")
    print(f"Main runner finished. Total time: {elapsed:.2f}s")
    if usage_log_path:
        print_usage_summary(usage_log_path)
    print(f"{'=' * 60}")


if __name__ == "__main__":
    _base_for_log = os.getenv(BASE_DIR_ENV)

    if _base_for_log:
        _log_path = Path(_base_for_log).expanduser().resolve() / "programFileOutput.txt"
        _log_path.parent.mkdir(parents=True, exist_ok=True)
        _tee = _Tee(_log_path, sys.stdout)
        sys.stdout = _tee
        sys.stderr = _Tee(_log_path, sys.stderr)
        _original_stdout = _tee._original
        _original_stderr = sys.stderr._original

        try:
            main()
        except Exception as e:
            print(f"\nFATAL ERROR: {e}")
            import traceback
            traceback.print_exc()
            sys.stdout = _original_stdout
            sys.stderr = _original_stderr
            _tee.close()
            sys.exit(1)

        sys.stdout = _original_stdout
        sys.stderr = _original_stderr
        _tee.close()
    else:
        main()
