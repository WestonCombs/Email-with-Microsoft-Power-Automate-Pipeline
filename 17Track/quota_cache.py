"""
17TRACK API quota with a simple in-process cache (refresh at most every 60 seconds).

Quota **dialogs** run only from the Email Sorter launcher session: see
:func:`quota_prefetch_gate` (start) and :func:`quota_session_end_notify` (after prefetch).

Requires ``SEVENTEEN_TRACK_API_KEY`` or ``17TRACK_API_KEY`` (same as ``trackingNumbersViewer``).

Usage (from the ``python_files`` directory — the folder name starts with a digit so
``python -m 17Track…`` is not valid)::

    python 17Track/quota_cache.py
"""

from __future__ import annotations

import json
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

_PYTHON_FILES = Path(__file__).resolve().parent.parent
if str(_PYTHON_FILES) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES))

from dotenv import load_dotenv

load_dotenv(_PYTHON_FILES / ".env")

from trackingNumbersViewer.seventeen_track_api import api_key_from_env  # noqa: E402

_GETQUOTA_URL = "https://api.17track.net/track/v2.4/getquota"

# Set by ``email_sorter_launcher`` on Run / Excel subprocesses so quota runs once before
# prefetch and once after (see ``createExcelDocument._prefetch_17track_for_excel_build``).
LAUNCHER_17TRACK_SESSION_ENV = "EMAIL_SORTER_17TRACK_QUOTA_SESSION"

# Warn when remaining credits fall to this level or below (still allow API use until 0).
LOW_QUOTA_THRESHOLD = 20

# Refresh live quota from the API at most this often (seconds).
CACHE_TTL_SEC = 60

last_quota_check = 0.0
cached_remaining: int | None = None


def get_quota(*, timeout: float = 25.0) -> dict:
    """POST ``/getquota`` — full JSON body (see 17TRACK v2.4 docs)."""
    api_key = api_key_from_env()
    if not api_key:
        raise ValueError(
            "No API key: set SEVENTEEN_TRACK_API_KEY or 17TRACK_API_KEY (e.g. in python_files/.env)."
        )

    headers = {
        "Content-Type": "application/json; charset=utf-8",
        "17token": api_key,
    }
    body = json.dumps({}).encode("utf-8")
    req = urllib.request.Request(_GETQUOTA_URL, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", errors="replace")
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {"_raw": raw, "parse_error": True}


def quota_remaining_from_payload(data: dict) -> int | None:
    """Extract remaining quota from getquota JSON (v2.4 uses ``quota_remain``)."""
    if not isinstance(data, dict):
        return None
    inner = data.get("data")
    if not isinstance(inner, dict):
        return None
    for key in ("quota_remain", "quota_remaining"):
        qr = inner.get(key)
        if qr is not None:
            try:
                return int(qr)
            except (TypeError, ValueError):
                return None
    return None


def get_cached_quota(ttl_sec: float = CACHE_TTL_SEC) -> int | None:
    """Return cached ``quota_remaining``, refreshing from the API at most every *ttl_sec*."""
    global last_quota_check, cached_remaining

    now = time.time()
    if now - last_quota_check > ttl_sec or cached_remaining is None:
        data = get_quota()
        cached_remaining = quota_remaining_from_payload(data)
        last_quota_check = now

    return cached_remaining


def fetch_quota_remaining_now(*, timeout: float = 25.0) -> int | None:
    """One live ``getquota`` call; updates :func:`get_cached_quota` globals."""
    global last_quota_check, cached_remaining
    try:
        data = get_quota(timeout=timeout)
    except ValueError:
        return None
    rem = quota_remaining_from_payload(data)
    last_quota_check = time.time()
    cached_remaining = rem
    return rem


def _notify_quota_console(title: str, body: str) -> None:
    print(f"{title}\n{body}\n", file=sys.stderr)


def _show_quota_dialog(fn, title: str, message: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        fn(title, message)
        root.destroy()
    except Exception:
        _notify_quota_console(title, message)


def notify_quota_level(remaining: int | None) -> None:
    """Tk dialogs (stderr fallback) when quota is low or exhausted."""
    if remaining is None:
        return
    if remaining <= 0:
        msg = (
            "Your 17TRACK API quota is exhausted (0 credits remaining).\n\n"
            "The tracking API cannot be used until quota is restored. Shipping status "
            "and tracking columns may not populate correctly, and the program may "
            "behave unpredictably.\n\n"
            "Contact an administrator to add quota at 17TRACK before relying on "
            "tracking updates."
        )
        try:
            from tkinter import messagebox

            _show_quota_dialog(messagebox.showerror, "17TRACK quota exhausted", msg)
        except Exception:
            _notify_quota_console("17TRACK quota exhausted", msg)
        return

    if remaining <= LOW_QUOTA_THRESHOLD:
        msg = (
            f"Your 17TRACK API quota is low.\n\n"
            f"Remaining credits: {remaining}\n\n"
            "Contact an administrator to add quota before tracking updates stop working."
        )
        try:
            from tkinter import messagebox

            _show_quota_dialog(messagebox.showerror, "17TRACK quota low", msg)
        except Exception:
            _notify_quota_console("17TRACK quota low", msg)


def quota_prefetch_gate() -> tuple[int | None, bool]:
    """Fetch quota once, notify user if low/zero.

    Returns ``(remaining, skip_network)``. When ``skip_network`` is True, callers should not
    issue 17TRACK tracking calls (quota is definitively 0).
    """
    if not api_key_from_env():
        return (None, False)
    try:
        rem = fetch_quota_remaining_now()
    except Exception:
        return (None, False)
    notify_quota_level(rem)
    skip = rem is not None and rem <= 0
    return (rem, skip)


def quota_session_end_notify() -> None:
    """Second quota check after launcher-driven 17TRACK prefetch completes (fresh getquota + dialogs)."""
    if not api_key_from_env():
        return
    try:
        rem = fetch_quota_remaining_now()
        notify_quota_level(rem)
    except Exception:
        pass


def _main() -> None:
    print("Live getquota (full JSON):")
    live = get_quota()
    print(json.dumps(live, indent=2, ensure_ascii=False))
    rem = quota_remaining_from_payload(live)
    print(f"\nquota_remaining (parsed): {rem!r}")
    print("\nCached helper (may skip network if called twice within TTL):")
    print(f"  get_cached_quota() -> {get_cached_quota()!r}")
    print(f"  get_cached_quota() -> {get_cached_quota()!r}")


if __name__ == "__main__":
    try:
        _main()
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
