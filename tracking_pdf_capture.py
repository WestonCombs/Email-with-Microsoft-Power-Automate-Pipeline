"""Legacy Playwright capture pipeline for package tracking PDFs.

The POD viewer now uses the assisted Chrome/CDP workflow in
``pdfCaptureFromChrome.html_capture`` so tracking sites are loaded in a visible
Chrome session and captured only after the user confirms readiness.
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import sys
import time
import traceback
from pathlib import Path
from typing import Any


STATE_FILE_NAME = "tracking_pdf_hands_free_state.json"
DEFAULT_TRACKING_CONTENT_WAIT_SEC = 5.0
DEFAULT_NETWORK_IDLE_WAIT_MS = 5000
DEFAULT_SLOW_MO_MS = 25
DEFAULT_POST_READY_DELAY_SEC = 0.2


def _project_root() -> Path:
    """Return the project root used by the rest of the email sorter."""
    from shared.project_paths import ensure_base_dir_in_environ

    return ensure_base_dir_in_environ()


def _pdf_dir() -> Path:
    """Return the existing email PDF directory."""
    path = _project_root() / "email_contents" / "pdf"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _state_path() -> Path:
    """Return the shared UI toggle state path."""
    path = _project_root() / "email_contents" / STATE_FILE_NAME
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _capture_log_path() -> Path:
    """Return the tracking PDF capture log path."""
    path = _project_root() / "email_contents" / "logs" / "tracking_pdf_capture.log"
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _append_capture_log(message: str) -> None:
    """Append one diagnostic line for capture startup and failures."""
    try:
        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with _capture_log_path().open("a", encoding="utf-8", newline="\n") as handle:
            handle.write(f"{stamp} {message}\n")
    except OSError:
        pass


def read_hands_free_capture_enabled(default: bool = False) -> bool:
    """Read the shared assisted PDF toggle state used by all viewer windows."""
    path = _state_path()
    if not path.is_file():
        return default
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default
    return bool(payload.get("enabled", default)) if isinstance(payload, dict) else default


def write_hands_free_capture_enabled(enabled: bool) -> None:
    """Persist the shared assisted PDF toggle state."""
    path = _state_path()
    payload = {"enabled": bool(enabled), "updated_at": time.time()}
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload, indent=2, ensure_ascii=True), encoding="utf-8")
    tmp.replace(path)


def _build_target_pdf_path(record: dict) -> Path:
    """Build the next convention filename with ``(n)`` collision handling."""
    tracking_number = str(record.get("tracking_number") or "").strip()
    if not tracking_number:
        raw_numbers = record.get("tracking_numbers")
        if isinstance(raw_numbers, list) and raw_numbers:
            tracking_number = str(raw_numbers[0] or "").strip()
    if tracking_number:
        from proofOfDelivery.pod_data import pod_pdf_basename

        filename = (
            pod_pdf_basename(
                record.get("company"),
                record.get("purchase_datetime"),
                tracking_number,
                record.get("carrier"),
            )
            + ".pdf"
        )
    else:
        from grabbingImportantEmailContent.grabbingImportantEmailContent import build_convention_filename

        filename = build_convention_filename(record)
    candidate = _pdf_dir() / filename
    stem = candidate.stem
    suffix = candidate.suffix or ".pdf"
    counter = 2
    while candidate.exists():
        candidate = candidate.with_name(f"{stem} ({counter}){suffix}")
        counter += 1
    return candidate


def _write_pdf_with_collision(record: dict, pdf_bytes: bytes) -> Path:
    """Write PDF bytes using the convention filename without overwriting a sibling."""
    while True:
        candidate = _build_target_pdf_path(record)
        try:
            with candidate.open("xb") as handle:
                handle.write(pdf_bytes)
            return candidate
        except FileExistsError:
            continue


def carrier_specific_wait(page: Any) -> None:
    """Placeholder for carrier-specific tracking page wait rules."""
    pass


def _candidate_browser_executables() -> list[Path]:
    """Return installed browser executables to try before bundled Chromium."""
    candidates: list[Path] = []
    raw_env_path = str(os.getenv("TRACKING_PDF_BROWSER_PATH") or "").strip()
    if raw_env_path:
        candidates.append(Path(raw_env_path).expanduser())

    try:
        from pdfCaptureFromChrome.launch_mitm_chrome import find_chrome_executable

        chrome_path = find_chrome_executable()
        if chrome_path is not None:
            candidates.append(Path(chrome_path))
    except Exception:
        pass

    program_files = Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
    program_files_x86 = Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
    local_app = Path(os.environ.get("LOCALAPPDATA", ""))
    candidates.extend(
        [
            program_files / "Google" / "Chrome" / "Application" / "chrome.exe",
            program_files_x86 / "Google" / "Chrome" / "Application" / "chrome.exe",
            local_app / "Google" / "Chrome" / "Application" / "chrome.exe",
            program_files / "Microsoft" / "Edge" / "Application" / "msedge.exe",
            program_files_x86 / "Microsoft" / "Edge" / "Application" / "msedge.exe",
            local_app / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        ]
    )

    seen: set[str] = set()
    out: list[Path] = []
    for candidate in candidates:
        try:
            resolved = candidate.expanduser().resolve()
        except OSError:
            continue
        key = str(resolved).casefold()
        if key in seen or not resolved.is_file():
            continue
        seen.add(key)
        out.append(resolved)
    return out


def _launch_browser(playwright: Any) -> Any:
    """Launch a headed browser with fallbacks for Chrome, Edge, and bundled Chromium."""
    launch_kwargs = {
        "headless": False,
        "slow_mo": DEFAULT_SLOW_MO_MS,
        "timeout": 120000,
        "args": [
            "--start-maximized",
            "--disable-dev-shm-usage",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-background-networking",
            "--disable-background-timer-throttling",
        ],
    }
    launch_errors: list[str] = []

    for browser_path in _candidate_browser_executables():
        try:
            _append_capture_log(f"Trying browser executable: {browser_path}")
            browser = playwright.chromium.launch(
                executable_path=str(browser_path),
                **launch_kwargs,
            )
            _append_capture_log(f"Launched browser executable: {browser_path}")
            return browser
        except Exception as exc:
            launch_errors.append(f"{browser_path.name}: {exc}")
            _append_capture_log(f"Browser launch failed for {browser_path}: {exc}")

    try:
        _append_capture_log("Trying bundled Playwright Chromium")
        browser = playwright.chromium.launch(**launch_kwargs)
        _append_capture_log("Launched bundled Playwright Chromium")
        return browser
    except Exception as exc:
        launch_errors.append(f"bundled chromium: {exc}")
        _append_capture_log(f"Bundled Chromium launch failed: {exc}")

    if launch_errors:
        _append_capture_log("All browser launch attempts failed: " + " | ".join(launch_errors))
    raise RuntimeError("Capture failed: Chrome is not running")


def _wait_for_tracking_content(page: Any, *, timeout_sec: float = DEFAULT_TRACKING_CONTENT_WAIT_SEC) -> None:
    """Best-effort wait for visible tracking content without carrier-specific logic."""
    deadline = time.monotonic() + timeout_sec
    keywords = (
        "tracking",
        "delivered",
        "shipped",
        "in transit",
        "out for delivery",
        "label created",
        "exception",
        "delayed",
    )
    while time.monotonic() < deadline:
        try:
            text = page.locator("body").inner_text(timeout=2000)
        except Exception:
            text = ""
        clean = " ".join(text.split())
        if len(clean) >= 120 and any(word in clean.casefold() for word in keywords):
            return
        time.sleep(0.4)


def _print_pdf_bytes(page: Any) -> bytes:
    """Print the current Chromium page to PDF bytes."""
    try:
        return page.pdf(format="Letter", print_background=True)
    except Exception:
        session = page.context.new_cdp_session(page)
        result = session.send(
            "Page.printToPDF",
            {
                "printBackground": True,
                "paperWidth": 8.5,
                "paperHeight": 11,
                "marginTop": 0.35,
                "marginBottom": 0.35,
                "marginLeft": 0.35,
                "marginRight": 0.35,
            },
        )
        return base64.b64decode(result["data"])


def capture_tracking_pdf(url: str, record: dict) -> str:
    """Capture a tracking URL to ``email_contents/pdf/`` using Playwright Chromium."""
    target_url = (url or "").strip()
    if not target_url:
        raise ValueError("Tracking URL is required")

    try:
        from shared.settings_store import apply_runtime_settings_from_json

        apply_runtime_settings_from_json()
    except Exception:
        pass

    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        from playwright.sync_api import sync_playwright
    except Exception as exc:  # pragma: no cover - depends on local install
        _append_capture_log(f"Playwright import failed: {exc}")
        raise RuntimeError("Capture failed: Chrome is not running") from exc

    browser = None
    target_pdf: Path | None = None
    try:
        with sync_playwright() as playwright:
            browser = _launch_browser(playwright)
            context = browser.new_context(
                viewport={"width": 1366, "height": 900},
                locale="en-US",
                timezone_id=os.getenv("TZ", "America/New_York"),
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
                extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
            )
            page = context.new_page()
            page.goto(target_url, wait_until="domcontentloaded", timeout=60000)
            try:
                page.wait_for_load_state("networkidle", timeout=DEFAULT_NETWORK_IDLE_WAIT_MS)
            except PlaywrightTimeoutError:
                pass
            carrier_specific_wait(page)
            _wait_for_tracking_content(page)
            page.mouse.move(320, 280)
            time.sleep(DEFAULT_POST_READY_DELAY_SEC)
            target_pdf = _write_pdf_with_collision(record, _print_pdf_bytes(page))
            context.close()
            browser.close()
            browser = None
    except RuntimeError:
        raise
    except Exception as exc:
        _append_capture_log(f"Capture failed for {target_url}: {exc}")
        if browser is None:
            raise RuntimeError("Capture failed: Chrome is not running") from exc
        raise RuntimeError(f"Capture failed: {exc}") from exc
    finally:
        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass

    return str(target_pdf)


def capture_with_retry(url: str, record: dict) -> str:
    """Placeholder for future retry/backoff logic."""
    return capture_tracking_pdf(url, record)


def process_tracking_capture(url: str, record: dict) -> str:
    """Capture, validate, and audit a tracking PDF."""
    from tracking_pdf_audit import log_tracking_pdf
    from tracking_pdf_validator import validate_pdf_with_llm

    pdf_path = capture_tracking_pdf(url, record)
    try:
        validation = validate_pdf_with_llm(pdf_path)
    except Exception as exc:
        validation = {
            "latest_tracking_info_visible": False,
            "confidence": 0,
            "status_found": "Unknown",
            "latest_update_found": None,
            "reason": f"Validation failed: {exc}",
        }
    log_tracking_pdf(pdf_path, record, validation)
    if not validation.get("latest_tracking_info_visible"):
        print("[WARN] PDF captured before latest tracking info was visible")
    return pdf_path


def _stdout_json(payload: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False))
    sys.stdout.write("\n")
    sys.stdout.flush()


def _run_cli_capture_from_stdin() -> int:
    """Run one capture request from stdin JSON for viewer subprocess launches."""
    try:
        payload = json.load(sys.stdin)
        if not isinstance(payload, dict):
            raise ValueError("Capture payload must be a JSON object")
        url = str(payload.get("url") or "").strip()
        record = payload.get("record")
        if not isinstance(record, dict):
            raise ValueError("Capture payload missing record object")
        pdf_path = process_tracking_capture(url, record)
        _stdout_json({"ok": True, "pdf_path": pdf_path})
        return 0
    except Exception as exc:
        tb = traceback.format_exc().strip().replace("\n", " | ")
        _append_capture_log(f"Capture subprocess failed: {tb}")
        _stdout_json({"ok": False, "error": str(exc)})
        return 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Legacy Playwright tracking PDF capture helper")
    parser.add_argument(
        "--stdin-json",
        action="store_true",
        help="Read {url, record} capture payload from stdin as JSON",
    )
    args = parser.parse_args(argv)
    if args.stdin_json:
        return _run_cli_capture_from_stdin()
    parser.print_help(sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
