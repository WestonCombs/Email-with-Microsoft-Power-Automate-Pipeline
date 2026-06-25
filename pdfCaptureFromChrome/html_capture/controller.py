"""
Isolated Chrome + global Ctrl+Shift+P: snapshot the focused tab to an expected .pdf path via CDP.
"""

from __future__ import annotations

import atexit
import json
import os
import re
import subprocess
import sys
import threading
import time
import traceback
import urllib.parse
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

# pdfCaptureFromChrome/ as import root for standalone script-style imports.
_PCAP = Path(__file__).resolve().parent.parent
if str(_PCAP) not in sys.path:
    sys.path.insert(0, str(_PCAP))

try:
    from ..chrome_devtools import (
        export_page_pdf,
        extract_outer_html_snippet,
        list_page_targets,
        page_has_focus,
        reserve_free_port,
        wait_for_debugger,
    )
    from ..launch_mitm_chrome import (
        launch_isolated_chrome_no_proxy,
        terminate_isolated_capture_chrome,
    )
    from ..paths import PDF_CAPTURE_SESSION_LOG
except ImportError:
    from chrome_devtools import (  # type: ignore[no-redef]  # noqa: E402
        export_page_pdf,
        extract_outer_html_snippet,
        list_page_targets,
        page_has_focus,
        reserve_free_port,
        wait_for_debugger,
    )
    from launch_mitm_chrome import (  # type: ignore[no-redef]  # noqa: E402
        launch_isolated_chrome_no_proxy,
        terminate_isolated_capture_chrome,
    )
    from paths import PDF_CAPTURE_SESSION_LOG  # type: ignore[no-redef]  # noqa: E402

try:
    from .hotkey_win32 import CAPTURE_HOTKEY_LABEL, CaptureHotkey, hotkey_capture_available
    from .pod_readiness import (
        highlight_pod_ready_element,
        highlight_pod_selector,
        pod_readiness_debug_enabled,
        pod_selector_candidates,
        record_user_selected_ready_element,
        remove_pod_debug_overlay,
        wait_for_pod_dom_ready,
    )
except ImportError:
    from hotkey_win32 import (  # type: ignore[no-redef]  # noqa: E402
        CAPTURE_HOTKEY_LABEL,
        CaptureHotkey,
        hotkey_capture_available,
    )
    from pod_readiness import (  # type: ignore[no-redef]  # noqa: E402
        highlight_pod_ready_element,
        highlight_pod_selector,
        pod_readiness_debug_enabled,
        pod_selector_candidates,
        record_user_selected_ready_element,
        remove_pod_debug_overlay,
        wait_for_pod_dom_ready,
    )


_BACKGROUND_POD_CHROME_ARGS = (
    "--start-minimized",
    "--window-position=-32000,-32000",
    "--window-size=1280,1600",
    "--disable-features=PaintHolding",
    "--disable-session-crashed-bubble",
    "--disable-restore-session-state",
)

_LOG_LOCK = threading.Lock()

_HTTP_PREFIX_RE = re.compile(r"^https?://", re.IGNORECASE)


@dataclass
class _DebugSelectorSession:
    target_id: str
    ws_url: str
    url: str
    expected_pdf: Path
    record: dict | None
    readiness: dict
    candidates: list[dict] = field(default_factory=list)
    index: int = 0
    approve_event: threading.Event = field(default_factory=threading.Event)
    approved: bool = False
    needs_recapture: bool = False
    pdf_saved: bool = False
    final_pdf: Path | None = None
    status: str = ""


def _norm_href(s: str) -> str:
    t = (s or "").strip().rstrip("/")
    t = _HTTP_PREFIX_RE.sub("", t, count=1)
    return t.casefold()


def _log_line(message: str) -> None:
    line = f"{datetime.now().isoformat(timespec='seconds')} [html_capture] {message}\n"
    try:
        with _LOG_LOCK:
            with open(PDF_CAPTURE_SESSION_LOG, "a", encoding="utf-8", newline="\n") as f:
                f.write(line)
    except OSError:
        pass


def _terminate_chrome_process(proc: subprocess.Popen | None) -> None:
    if proc is None or proc.poll() is not None:
        return
    if sys.platform == "win32":
        kwargs: dict = {"capture_output": True}
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        subprocess.run(
            ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
            **kwargs,
        )
    else:
        proc.terminate()
        try:
            proc.wait(timeout=8)
        except subprocess.TimeoutExpired:
            proc.kill()


def _http_json_new_tab(debug_port: int, url: str) -> str | None:
    enc = urllib.parse.quote(url, safe="")
    req_url = f"http://127.0.0.1:{debug_port}/json/new?{enc}"
    try:
        with urllib.request.urlopen(req_url, timeout=60) as r:
            data = json.loads(r.read().decode("utf-8", errors="replace"))
    except OSError as e:
        _log_line(f"json/new error: {e!r} url={req_url[:200]}")
        return None
    if isinstance(data, dict):
        tid = data.get("id")
        if tid is not None:
            return str(tid)
    _log_line(f"json/new unexpected: {data!r}")
    return None


def _http_close_tab(debug_port: int, target_id: str) -> bool:
    tid = urllib.parse.quote(str(target_id), safe="")
    req_url = f"http://127.0.0.1:{debug_port}/json/close/{tid}"
    try:
        with urllib.request.urlopen(req_url, timeout=4) as r:
            r.read()
        return True
    except OSError as e:
        _log_line(f"json/close error: {e!r} target_id={target_id}")
        return False


def _http_activate_tab(debug_port: int, target_id: str) -> bool:
    tid = urllib.parse.quote(str(target_id), safe="")
    req_url = f"http://127.0.0.1:{debug_port}/json/activate/{tid}"
    try:
        with urllib.request.urlopen(req_url, timeout=4) as r:
            r.read()
        return True
    except OSError as e:
        _log_line(f"json/activate error: {e!r} target_id={target_id}")
        return False


def _first_page_target_id_for_url(debug_port: int, want_url: str) -> str | None:
    want = _norm_href(want_url)
    try:
        targets = list_page_targets(debug_port)
    except OSError as e:
        _log_line(f"list_page_targets error: {e!r}")
        return None
    best: str | None = None
    for t in targets:
        u = str(t.get("url") or "")
        tid = t.get("id")
        wu = _norm_href(u)
        if not tid or not wu or u.startswith("chrome-devtools://"):
            continue
        if want and (want in wu or wu in want or want in u.casefold() or u.casefold() in want):
            return str(tid)
        if best is None and not wu.startswith("chrome://") and wu not in ("about:blank", ""):
            best = str(tid)
    if best is not None:
        return best
    for t in targets:
        if t.get("id") and t.get("type") == "page":
            return str(t.get("id"))
    return None


def _apply_runtime_settings() -> None:
    try:
        from shared.settings_store import apply_runtime_settings_from_json

        apply_runtime_settings_from_json()
    except Exception:
        pass


def _unique_sibling_path(path: Path, marker: str) -> Path:
    suffix = path.suffix or ".pdf"
    candidate = path.with_name(f"{path.stem}{marker}{suffix}")
    if not candidate.exists():
        return candidate
    idx = 2
    while True:
        candidate = path.with_name(f"{path.stem}{marker} ({idx}){suffix}")
        if not candidate.exists():
            return candidate
        idx += 1


def _audit_unavailable(reason: str) -> bool:
    text = reason.casefold()
    return any(
        marker in text
        for marker in (
            "openai_api_key is not configured",
            "openai is required",
            "pypdf2 is required",
            "validation failed:",
        )
    )


def _finalize_pdf_with_audit(
    staged_pdf: Path,
    expected_pdf: Path,
    record: dict | None,
) -> tuple[Path, bool, str, dict]:
    if not record:
        staged_pdf.replace(expected_pdf)
        return expected_pdf, True, "", {}

    try:
        from tracking_pdf_audit import log_tracking_pdf
        from tracking_pdf_validator import validate_pdf_with_llm
    except Exception as exc:
        staged_pdf.replace(expected_pdf)
        _log_line(f"audit import unavailable: {exc!r}")
        return expected_pdf, True, f"Audit unavailable: {exc!s}", {}

    try:
        validation = validate_pdf_with_llm(str(staged_pdf))
    except Exception as exc:
        validation = {
            "latest_tracking_info_visible": False,
            "confidence": 0,
            "status_found": "Unknown",
            "latest_update_found": None,
            "reason": f"Validation failed: {exc}",
        }

    validation = validation if isinstance(validation, dict) else {}
    visible = bool(validation.get("latest_tracking_info_visible"))
    reason = str(validation.get("reason") or "").strip()
    can_advance = visible or _audit_unavailable(reason)
    final_pdf = expected_pdf if can_advance else _unique_sibling_path(expected_pdf, "_needs_review")
    staged_pdf.replace(final_pdf)

    try:
        log_tracking_pdf(str(final_pdf), record, validation)
    except Exception as exc:
        _log_line(f"audit log failed: {exc!r}")

    if visible:
        status = str(validation.get("status_found") or "Unknown").strip() or "Unknown"
        return final_pdf, True, f"AI audit passed ({status}).", validation
    if can_advance:
        return final_pdf, True, reason or "AI audit was unavailable.", validation
    return final_pdf, False, reason or "AI audit did not confirm visible tracking details.", validation


class HtmlCaptureController:
    """
    - ``start()`` registers Ctrl+Shift+P (Win32) and does not start Chrome until the first ``enqueue_capture``.
    - Each ``enqueue_capture`` opens a new tab (or the first load after a fresh launch) to ``url`` and
      records the target id → expected output ``.pdf`` path.
    - On Ctrl+Shift+P, the focused page is printed to PDF and written to the mapped path.
    - With ``auto_print_pdf=True``, a background thread watches DOM/CSS readiness, then prints
      to PDF without requiring Ctrl+Shift+P and writes an optional HTML snapshot beside the PDF.
    """

    def __init__(
        self,
        *,
        on_notify: Callable[[str, str], None] | None = None,
        on_saved: Callable[[], None] | None = None,
        on_review_saved: Callable[[dict], None] | None = None,
        on_selector_ready: Callable[[str, dict], None] | None = None,
        verbose: bool = False,
    ) -> None:
        self._on_notify = on_notify
        self._on_saved = on_saved
        self._on_review_saved = on_review_saved
        self._on_selector_ready = on_selector_ready
        self._verbose = verbose
        self._lock = threading.Lock()
        self._active = False
        self._chrome: subprocess.Popen | None = None
        self._debug_port: int = 0
        self._first_enqueued: bool = False
        self._target_to_path: dict[str, Path] = {}
        self._target_to_record: dict[str, dict] = {}
        self._target_to_url: dict[str, str] = {}
        self._order: list[str] = []
        self._hotkey: CaptureHotkey | None = None
        self._debug_sessions: dict[str, _DebugSelectorSession] = {}
        atexit.register(self.stop)

    @staticmethod
    def _env_debug_port() -> int:
        _apply_runtime_settings()
        raw = (os.environ.get("HTML_CAPTURE_DEBUG_PORT") or "").strip()
        if raw:
            try:
                p = int(raw)
                if 0 < p < 65536:
                    return p
            except ValueError:
                pass
        return reserve_free_port()

    def _emit(self, level: str, message: str) -> None:
        _log_line(f"{level} {message}")
        if self._on_notify is not None:
            try:
                self._on_notify(level, message)
            except Exception:
                pass

    def _emit_review_saved(
        self,
        *,
        final_pdf: Path,
        record: dict | None,
        audit_message: str,
        validation: dict,
    ) -> None:
        if self._on_review_saved is None:
            return
        payload = {
            "path": str(final_pdf),
            "filename": final_pdf.name,
            "record": dict(record or {}),
            "reason": audit_message,
            "validation": validation if isinstance(validation, dict) else {},
        }
        try:
            self._on_review_saved(payload)
        except Exception:
            pass

    def start(self) -> bool:
        with self._lock:
            if self._active:
                return True
            if not hotkey_capture_available():
                self._emit("error", f"HTML capture ({CAPTURE_HOTKEY_LABEL}) requires Windows.")
                return False
            try:
                terminate_isolated_capture_chrome()
            except Exception as exc:
                _log_line(f"stale chrome cleanup failed on start: {exc!r}")
            self._debug_port = self._env_debug_port()
            self._target_to_path.clear()
            self._target_to_record.clear()
            self._target_to_url.clear()
            self._order.clear()
            self._debug_sessions.clear()
            self._first_enqueued = False
            self._hotkey = CaptureHotkey(self._schedule_snapshot)
            ok = self._hotkey.start() if self._hotkey is not None else False
            if not ok and self._hotkey is not None and self._hotkey.start_error:
                self._emit("error", self._hotkey.start_error)
            if not ok:
                self._hotkey = None
                return False
            self._active = True
            if self._verbose:
                _log_line(f"started (debug port {self._debug_port})")
            return True

    def stop(self) -> None:
        with self._lock:
            h = self._hotkey
            c = self._chrome
            self._active = False
            self._hotkey = None
            self._first_enqueued = False
            self._target_to_path.clear()
            self._target_to_record.clear()
            self._target_to_url.clear()
            self._order.clear()
            self._debug_sessions.clear()
            self._chrome = None
        if h is not None:
            h.stop()
        if c is not None:
            _terminate_chrome_process(c)
        try:
            terminate_isolated_capture_chrome()
        except Exception as exc:
            _log_line(f"stale chrome cleanup failed on stop: {exc!r}")

    def _find_existing_target_id(self, debug_port: int, want_url: str, expected_pdf: Path) -> str | None:
        try:
            targets = list_page_targets(debug_port)
        except OSError as exc:
            _log_line(f"existing-target lookup failed: {exc!r}")
            return None

        want = _norm_href(want_url)
        with self._lock:
            known_paths = dict(self._target_to_path)

        url_fallback: str | None = None
        for target in targets:
            tid = str(target.get("id") or "")
            if not tid:
                continue
            if known_paths.get(tid) == expected_pdf:
                return tid
            current_url = str(target.get("url") or "")
            current_norm = _norm_href(current_url)
            if not want or not current_norm:
                continue
            if want in current_norm or current_norm in want or want in current_url.casefold():
                if tid in known_paths:
                    return tid
                if url_fallback is None:
                    url_fallback = tid
        return url_fallback

    def enqueue_capture(
        self,
        url: str,
        expected_pdf: Path,
        *,
        record: dict | None = None,
        auto_print_pdf: bool = False,
        visible_chrome: bool = False,
    ) -> bool:
        if not self._active:
            self._emit("error", "HTML capture is not started (turn Assisted PDF Capture on first).")
            return False
        u = (url or "").strip()
        if not u:
            return False
        with self._lock:
            dport = self._debug_port
        if not dport:
            return False

        with self._lock:
            ch = self._chrome

        existing_target_id: str | None = None
        if ch is not None and ch.poll() is None:
            existing_target_id = self._find_existing_target_id(dport, u, expected_pdf)

        if existing_target_id:
            target_id = existing_target_id
            _http_activate_tab(dport, target_id)
        elif ch is None or ch.poll() is not None:
            if ch is not None:
                _terminate_chrome_process(ch)
            with self._lock:
                self._chrome = None
                self._first_enqueued = False
            proc = launch_isolated_chrome_no_proxy(
                start_url=u,
                remote_debugging_port=dport,
                extra_args=(
                    _BACKGROUND_POD_CHROME_ARGS
                    if auto_print_pdf and not visible_chrome and not pod_readiness_debug_enabled()
                    else None
                ),
                verbose=self._verbose,
            )
            if proc is None:
                self._emit("error", "Could not start isolated Chrome. Is Google Chrome installed?")
                return False
            with self._lock:
                self._chrome = proc
            if not wait_for_debugger(dport, timeout=15.0):
                with self._lock:
                    c2 = self._chrome
                _terminate_chrome_process(c2)
                with self._lock:
                    self._chrome = None
                self._emit("error", "Chrome DevTools did not start (check HTML_CAPTURE_DEBUG_PORT in .env for a free port).")
                return False
            time.sleep(0.8)
            target_id = _first_page_target_id_for_url(dport, u)
            if not target_id:
                self._emit("error", "Could not find a DevTools page target for the new Chrome tab.")
                return False
        else:
            target_id = _http_json_new_tab(dport, u) or _first_page_target_id_for_url(dport, u)
            if not target_id:
                self._emit("error", "Failed to open a new tab for the tracking page.")
                return False

        with self._lock:
            self._target_to_path[str(target_id)] = expected_pdf
            self._target_to_url[str(target_id)] = u
            if record is not None:
                self._target_to_record[str(target_id)] = dict(record)
            else:
                self._target_to_record.pop(str(target_id), None)
            if str(target_id) not in self._order:
                self._order.append(str(target_id))
            if not self._first_enqueued:
                self._first_enqueued = True
        if existing_target_id:
            _log_line(f"reused target_id={target_id} -> {expected_pdf} url={u[:120]}")
        else:
            _log_line(f"enqueued target_id={target_id} -> {expected_pdf} url={u[:120]}")
        if self._verbose:
            _log_line("enqueue: ok")
        if auto_print_pdf:
            tid_s = str(target_id)
            pdf_p = Path(expected_pdf)

            def _run() -> None:
                self._thread_auto_pod_capture(tid_s, pdf_p)

            threading.Thread(
                target=_run,
                name="html-capture-auto-pod",
                daemon=True,
            ).start()
        return True

    def _websocket_for_target_id(self, target_id: str) -> str | None:
        with self._lock:
            dport = self._debug_port
        if not dport:
            return None
        try:
            for t in list_page_targets(dport):
                if str(t.get("id") or "") != target_id:
                    continue
                w = t.get("webSocketDebuggerUrl")
                if isinstance(w, str) and w:
                    return w
        except OSError:
            return None
        return None

    def _close_target_id(self, target_id: str) -> None:
        with self._lock:
            dport = self._debug_port
            self._target_to_path.pop(target_id, None)
            self._target_to_record.pop(target_id, None)
            self._target_to_url.pop(target_id, None)
            self._debug_sessions.pop(target_id, None)
            self._order = [tid for tid in self._order if tid != target_id]
        if dport:
            _http_close_tab(dport, target_id)

    def close_debug_target(self, target_id: str) -> None:
        """Close a supervised debug tab after its PDF has been saved."""
        self._close_target_id(str(target_id))

    def _debug_session_status_locked(self, session: _DebugSelectorSession) -> dict:
        total = len(session.candidates)
        current = session.candidates[session.index] if total else {}
        return {
            "target_id": session.target_id,
            "index": session.index,
            "total": total,
            "current": dict(current),
            "approved": session.approved,
            "needs_recapture": session.needs_recapture,
            "pdf_saved": session.pdf_saved,
            "final_pdf": str(session.final_pdf) if session.final_pdf is not None else "",
            "status": session.status,
            "readiness": dict(session.readiness),
        }

    def debug_selector_status(self, target_id: str) -> dict | None:
        with self._lock:
            session = self._debug_sessions.get(str(target_id))
            if session is None:
                return None
            return self._debug_session_status_locked(session)

    def select_debug_element(self, target_id: str, step: int) -> dict | None:
        """Move the blue POD selector box by *step* and return the new selection status."""
        tid = str(target_id)
        with self._lock:
            session = self._debug_sessions.get(tid)
            if session is None:
                return None
            if session.candidates:
                session.index = (session.index + int(step)) % len(session.candidates)
                current = dict(session.candidates[session.index])
            else:
                current = {}
            ws_url = session.ws_url
            url = session.url
            record = dict(session.record) if isinstance(session.record, dict) else None
        selector = str(current.get("selector") or "").strip()
        if selector:
            highlight_pod_selector(ws_url, selector=selector, url=url, record=record)
        with self._lock:
            session = self._debug_sessions.get(tid)
            if session is None:
                return None
            session.status = f"Selected element {session.index + 1} of {len(session.candidates)}"
            return self._debug_session_status_locked(session)

    def approve_debug_element(self, target_id: str) -> bool:
        """Accept the currently highlighted selector and let the waiting capture thread print."""
        tid = str(target_id)
        with self._lock:
            session = self._debug_sessions.get(tid)
            if session is None or session.pdf_saved:
                return False
            if session.needs_recapture:
                session.approved = True
                session.needs_recapture = False
                session.status = "Selection accepted. Saving PDF again..."
                threading.Thread(
                    target=self._thread_debug_recapture,
                    args=(tid,),
                    name="html-capture-debug-retry",
                    daemon=True,
                ).start()
                return True
            session.approved = True
            session.status = "Selection accepted. Saving PDF..."
            session.approve_event.set()
            return True

    def _prepare_debug_selector_session(
        self,
        *,
        target_id: str,
        ws_url: str,
        expected_pdf: Path,
        record: dict | None,
        readiness: dict,
    ) -> _DebugSelectorSession | None:
        with self._lock:
            url = self._target_to_url.get(target_id, "")
        candidates = pod_selector_candidates(ws_url, url=url, record=record, readiness=readiness)
        if not candidates:
            ready_element = readiness.get("ready_element") if isinstance(readiness, dict) else None
            if isinstance(ready_element, dict) and ready_element.get("selector"):
                candidates = [dict(ready_element)]
        highlight_selector = str((candidates[0] if candidates else {}).get("selector") or "").strip()
        if highlight_selector:
            highlight_pod_selector(ws_url, selector=highlight_selector, url=url, record=record)
        else:
            highlight_pod_ready_element(ws_url, url=url, record=record, readiness=readiness)

        session = _DebugSelectorSession(
            target_id=target_id,
            ws_url=ws_url,
            url=url,
            expected_pdf=Path(expected_pdf),
            record=dict(record) if isinstance(record, dict) else None,
            readiness=dict(readiness),
            candidates=candidates,
            index=0,
            status="Choose the ready element, then capture.",
        )
        with self._lock:
            if not self._active:
                return None
            self._debug_sessions[target_id] = session
            status = self._debug_session_status_locked(session)
        if self._on_selector_ready is not None:
            try:
                self._on_selector_ready(target_id, status)
            except Exception:
                pass
        self._emit(
            "selector",
            "POD selector debug is ready. Use the selector controls to choose the page element.",
        )
        return session

    def _wait_for_debug_selector_approval(self, session: _DebugSelectorSession) -> dict | None:
        while True:
            with self._lock:
                active = self._active
                current = self._debug_sessions.get(session.target_id)
            if not active or current is None:
                return None
            if session.approve_event.wait(timeout=0.25):
                break
        with self._lock:
            current = self._debug_sessions.get(session.target_id)
            if current is None:
                return None
            if current.candidates:
                selected = dict(current.candidates[current.index])
            else:
                ready_element = current.readiness.get("ready_element")
                selected = dict(ready_element) if isinstance(ready_element, dict) else {}
            current.status = "Selection accepted. Waiting 1 second before PDF capture..."
        if selected:
            record_user_selected_ready_element(
                url=session.url,
                record=session.record,
                selected_element=selected,
                elapsed_seconds=float(session.readiness.get("elapsed_seconds") or 0.0),
            )
            highlight_pod_selector(
                session.ws_url,
                selector=str(selected.get("selector") or ""),
                url=session.url,
                record=session.record,
            )
        return selected

    def _thread_debug_recapture(self, target_id: str) -> None:
        try:
            with self._lock:
                session = self._debug_sessions.get(target_id)
                if session is None:
                    return
                ws_url = session.ws_url
                expected_pdf = Path(session.expected_pdf)
                record = dict(session.record) if isinstance(session.record, dict) else None
                url = session.url
                if session.candidates:
                    selected = dict(session.candidates[session.index])
                else:
                    selected = {}
            if selected:
                record_user_selected_ready_element(
                    url=url,
                    record=record,
                    selected_element=selected,
                    elapsed_seconds=None,
                )
                highlight_pod_selector(
                    ws_url,
                    selector=str(selected.get("selector") or ""),
                    url=url,
                    record=record,
                )
            self._emit("progress", "Selection accepted. Waiting 1 second, then saving PDF again...")
            time.sleep(1.0)

            if expected_pdf.is_file() and not self._verbose:
                with self._lock:
                    session = self._debug_sessions.get(target_id)
                    if session is not None:
                        session.pdf_saved = True
                        session.final_pdf = expected_pdf
                        session.status = "PDF already exists. Use Next POD when ready."
                if self._on_saved is not None:
                    try:
                        self._on_saved()
                    except Exception:
                        pass
                return

            expected_pdf.parent.mkdir(parents=True, exist_ok=True)
            remove_pod_debug_overlay(ws_url)
            pdf_bytes = export_page_pdf(ws_url)
            staged_pdf = _unique_sibling_path(expected_pdf, "_capture_pending")
            staged_pdf.write_bytes(pdf_bytes)
            final_pdf, can_advance, audit_message, validation = _finalize_pdf_with_audit(
                staged_pdf,
                expected_pdf,
                record,
            )
            _log_line(f"debug re-saved PDF {final_pdf} ({len(pdf_bytes)} bytes)")

            html_path = final_pdf.with_name(final_pdf.stem + "_capture.html")
            html_snip = extract_outer_html_snippet(ws_url, max_chars=350_000)
            if html_snip:
                try:
                    html_path.write_text(html_snip, encoding="utf-8", errors="replace")
                    _log_line(f"debug re-saved HTML snapshot {html_path}")
                except OSError as exc:
                    _log_line(f"debug html snapshot write failed: {exc!r}")

            with self._lock:
                session = self._debug_sessions.get(target_id)
                if session is not None:
                    if can_advance:
                        session.pdf_saved = True
                        session.final_pdf = final_pdf
                        session.status = "PDF saved. Use Next POD when ready."
                    else:
                        session.pdf_saved = True
                        session.final_pdf = final_pdf
                        session.status = "PDF saved for manual review. Use Next POD when ready."

            extra = f"\nHTML snapshot:\n{html_path.name}" if html_snip else ""
            if can_advance:
                if self._on_saved is not None:
                    try:
                        self._on_saved()
                    except Exception:
                        pass
                audit_line = f"\n{audit_message}" if audit_message else ""
                self._emit("info", f"Proof-of-delivery PDF saved:\n{final_pdf.name}{extra}{audit_line}")
            else:
                self._emit_review_saved(
                    final_pdf=final_pdf,
                    record=record,
                    audit_message=audit_message,
                    validation=validation,
                )
                self._emit(
                    "info",
                    "POD PDF saved for manual review:\n"
                    f"{final_pdf.name}{extra}\n\n{audit_message}",
                )
        except Exception as e:
            _log_line("debug recapture error:\n" + traceback.format_exc())
            with self._lock:
                session = self._debug_sessions.get(target_id)
                if session is not None:
                    session.approved = False
                    session.needs_recapture = True
                    session.status = "Recapture failed. Adjust selection and try again."
            self._emit("error", f"Automatic print to PDF failed: {e!s}")

    def _thread_auto_pod_capture(self, target_id: str, expected_pdf: Path) -> None:
        self._emit(
            "info",
            "Capture: opened carrier tab - watching page elements before saving PDF...",
        )
        try:
            ws_url = self._websocket_for_target_id(target_id)
            if not ws_url:
                self._emit("error", "Lost the capture tab before reading the page.")
                return
            with self._lock:
                record = self._target_to_record.get(target_id)
                target_url = self._target_to_url.get(target_id, "")

            readiness = wait_for_pod_dom_ready(
                ws_url,
                url=target_url,
                record=record,
                timeout_seconds=75.0,
                notify=self._emit,
            )
            _log_line(
                "pod readiness "
                f"target_id={target_id} mode={readiness.get('mode')} "
                f"elapsed={readiness.get('elapsed_seconds')} "
                f"selector={readiness.get('ready_selector')}"
            )

            debug_selector = pod_readiness_debug_enabled()
            if debug_selector:
                session = self._prepare_debug_selector_session(
                    target_id=target_id,
                    ws_url=ws_url,
                    expected_pdf=expected_pdf,
                    record=record,
                    readiness=readiness,
                )
                if session is None:
                    return
                selected = self._wait_for_debug_selector_approval(session)
                if selected is None:
                    return
                self._emit("progress", "Selection accepted. Waiting 1 second, then saving PDF...")
                time.sleep(1.0)
            else:
                self._emit("progress", "Page looks ready. Saving POD PDF automatically...")

            if expected_pdf.is_file() and not self._verbose:
                self._emit("info", f"File already exists:\n{expected_pdf.name}")
                if debug_selector:
                    with self._lock:
                        session = self._debug_sessions.get(target_id)
                        if session is not None:
                            session.pdf_saved = True
                            session.final_pdf = expected_pdf
                            session.status = "PDF already exists. Use Next POD when ready."
                else:
                    self._close_target_id(target_id)
                if self._on_saved is not None:
                    try:
                        self._on_saved()
                    except Exception:
                        pass
                return
            expected_pdf.parent.mkdir(parents=True, exist_ok=True)

            remove_pod_debug_overlay(ws_url)
            pdf_bytes = export_page_pdf(ws_url)
            staged_pdf = _unique_sibling_path(expected_pdf, "_capture_pending")
            staged_pdf.write_bytes(pdf_bytes)
            final_pdf, can_advance, audit_message, validation = _finalize_pdf_with_audit(
                staged_pdf,
                expected_pdf,
                record,
            )
            _log_line(f"auto-saved PDF {final_pdf} ({len(pdf_bytes)} bytes)")

            html_path = final_pdf.with_name(final_pdf.stem + "_capture.html")
            html_snip = extract_outer_html_snippet(ws_url, max_chars=350_000)
            if html_snip:
                try:
                    html_path.write_text(html_snip, encoding="utf-8", errors="replace")
                    _log_line(f"auto-saved HTML snapshot {html_path}")
                except OSError as exc:
                    _log_line(f"html snapshot write failed: {exc!r}")

            if can_advance:
                if debug_selector:
                    with self._lock:
                        session = self._debug_sessions.get(target_id)
                        if session is not None:
                            session.pdf_saved = True
                            session.final_pdf = final_pdf
                            session.status = "PDF saved. Use Next POD when ready."
                else:
                    self._close_target_id(target_id)
                if self._on_saved is not None:
                    try:
                        self._on_saved()
                    except Exception:
                        pass
            elif debug_selector:
                with self._lock:
                    session = self._debug_sessions.get(target_id)
                    if session is not None:
                        session.pdf_saved = True
                        session.final_pdf = final_pdf
                        session.status = "PDF saved for manual review. Use Next POD when ready."
            else:
                self._close_target_id(target_id)
            extra = f"\nHTML snapshot:\n{html_path.name}" if html_snip else ""
            if can_advance:
                audit_line = f"\n{audit_message}" if audit_message else ""
                self._emit("info", f"Proof-of-delivery PDF saved:\n{final_pdf.name}{extra}{audit_line}")
            else:
                self._emit_review_saved(
                    final_pdf=final_pdf,
                    record=record,
                    audit_message=audit_message,
                    validation=validation,
                )
                self._emit(
                    "info",
                    "POD PDF saved for manual review:\n"
                    f"{final_pdf.name}{extra}\n\n{audit_message}",
                )
        except Exception as e:
            _log_line("auto pod capture error:\n" + traceback.format_exc())
            self._emit("error", f"Automatic print to PDF failed: {e!s}")

    def _schedule_snapshot(self) -> None:
        self._emit("progress", f"{CAPTURE_HOTKEY_LABEL} received. Starting PDF capture...")
        threading.Thread(target=self._do_snapshot, name="html-capture-cdp", daemon=True).start()

    def _resolve_capture_for_focused_tab(self) -> tuple[str, Path] | None:
        with self._lock:
            dport = self._debug_port
            tmap = dict(self._target_to_path)
            order = list(self._order)
        if not dport or not tmap or not self._active:
            return None
        try:
            targets = list_page_targets(dport)
        except OSError as e:
            _log_line(f"snapshot: list_page_targets: {e!r}")
            return None
        current_ids = {str(t.get("id") or "") for t in targets if t.get("id")}
        for t in targets:
            ws = t.get("webSocketDebuggerUrl")
            tid = str(t.get("id") or "")
            if not isinstance(ws, str) or not tid or tid not in tmap:
                continue
            if page_has_focus(ws):
                return tid, tmap[tid]
        for tid in reversed(order):
            if tid in tmap and tid in current_ids:
                return tid, tmap[tid]
        if order and order[-1] in tmap and order[-1] in current_ids:
            return order[-1], tmap[order[-1]]
        return None

    def _websocket_for_path(self, out_path: Path) -> str | None:
        with self._lock:
            dport = self._debug_port
            tmap = dict(self._target_to_path)
        if not dport:
            return None
        want_ids = {k for k, p in tmap.items() if p == out_path}
        if not want_ids:
            return None
        try:
            targets = list_page_targets(dport)
        except OSError as e:
            _log_line(f"websocket For path: {e!r}")
            return None
        for t in targets:
            tid = str(t.get("id") or "")
            w = t.get("webSocketDebuggerUrl")
            if tid in want_ids and isinstance(w, str) and w and page_has_focus(w):
                return w
        for t in targets:
            tid = str(t.get("id") or "")
            w = t.get("webSocketDebuggerUrl")
            if tid in want_ids and isinstance(w, str) and w:
                return w
        return None

    def _do_snapshot(self) -> None:
        try:
            with self._lock:
                dport = self._debug_port
                c = self._chrome
            if c is not None and c.poll() is not None:
                self._emit("info", "The capture Chrome was closed. Double-click a row to reopen it.")
                with self._lock:
                    self._chrome = None
                    self._first_enqueued = False
                    self._target_to_path.clear()
                    self._target_to_record.clear()
                    self._target_to_url.clear()
                    self._order.clear()
                    self._debug_sessions.clear()
                return

            capture = self._resolve_capture_for_focused_tab()
            if capture is None:
                self._emit(
                    "error",
                    f"No matching capture tab. Double-click a tracking row, then try {CAPTURE_HOTKEY_LABEL} in Chrome.",
                )
                return
            target_id, out_path = capture
            self._emit("progress", "Matched the focused Chrome tab. Preparing to print the page...")
            if out_path.is_file() and not self._verbose:
                self._emit("info", f"File already exists:\n{out_path.name}")
                self._close_target_id(target_id)
                if self._on_saved is not None:
                    try:
                        self._on_saved()
                    except Exception:
                        pass
                return
            out_path.parent.mkdir(parents=True, exist_ok=True)

            ws_url = self._websocket_for_path(out_path)
            if not ws_url and dport:
                try:
                    for t in list_page_targets(dport):
                        tid = str(t.get("id") or "")
                        w = t.get("webSocketDebuggerUrl")
                        with self._lock:
                            tmap = dict(self._target_to_path)
                        if (
                            isinstance(w, str)
                            and w
                            and tid
                            and tid in tmap
                            and tmap.get(tid) == out_path
                        ):
                            ws_url = w
                            break
                except OSError:
                    pass
            if not ws_url:
                self._emit("error", "Could not find the Chrome tab to print (is DevTools connected?).")
                return

            self._emit("progress", "Printing the displayed shipping page to PDF...")
            remove_pod_debug_overlay(ws_url)
            pdf_bytes = export_page_pdf(ws_url)
            staged_pdf = _unique_sibling_path(out_path, "_capture_pending")
            staged_pdf.write_bytes(pdf_bytes)
            with self._lock:
                record = self._target_to_record.get(target_id)
            self._emit("progress", "Checking the captured PDF before saving...")
            final_pdf, can_advance, audit_message, validation = _finalize_pdf_with_audit(
                staged_pdf,
                out_path,
                record,
            )
            _log_line(f"saved PDF {final_pdf} ({len(pdf_bytes)} bytes)")
            if can_advance:
                self._close_target_id(target_id)
                if self._on_saved is not None:
                    try:
                        self._on_saved()
                    except Exception:
                        pass
                audit_line = f"\n{audit_message}" if audit_message else ""
                self._emit("info", f"Proof-of-delivery PDF saved:\n{final_pdf.name}{audit_line}")
            else:
                self._close_target_id(target_id)
                self._emit_review_saved(
                    final_pdf=final_pdf,
                    record=record,
                    audit_message=audit_message,
                    validation=validation,
                )
                self._emit(
                    "info",
                    "POD PDF saved for manual review:\n"
                    f"{final_pdf.name}\n\n{audit_message}",
                )
        except Exception as e:
            _log_line("snapshot error:\n" + traceback.format_exc())
            self._emit("error", f"Print to PDF failed: {e!s}")
