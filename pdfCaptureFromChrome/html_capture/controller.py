"""
Isolated Chrome + global Ctrl+Enter: snapshot the focused tab to an expected .pdf path via CDP.
"""

from __future__ import annotations

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
        inspect_page,
        list_page_targets,
        page_has_focus,
        reserve_free_port,
        wait_for_debugger,
    )
    from ..launch_mitm_chrome import launch_isolated_chrome_no_proxy
    from ..paths import PDF_CAPTURE_SESSION_LOG
except ImportError:
    from chrome_devtools import (  # type: ignore[no-redef]  # noqa: E402
        export_page_pdf,
        extract_outer_html_snippet,
        inspect_page,
        list_page_targets,
        page_has_focus,
        reserve_free_port,
        wait_for_debugger,
    )
    from launch_mitm_chrome import launch_isolated_chrome_no_proxy  # type: ignore[no-redef]  # noqa: E402
    from paths import PDF_CAPTURE_SESSION_LOG  # type: ignore[no-redef]  # noqa: E402

try:
    from .hotkey_win32 import CtrlEnterHotkey, hotkey_ctrl_enter_available
except ImportError:
    from hotkey_win32 import CtrlEnterHotkey, hotkey_ctrl_enter_available  # type: ignore[no-redef]  # noqa: E402

_LOG_LOCK = threading.Lock()

_HTTP_PREFIX_RE = re.compile(r"^https?://", re.IGNORECASE)


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


class HtmlCaptureController:
    """
    - ``start()`` registers Ctrl+Enter (Win32) and does not start Chrome until the first ``enqueue_capture``.
    - Each ``enqueue_capture`` opens a new tab (or the first load after a fresh launch) to ``url`` and
      records the target id → expected output ``.pdf`` path.
    - On Ctrl+Enter, the focused page is printed to PDF and written to the mapped path.
    - With ``auto_print_pdf=True``, a background thread waits for the tab to settle (CDP ``inspect_page``),
      then prints to PDF without requiring Ctrl+Enter, and writes an optional HTML snapshot beside the PDF.
    """

    def __init__(
        self,
        *,
        on_notify: Callable[[str, str], None] | None = None,
        on_saved: Callable[[], None] | None = None,
        verbose: bool = False,
    ) -> None:
        self._on_notify = on_notify
        self._on_saved = on_saved
        self._verbose = verbose
        self._lock = threading.Lock()
        self._active = False
        self._chrome: subprocess.Popen | None = None
        self._debug_port: int = 0
        self._first_enqueued: bool = False
        self._target_to_path: dict[str, Path] = {}
        self._order: list[str] = []
        self._hotkey: CtrlEnterHotkey | None = None

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

    def start(self) -> bool:
        with self._lock:
            if self._active:
                return True
            if not hotkey_ctrl_enter_available():
                self._emit("error", "HTML capture (Ctrl+Enter) requires Windows.")
                return False
            self._debug_port = self._env_debug_port()
            self._target_to_path.clear()
            self._order.clear()
            self._first_enqueued = False
            self._hotkey = CtrlEnterHotkey(self._schedule_snapshot)
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
            self._order.clear()
            self._chrome = None
        if h is not None:
            h.stop()
        if c is not None:
            _terminate_chrome_process(c)

    def enqueue_capture(
        self, url: str, expected_pdf: Path, *, auto_print_pdf: bool = False
    ) -> bool:
        if not self._active:
            self._emit("error", "HTML capture is not started (toggle PDF capture on first).")
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

        if ch is None or ch.poll() is not None:
            if ch is not None:
                _terminate_chrome_process(ch)
            with self._lock:
                self._chrome = None
                self._first_enqueued = False
            proc = launch_isolated_chrome_no_proxy(
                start_url=u,
                remote_debugging_port=dport,
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
            if str(target_id) not in self._order:
                self._order.append(str(target_id))
            if not self._first_enqueued:
                self._first_enqueued = True
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

    def _thread_auto_pod_capture(self, target_id: str, expected_pdf: Path) -> None:
        self._emit(
            "info",
            "Capture: opened carrier tab — waiting for the page to settle, then saving PDF …",
        )
        min_ready = time.monotonic() + 1.25
        deadline = time.monotonic() + 75.0
        last_sig: tuple[int, str] | None = None
        stable = 0
        try:
            while time.monotonic() < deadline:
                with self._lock:
                    if not self._active:
                        return
                ws_url = self._websocket_for_target_id(target_id)
                if not ws_url:
                    time.sleep(0.45)
                    continue
                info = inspect_page(ws_url, text_preview_chars=16000)
                if info and str(info.get("readyState") or "") == "complete":
                    text = str(info.get("text") or "")
                    if len(text) >= 220:
                        sig = (len(text), str(info.get("title") or "")[:120])
                        if sig == last_sig:
                            stable += 1
                        else:
                            stable = 0
                        last_sig = sig
                        if stable >= 2 and time.monotonic() >= min_ready:
                            break
                time.sleep(0.72)
            else:
                self._emit(
                    "error",
                    "Timed out waiting for the tracking page to finish loading.",
                )
                return

            time.sleep(0.4)
            ws_url = self._websocket_for_target_id(target_id)
            if not ws_url:
                self._emit("error", "Lost the capture tab before saving the PDF.")
                return

            if expected_pdf.is_file() and not self._verbose:
                self._emit("info", f"File already exists:\n{expected_pdf.name}")
                return
            expected_pdf.parent.mkdir(parents=True, exist_ok=True)

            pdf_bytes = export_page_pdf(ws_url)
            expected_pdf.write_bytes(pdf_bytes)
            _log_line(f"auto-saved PDF {expected_pdf} ({len(pdf_bytes)} bytes)")

            html_path = expected_pdf.with_name(expected_pdf.stem + "_capture.html")
            html_snip = extract_outer_html_snippet(ws_url, max_chars=350_000)
            if html_snip:
                try:
                    html_path.write_text(html_snip, encoding="utf-8", errors="replace")
                    _log_line(f"auto-saved HTML snapshot {html_path}")
                except OSError as exc:
                    _log_line(f"html snapshot write failed: {exc!r}")

            if self._on_saved is not None:
                try:
                    self._on_saved()
                except Exception:
                    pass
            extra = f"\nHTML snapshot:\n{html_path.name}" if html_snip else ""
            self._emit("info", f"Proof-of-delivery PDF saved:\n{expected_pdf.name}{extra}")
        except Exception as e:
            _log_line("auto pod capture error:\n" + traceback.format_exc())
            self._emit("error", f"Automatic print to PDF failed: {e!s}")

    def _schedule_snapshot(self) -> None:
        threading.Thread(target=self._do_snapshot, name="html-capture-cdp", daemon=True).start()

    def _resolve_path_for_focused_tab(self) -> Path | None:
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
                return tmap[tid]
        for tid in reversed(order):
            if tid in tmap and tid in current_ids:
                return tmap[tid]
        if order and order[-1] in tmap and order[-1] in current_ids:
            return tmap[order[-1]]
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
                    self._order.clear()
                return

            out_path = self._resolve_path_for_focused_tab()
            if out_path is None:
                self._emit(
                    "error",
                    "No matching capture tab. Double-click a tracking row, then try Ctrl+Enter in Chrome.",
                )
                return
            if out_path.is_file() and not self._verbose:
                self._emit("info", f"File already exists:\n{out_path.name}")
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

            pdf_bytes = export_page_pdf(ws_url)
            out_path.write_bytes(pdf_bytes)
            _log_line(f"saved PDF {out_path} ({len(pdf_bytes)} bytes)")
            if self._on_saved is not None:
                try:
                    self._on_saved()
                except Exception:
                    pass
            self._emit("info", f"Proof-of-delivery PDF saved:\n{out_path.name}")
        except Exception as e:
            _log_line("snapshot error:\n" + traceback.format_exc())
            self._emit("error", f"Print to PDF failed: {e!s}")
