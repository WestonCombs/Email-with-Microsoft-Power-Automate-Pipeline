"""Automated saved-email HTML to PDF conversion through Chrome DevTools."""

from __future__ import annotations

import json
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path

from .chrome_devtools import (
    _CdpSession,
    export_page_pdf,
    list_page_targets,
    reserve_free_port,
    wait_for_debugger,
)
from .launch_mitm_chrome import launch_isolated_chrome_no_proxy


@dataclass(frozen=True)
class EmailHtmlPdfResult:
    html_path: Path
    pdf_path: Path
    ready_state: str
    image_count: int
    images_complete: int
    images_loaded: int
    images_failed: int
    elapsed_seconds: float

    @property
    def image_summary(self) -> str:
        return (
            f"images_loaded={self.images_loaded}/{self.image_count} "
            f"complete={self.images_complete}/{self.image_count} "
            f"failed={self.images_failed}"
        )


def convert_email_html_to_pdf(
    html_path: Path,
    pdf_path: Path,
    *,
    wait_timeout: float = 35.0,
    startup_timeout: float = 15.0,
) -> EmailHtmlPdfResult:
    """Load ``html_path`` in the capture Chrome profile and save ``pdf_path``.

    This intentionally uses non-headless Chrome plus CDP ``Page.printToPDF``.
    Some retail email image hosts block headless/file-print paths but load when
    the existing capture profile and browser context are used.
    """
    source = Path(html_path).expanduser().resolve()
    output = Path(pdf_path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"HTML file not found: {source}")
    output.parent.mkdir(parents=True, exist_ok=True)

    port = reserve_free_port()
    proc = launch_isolated_chrome_no_proxy(
        start_url="about:blank",
        remote_debugging_port=port,
        extra_args=(
            "--start-minimized",
            "--window-position=-32000,-32000",
            "--window-size=1280,1600",
            "--allow-file-access-from-files",
            "--blink-settings=imagesEnabled=true",
            "--disable-features=PaintHolding",
            "--disable-session-crashed-bubble",
            "--disable-restore-session-state",
        ),
        verbose=False,
    )
    if proc is None:
        raise RuntimeError("Could not launch capture Chrome")

    started = time.time()
    tmp_pdf = output.with_name(f".{output.name}.cdp_tmp")
    try:
        if not wait_for_debugger(port, timeout=startup_timeout):
            raise RuntimeError("Chrome DevTools did not become available")

        target = _first_page_target(port)
        if target is None:
            raise RuntimeError("Could not find a Chrome page target")
        ws_url = str(target.get("webSocketDebuggerUrl") or "")
        if not ws_url:
            raise RuntimeError("Chrome page target has no DevTools websocket")

        with _CdpSession(ws_url) as cdp:
            cdp.call("Page.enable")
            cdp.call("Runtime.enable")
            cdp.call("Page.navigate", {"url": source.as_uri()})

        stats = _wait_for_document_and_images(ws_url, timeout=wait_timeout)
        _scroll_to_top(ws_url)

        pdf_bytes = export_page_pdf(ws_url)
        if not pdf_bytes:
            raise RuntimeError("Chrome DevTools returned an empty PDF")
        tmp_pdf.write_bytes(pdf_bytes)
        tmp_pdf.replace(output)

        elapsed = time.time() - started
        return EmailHtmlPdfResult(
            html_path=source,
            pdf_path=output,
            ready_state=str(stats.get("readyState") or ""),
            image_count=int(stats.get("imageCount") or 0),
            images_complete=int(stats.get("imagesComplete") or 0),
            images_loaded=int(stats.get("imagesLoaded") or 0),
            images_failed=int(stats.get("imagesFailed") or 0),
            elapsed_seconds=round(elapsed, 3),
        )
    finally:
        try:
            if tmp_pdf.exists():
                tmp_pdf.unlink()
        except OSError:
            pass
        _terminate_process_tree(proc)


def _first_page_target(port: int) -> dict | None:
    deadline = time.time() + 10.0
    while time.time() < deadline:
        for target in list_page_targets(port):
            url = str(target.get("url") or "")
            if url.startswith("chrome-devtools://"):
                continue
            return target
        time.sleep(0.25)
    return None


def _wait_for_document_and_images(ws_url: str, *, timeout: float) -> dict:
    deadline = time.time() + max(timeout, 1.0)
    last_signature: tuple[int, ...] | None = None
    stable_polls = 0
    stats: dict = {}

    while True:
        stats = _page_image_stats(ws_url)
        signature = (
            int(stats.get("imageCount") or 0),
            int(stats.get("imagesComplete") or 0),
            int(stats.get("imagesLoaded") or 0),
            int(stats.get("imagesFailed") or 0),
            int(stats.get("scrollHeight") or 0),
        )
        ready = str(stats.get("readyState") or "") == "complete"
        all_images_finished = signature[1] >= signature[0]
        if ready and signature == last_signature:
            stable_polls += 1
        else:
            stable_polls = 0
            last_signature = signature

        if ready and all_images_finished and stable_polls >= 2:
            return stats
        if time.time() >= deadline:
            return stats

        _nudge_lazy_images(ws_url)
        time.sleep(0.65)


def _page_image_stats(ws_url: str) -> dict:
    expr = r"""
(() => {
  const images = Array.from(document.images || []);
  const rows = images.map((img, index) => ({
    index,
    src: img.currentSrc || img.src || "",
    complete: Boolean(img.complete),
    naturalWidth: Number(img.naturalWidth || 0),
    naturalHeight: Number(img.naturalHeight || 0),
    renderedWidth: Math.round((img.getBoundingClientRect && img.getBoundingClientRect().width) || 0),
    renderedHeight: Math.round((img.getBoundingClientRect && img.getBoundingClientRect().height) || 0)
  }));
  return JSON.stringify({
    readyState: document.readyState || "",
    imageCount: rows.length,
    imagesComplete: rows.filter((img) => img.complete).length,
    imagesLoaded: rows.filter((img) => img.complete && img.naturalWidth > 0).length,
    imagesFailed: rows.filter((img) => img.complete && img.naturalWidth <= 0).length,
    scrollHeight: Math.max(
      document.body ? document.body.scrollHeight : 0,
      document.documentElement ? document.documentElement.scrollHeight : 0
    ),
    images: rows
  });
})()
"""
    with _CdpSession(ws_url) as cdp:
        result = cdp.call("Runtime.evaluate", {"expression": expr, "returnByValue": True})
    payload = result.get("result") if isinstance(result, dict) else None
    value = payload.get("value") if isinstance(payload, dict) else None
    if not isinstance(value, str):
        return {}
    try:
        decoded = json.loads(value)
    except json.JSONDecodeError:
        return {}
    return decoded if isinstance(decoded, dict) else {}


def _nudge_lazy_images(ws_url: str) -> None:
    expr = r"""
(() => {
  document.querySelectorAll('img').forEach((img) => {
    img.loading = 'eager';
    if (!img.getAttribute('src')) {
      ['data-src', 'data-original', 'data-lazy-src', 'data-url'].some((name) => {
        const value = img.getAttribute(name);
        if (value) {
          img.setAttribute('src', value);
          return true;
        }
        return false;
      });
    }
  });
  const maxY = Math.max(
    document.body ? document.body.scrollHeight : 0,
    document.documentElement ? document.documentElement.scrollHeight : 0
  );
  window.scrollTo(0, maxY);
  return true;
})()
"""
    try:
        with _CdpSession(ws_url) as cdp:
            cdp.call("Runtime.evaluate", {"expression": expr, "returnByValue": True})
    except Exception:
        pass


def _scroll_to_top(ws_url: str) -> None:
    try:
        with _CdpSession(ws_url) as cdp:
            cdp.call(
                "Runtime.evaluate",
                {"expression": "window.scrollTo(0, 0); true", "returnByValue": True},
            )
    except Exception:
        pass
    time.sleep(0.35)


def _terminate_process_tree(proc: subprocess.Popen) -> None:
    if proc.poll() is not None:
        return
    if sys.platform == "win32":
        try:
            subprocess.run(
                ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                capture_output=True,
                text=True,
                stdin=subprocess.DEVNULL,
                timeout=10,
            )
            return
        except (OSError, subprocess.SubprocessError):
            pass
    proc.terminate()
    try:
        proc.wait(timeout=5)
    except subprocess.TimeoutExpired:
        proc.kill()
