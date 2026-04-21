"""Minimal Chrome DevTools helpers for print-preview fallback capture."""

from __future__ import annotations

import base64
import json
import socket
import time
import urllib.error
import urllib.request
from pathlib import Path

import websocket

PRINT_KEYWORDS = (
    "proof of delivery",
    "tracking number",
    "delivered on",
    "delivery location",
    "print this page",
    "print",
    "label",
    "shipment",
    "receipt",
    "pod",
)


def reserve_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _json_get(url: str, *, timeout: float = 2.0) -> object:
    with urllib.request.urlopen(url, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def wait_for_debugger(port: int, *, timeout: float = 10.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            _json_get(f"http://127.0.0.1:{port}/json/version", timeout=1.5)
            return True
        except (OSError, urllib.error.URLError, json.JSONDecodeError):
            time.sleep(0.25)
    return False


def list_page_targets(port: int) -> list[dict]:
    data = _json_get(f"http://127.0.0.1:{port}/json/list", timeout=2.0)
    if not isinstance(data, list):
        return []
    out: list[dict] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        if item.get("type") != "page":
            continue
        if not item.get("webSocketDebuggerUrl"):
            continue
        out.append(item)
    return out


class _CdpSession:
    def __init__(self, ws_url: str, *, timeout: float = 4.0) -> None:
        self._ws = websocket.create_connection(ws_url, timeout=timeout)
        self._next_id = 0

    def close(self) -> None:
        self._ws.close()

    def call(self, method: str, params: dict | None = None) -> dict:
        self._next_id += 1
        msg_id = self._next_id
        payload = {"id": msg_id, "method": method}
        if params:
            payload["params"] = params
        self._ws.send(json.dumps(payload))
        while True:
            raw = self._ws.recv()
            data = json.loads(raw)
            if not isinstance(data, dict):
                continue
            if data.get("id") != msg_id:
                continue
            if "error" in data:
                raise RuntimeError(str(data["error"]))
            result = data.get("result")
            return result if isinstance(result, dict) else {}

    def __enter__(self) -> _CdpSession:
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()


def inspect_page(ws_url: str) -> dict | None:
    expr = """
(() => {
  const body = document.body ? document.body.innerText : "";
  return JSON.stringify({
    href: location.href,
    title: document.title || "",
    readyState: document.readyState || "",
    text: body.slice(0, 4000)
  });
})()
"""
    try:
        with _CdpSession(ws_url) as cdp:
            result = cdp.call("Runtime.evaluate", {"expression": expr, "returnByValue": True})
    except Exception:
        return None
    payload = result.get("result")
    if not isinstance(payload, dict):
        return None
    value = payload.get("value")
    if not isinstance(value, str):
        return None
    try:
        decoded = json.loads(value)
    except json.JSONDecodeError:
        return None
    return decoded if isinstance(decoded, dict) else None


def looks_like_print_preview(target: dict, page_info: dict | None) -> bool:
    url = str((page_info or {}).get("href") or target.get("url") or "").lower()
    title = str((page_info or {}).get("title") or target.get("title") or "").lower()
    text = str((page_info or {}).get("text") or "").lower()

    has_doc_keyword = any(k in f"{title}\n{text}" for k in PRINT_KEYWORDS)
    looks_like_popup = (
        url == "about:blank"
        or url.startswith("chrome://print")
        or "print" in url
        or "proof" in url
    )
    preview_hits = sum(k in text for k in ("destination", "pages", "color", "cancel"))

    return has_doc_keyword and (looks_like_popup or preview_hits >= 2)


def export_page_pdf(ws_url: str) -> bytes:
    with _CdpSession(ws_url) as cdp:
        cdp.call("Page.enable")
        result = cdp.call(
            "Page.printToPDF",
            {
                "printBackground": True,
                "preferCSSPageSize": True,
            },
        )
    data = result.get("data")
    if not isinstance(data, str) or not data:
        raise RuntimeError("Page.printToPDF returned no data")
    return base64.b64decode(data)


def next_pdf_path(output_dir: Path, basename: str) -> Path:
    root = Path(output_dir)
    root.mkdir(parents=True, exist_ok=True)
    stem = basename.strip() or "captured"
    first = root / f"{stem}.pdf"
    if not first.exists():
        return first
    idx = 2
    while True:
        candidate = root / f"{stem}_{idx}.pdf"
        if not candidate.exists():
            return candidate
        idx += 1
