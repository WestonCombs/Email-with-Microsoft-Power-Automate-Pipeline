# -*- coding: utf-8 -*-
"""
Passive PDF interception for mitmproxy (UPS/FedEx proof-of-delivery, labels, etc.).

Run mitmdump with this package directory as **current working directory** so paths resolve.

Logs (when started via ``run_pdf_capture.py``): ``mitmdump.stdout.log``, ``mitmdump.stderr.log``.

CLI:

    cd pdfCaptureFromChrome
    python run_pdf_capture.py
    python run_pdf_capture.py "https://example.com/track"

Manual mitmproxy:

    mitmdump -s mitm_pdf_interceptor/mitm_pdf_addon.py --listen-port 8080
    mitmproxy -s mitm_pdf_interceptor/mitm_pdf_addon.py --listen-port 8080

Environment variables (set before mitmproxy, or rely on ``run_pdf_capture.py`` defaults):

    PDF_INTERCEPTOR_OUTPUT_DIR   Directory to write PDFs (default: <BASE_DIR>/email_contents/pdf)
    PDF_INTERCEPTOR_BASENAME     Filename stem, e.g. tracking number (default: captured)
    PDF_INTERCEPTOR_URL_KEYWORDS Comma-separated URL substrings (default includes tracking,fedex,ups,…)
    PDF_INTERCEPTOR_REQUIRE_URL_KEYWORD  If 1/true, only save when URL matches a keyword (default: 1)
    PDF_INTERCEPTOR_MAX_PDFS     Max saves this run; 0 or empty = unlimited (default: unlimited)
    PDF_CAPTURE_DONE_FILE        If set, ``run_pdf_capture.py`` writes a JSON marker here after the first save.

Chrome proxy (Windows):
    Settings → System → Open your computer's proxy settings → Manual proxy
    HTTP and HTTPS proxy: 127.0.0.1  Port: 8080
    Or: chrome.exe --proxy-server="http=127.0.0.1:8080;https=127.0.0.1:8080"

HTTPS: install mitmproxy's CA once:
    mitmproxy shows how; typically run mitmproxy once and open http://mitm.it in Chrome
    to download and install the certificate for "Windows".

This module does not drive a browser; use your own browser pointed at the proxy.
"""

from __future__ import annotations

import base64
import binascii
import json
import logging
import os
import re
import sys
from pathlib import Path
from typing import Iterable, NamedTuple, Optional

# mitmproxy types (script runs inside mitmproxy's Python)
from mitmproxy import http

_PDF_CAPTURE_ROOT = Path(__file__).resolve().parent.parent
if str(_PDF_CAPTURE_ROOT) not in sys.path:
    sys.path.insert(0, str(_PDF_CAPTURE_ROOT))

from paths import default_pdf_output_dir

logger = logging.getLogger("pdf_interceptor")


def _env_bool(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in ("1", "true", "yes", "on")


# ---------------------------------------------------------------------------
# Detection: Content-Type and Content-Disposition
# ---------------------------------------------------------------------------


class PdfDetection(NamedTuple):
    """Result of inspecting response headers (and optional body sniff)."""

    reason: str
    content_type: Optional[str]
    content_disposition: Optional[str]
    inferred_filename: Optional[str]


class PdfDetector:
    """Pure header/body sniffing — no I/O."""

    _pdf_magic = b"%PDF"
    _disp_filename = re.compile(
        r'filename\*?=(?:UTF-8\'\')?"?([^";]+)"?',
        re.IGNORECASE,
    )

    def detect(self, flow: http.HTTPFlow) -> Optional[PdfDetection]:
        if flow.response is None:
            return None

        h = flow.response.headers
        ct_raw = h.get("content-type")
        ct_display = str(ct_raw) if ct_raw is not None else None
        ct_lc = ct_display.lower() if ct_display else None

        cd_raw = h.get("content-disposition")
        cd_display = str(cd_raw) if cd_raw is not None else None

        body = _response_body_bytes(flow.response)

        reasons: list[str] = []

        if ct_lc and "application/pdf" in ct_lc:
            reasons.append("Content-Type: application/pdf")

        fname = self._parse_filename_from_disposition(cd_display)
        if fname and fname.lower().endswith(".pdf"):
            reasons.append("Content-Disposition filename ends with .pdf")

        # APIs often use application/octet-stream for PDF bytes; sniff magic.
        if body and body.startswith(self._pdf_magic):
            reasons.append("Body starts with %PDF (sniff)")

        if not reasons:
            return None

        if body and not body.startswith(self._pdf_magic):
            # Mislabeled HTML/errors — skip saving garbage
            logger.warning(
                "PDF-like headers but body does not start with %%PDF — skipping url=%s",
                flow.request.pretty_url,
            )
            return None

        reason = "; ".join(reasons)
        return PdfDetection(
            reason=reason,
            content_type=ct_display,
            content_disposition=cd_display,
            inferred_filename=fname,
        )

    @staticmethod
    def _parse_filename_from_disposition(cd: Optional[str]) -> Optional[str]:
        if not cd:
            return None
        m = PdfDetector._disp_filename.search(cd)
        if not m:
            # Fallback: any segment ending in .pdf
            m2 = re.search(r"([\w\-./%]+\.pdf)", cd, re.IGNORECASE)
            return m2.group(1) if m2 else None
        raw = m.group(1).strip().strip('"')
        # RFC 5987 * form may still have encoding prefix stripped by regex
        return os.path.basename(raw)


# ---------------------------------------------------------------------------
# Filtering: URL keywords
# ---------------------------------------------------------------------------


class UrlKeywordFilter:
    """Prefer carrier / document URLs; optional strict mode."""

    def __init__(
        self,
        keywords: Iterable[str],
        require_match: bool = True,
    ) -> None:
        self.keywords = tuple(k.strip().lower() for k in keywords if k.strip())
        self.require_match = require_match

    def allows(self, url: str) -> bool:
        if not self.require_match:
            return True
        u = url.lower()
        return any(k in u for k in self.keywords)


# ---------------------------------------------------------------------------
# JSON APIs: base64-wrapped PDF (e.g. FedEx POST …/trackingdocument)
# ---------------------------------------------------------------------------


def _decode_b64_if_pdf(s: str) -> Optional[bytes]:
    """If ``s`` is base64 data for a PDF, return decoded bytes."""
    s = s.strip()
    if "base64," in s:
        s = s.split("base64,", 1)[1].strip()
    t = "".join(s.split())
    if len(t) < 50:
        return None
    pad = (-len(t)) % 4
    padded = t + ("=" * pad)
    for dec in (
        lambda b: base64.b64decode(b, validate=False),
        base64.urlsafe_b64decode,
    ):
        try:
            raw = dec(padded)
            if raw.startswith(b"%PDF"):
                return raw
        except (binascii.Error, ValueError):
            continue
    return None


def _walk_json_for_pdf(obj: object) -> Optional[bytes]:
    if isinstance(obj, dict):
        for v in obj.values():
            found = _walk_json_for_pdf(v)
            if found is not None:
                return found
    elif isinstance(obj, list):
        for item in obj:
            found = _walk_json_for_pdf(item)
            if found is not None:
                return found
    elif isinstance(obj, str):
        return _decode_b64_if_pdf(obj)
    return None


def _try_extract_pdf_from_json(body: bytes) -> Optional[bytes]:
    """Walk JSON text and return the first decoded PDF payload, if any."""
    if not body:
        return None
    stripped = body.lstrip()
    if not stripped[:1] in (b"{", b"["):
        return None
    try:
        text = body.decode("utf-8")
    except UnicodeDecodeError:
        return None
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return None
    return _walk_json_for_pdf(data)


# ---------------------------------------------------------------------------
# Saving: disk writes with collision-safe names
# ---------------------------------------------------------------------------


class PdfFileSaver:
    """Writes bytes to output_dir with basename and numeric suffixes."""

    def __init__(self, output_dir: Path, basename: str) -> None:
        self.output_dir = Path(output_dir)
        self.basename = self._sanitize_stem(basename)
        self._saved_count = 0

    @staticmethod
    def _sanitize_stem(name: str) -> str:
        # Windows-forbidden in filenames
        bad = '<>:"/\\|?*'
        out = "".join("_" if c in bad else c for c in name.strip()) or "capture"
        return out[:200]

    def next_path(self) -> Path:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._saved_count += 1
        if self._saved_count == 1:
            return self.output_dir / f"{self.basename}.pdf"
        return self.output_dir / f"{self.basename}_{self._saved_count}.pdf"

    def save(self, data: bytes) -> Path:
        path = self.next_path()
        path.write_bytes(data)
        return path


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------


class PdfInterceptor:
    """
    Plug-in friendly facade: detection → filter → save, with logging.

    Construct manually or use ``from_env()`` for mitmproxy CLI usage.
    """

    def __init__(
        self,
        output_dir: Path | str,
        basename: str = "captured",
        url_keywords: Iterable[str] | None = None,
        require_url_keyword: bool = True,
        max_pdfs: Optional[int] = None,
    ) -> None:
        self.detector = PdfDetector()
        kw = url_keywords or (
            "proof",
            "delivery",
            "pod",
            "label",
            "tracking",
            "shipment",
            "invoice",
            "fedex",
            "ups",
        )
        self.url_filter = UrlKeywordFilter(kw, require_match=require_url_keyword)
        self.saver = PdfFileSaver(Path(output_dir), basename)
        self.max_pdfs = max_pdfs  # None = unlimited
        self._total_saved = 0

    @classmethod
    def from_env(cls) -> "PdfInterceptor":
        out = os.environ.get("PDF_INTERCEPTOR_OUTPUT_DIR", "").strip() or str(default_pdf_output_dir())
        base = os.environ.get("PDF_INTERCEPTOR_BASENAME", "captured")
        kw_raw = os.environ.get(
            "PDF_INTERCEPTOR_URL_KEYWORDS",
            "proof,delivery,pod,label,tracking,shipment,invoice,fedex,ups",
        )
        keywords = [x.strip() for x in kw_raw.split(",")]
        req = _env_bool("PDF_INTERCEPTOR_REQUIRE_URL_KEYWORD", True)
        max_raw = os.environ.get("PDF_INTERCEPTOR_MAX_PDFS", "").strip()
        max_pdfs: Optional[int]
        if not max_raw:
            max_pdfs = None
        else:
            try:
                v = int(max_raw)
                max_pdfs = None if v <= 0 else v
            except ValueError:
                max_pdfs = None
        return cls(
            output_dir=out,
            basename=base,
            url_keywords=keywords,
            require_url_keyword=req,
            max_pdfs=max_pdfs,
        )

    def handle_response(self, flow: http.HTTPFlow) -> None:
        if flow.response is None:
            return
        if self.max_pdfs is not None and self._total_saved >= self.max_pdfs:
            return

        url = flow.request.pretty_url
        body = _response_body_bytes(flow.response)

        detection = self.detector.detect(flow)
        save_bytes = body

        if detection is None and body and self.url_filter.allows(url):
            extracted = _try_extract_pdf_from_json(body)
            if extracted:
                detection = PdfDetection(
                    reason="Base64 PDF inside JSON (carrier API)",
                    content_type=None,
                    content_disposition=None,
                    inferred_filename=None,
                )
                save_bytes = extracted

        if detection is None:
            return

        if not self.url_filter.allows(url):
            logger.info(
                "PDF detected but URL filter rejected (require keyword): url=%s headers=%s",
                url,
                _headers_for_log(flow.response.headers),
            )
            return

        if not save_bytes:
            logger.warning("PDF headers but empty body: url=%s", url)
            return

        path = self.saver.save(save_bytes)
        self._total_saved += 1

        _write_pdf_capture_done(path)

        logger.info("Saved PDF (%s)", detection.reason)
        logger.info("  URL: %s", url)
        logger.info("  Response headers: %s", _headers_for_log(flow.response.headers))
        logger.info("  Saved to: %s", path.resolve())
        if self.max_pdfs is not None and self._total_saved >= self.max_pdfs:
            logger.info(
                "Reached PDF_INTERCEPTOR_MAX_PDFS=%s — further PDFs ignored until restart.",
                self.max_pdfs,
            )

    # Explicit hooks for custom pipelines
    def detect_only(self, flow: http.HTTPFlow) -> Optional[PdfDetection]:
        if flow.response is None:
            return None
        return self.detector.detect(flow)

    def filter_allows_url(self, url: str) -> bool:
        return self.url_filter.allows(url)


class MitmPdfAddon:
    """mitmproxy addon: delegates to PdfInterceptor."""

    def __init__(self, interceptor: Optional[PdfInterceptor] = None) -> None:
        self.interceptor = interceptor or PdfInterceptor.from_env()

    def load(self, loader) -> None:
        # Ensure INFO lines show when running mitmdump/mitmproxy
        if not logging.root.handlers:
            logging.basicConfig(
                level=logging.INFO,
                format="%(levelname)s [pdf_interceptor] %(message)s",
                stream=sys.stdout,
            )
        p = self.interceptor
        logger.info("MitmPdfAddon loaded")
        logger.info("  output_dir=%s", Path(p.saver.output_dir).resolve())
        logger.info("  basename=%s", p.saver.basename)
        logger.info("  url_keywords=%s require=%s max_pdfs=%s", p.url_filter.keywords, p.url_filter.require_match, p.max_pdfs)
        done = os.environ.get("PDF_CAPTURE_DONE_FILE", "").strip()
        if done:
            logger.info("  pdf_capture_done_file=%s", done)

    def response(self, flow: http.HTTPFlow) -> None:
        self.interceptor.handle_response(flow)


# ---------------------------------------------------------------------------
# mitmproxy entrypoint: must expose ``addons``
# ---------------------------------------------------------------------------

addons = [MitmPdfAddon()]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _write_pdf_capture_done(saved_path: Path) -> None:
    """Notify ``run_pdf_capture.py`` that a PDF was written (one-shot automation)."""
    raw = os.environ.get("PDF_CAPTURE_DONE_FILE", "").strip()
    if not raw:
        return
    try:
        payload = {
            "path": str(saved_path.resolve()),
            "filename": saved_path.name,
        }
        Path(raw).write_text(json.dumps(payload), encoding="utf-8", newline="\n")
    except OSError as e:
        logger.warning("Could not write PDF_CAPTURE_DONE_FILE: %s", e)


def _headers_for_log(headers: http.Headers) -> dict[str, str]:
    # Stable snapshot for logging (mitmproxy Headers is multi-valued)
    out: dict[str, str] = {}
    for k, v in headers.items(True):
        out[str(k)] = str(v)
    return out


def _response_body_bytes(response: http.Response) -> bytes:
    # Binary-safe body for mitmproxy 9+
    if hasattr(response, "content") and isinstance(response.content, (bytes, bytearray)):
        return bytes(response.content)
    if hasattr(response, "raw_content"):
        return bytes(response.raw_content)  # type: ignore[arg-type]
    if hasattr(response, "get_content"):
        return response.get_content()  # type: ignore[no-any-return]
    return b""
