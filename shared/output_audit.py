"""Audit HTML/PDF output integrity and attempt safe auto-fixes."""

from __future__ import annotations

import re
import subprocess
import os
from pathlib import Path


def _find_browser() -> Path | None:
    candidates: list[Path] = []
    for env_var in ("PROGRAMFILES(X86)", "PROGRAMFILES"):
        raw = (os.environ.get(env_var) or "").strip()
        if not raw:
            continue
        base = Path(raw)
        candidates.append(base / "Microsoft" / "Edge" / "Application" / "msedge.exe")
        candidates.append(base / "Google" / "Chrome" / "Application" / "chrome.exe")
    for p in candidates:
        if p.is_file():
            return p
    return None


def _looks_like_html_bytes(data: bytes) -> bool:
    snippet = data[:512].decode("utf-8", errors="ignore").lower()
    return "<html" in snippet or "<!doctype html" in snippet or "<body" in snippet


def _rebuild_pdf_from_html(html_path: Path, pdf_path: Path) -> bool:
    browser = _find_browser()
    if browser is None:
        return False
    try:
        subprocess.run(
            [
                str(browser),
                "--headless",
                "--disable-gpu",
                "--no-pdf-header-footer",
                f"--print-to-pdf={pdf_path}",
                html_path.resolve().as_uri(),
            ],
            capture_output=True,
            timeout=30,
        )
    except Exception:
        return False
    return pdf_path.is_file() and pdf_path.stat().st_size > 0


def _normalize_key(name: str) -> str:
    stem = Path(name).stem
    stem = re.sub(r"\s+\(\d+\)$", "", stem).strip()
    return stem.casefold()


def audit_email_outputs(directory_path: str | Path) -> dict:
    """Validate HTML↔PDF pairs and auto-fix missing/malformed PDFs when possible."""
    root = Path(directory_path).expanduser().resolve()
    html_dir = root / "html"
    pdf_dir = root / "pdf"
    report = {
        "html_only": [],
        "pdf_only": [],
        "malformed_pdf": [],
        "fixed_pdf": [],
        "needs_review": [],
    }

    if not html_dir.is_dir() or not pdf_dir.is_dir():
        report["needs_review"].append(
            f"Expected sibling folders missing under {root} (need html/ and pdf/)."
        )
        return report

    html_files = [p for p in html_dir.glob("*.html") if p.is_file()]
    pdf_files = [p for p in pdf_dir.glob("*.pdf") if p.is_file()]

    html_by_key = {_normalize_key(p.name): p for p in html_files}
    pdf_by_key = {_normalize_key(p.name): p for p in pdf_files}

    html_keys = set(html_by_key)
    pdf_keys = set(pdf_by_key)

    for key in sorted(html_keys - pdf_keys):
        html_path = html_by_key[key]
        pdf_path = pdf_dir / f"{html_path.stem}.pdf"
        report["html_only"].append(str(html_path))
        if _rebuild_pdf_from_html(html_path, pdf_path):
            report["fixed_pdf"].append(str(pdf_path))
        else:
            report["needs_review"].append(f"Missing PDF for HTML: {html_path}")

    for key in sorted(pdf_keys - html_keys):
        pdf_path = pdf_by_key[key]
        report["pdf_only"].append(str(pdf_path))
        report["needs_review"].append(f"Missing HTML for PDF: {pdf_path}")

    for key in sorted(pdf_keys):
        pdf_path = pdf_by_key[key]
        try:
            raw = pdf_path.read_bytes()
        except OSError:
            report["needs_review"].append(f"Unreadable PDF file: {pdf_path}")
            continue
        if raw.startswith(b"%PDF-"):
            continue
        if _looks_like_html_bytes(raw):
            report["malformed_pdf"].append(str(pdf_path))
            html_peer = html_dir / (pdf_path.stem + ".html")
            try:
                if not html_peer.exists():
                    html_peer.write_bytes(raw)
                if _rebuild_pdf_from_html(html_peer, pdf_path):
                    report["fixed_pdf"].append(str(pdf_path))
                else:
                    report["needs_review"].append(f"Could not regenerate malformed PDF: {pdf_path}")
            except OSError:
                report["needs_review"].append(f"Failed to recover malformed PDF: {pdf_path}")

    return report
