"""Email message model and helpers — mail is fetched via Microsoft Graph (OAuth 2.0)."""

from __future__ import annotations

import html
import base64
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path


@dataclass
class EmailMessage:
    """One fetched email with parsed fields."""

    from_raw: str
    subject: str
    body_html: str
    attachments: list[tuple[str, bytes]] = field(default_factory=list)
    # Outlook-style header (optional; filled by Graph fetcher when available)
    to_line: str = ""
    sent_line: str = ""
    received_line: str = ""
    header_title: str = ""
    sent_datetime_iso: str = ""
    received_datetime_iso: str = ""

    @property
    def sender_email(self) -> str:
        return extract_email(self.from_raw)

    @property
    def sender_name(self) -> str:
        return extract_sender_name(self.from_raw)


def format_graph_datetime_local(iso_ts: str | None) -> str:
    """Graph ISO timestamp (e.g. ending in Z) → local 'Wednesday, April 1, 2026 5:30 PM'."""
    if not iso_ts or not iso_ts.strip():
        return ""
    try:
        s = iso_ts.strip().replace("Z", "+00:00")
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        local = dt.astimezone()
        if sys.platform == "win32":
            time_part = local.strftime("%#I:%M %p")
        else:
            time_part = local.strftime("%-I:%M %p")
        return f"{local.strftime('%A, %B ')}{local.day}, {local.year} {time_part}"
    except ValueError:
        return iso_ts.strip()


def _build_outlook_header_fragment(msg: EmailMessage) -> str:
    """Return an isolated, print-safe Outlook-style metadata block.

    Email bodies commonly contain broad rules for ``table``, ``tr``, and
    ``td`` elements.  Using uniquely-classed div/span elements with defensive
    inline styles keeps those rules from hiding the metadata during PDF print.
    Every expected row is rendered so missing Graph data is explicit instead
    of making the header appear partially broken.
    """

    unavailable = "(not available)"

    def display_value(value: str | None) -> str:
        cleaned = (value or "").strip()
        return cleaned or unavailable

    rows = [
        ("From", display_value(msg.from_raw)),
        ("Sent", display_value(msg.sent_line)),
        ("Received", display_value(msg.received_line)),
        ("To", display_value(msg.to_line)),
        ("Subject", display_value(msg.subject)),
    ]

    row_style = (
        "display:table-row !important;visibility:visible !important;"
        "opacity:1 !important;position:static !important;height:auto !important;"
        "overflow:visible !important;color:#000 !important;"
        "font-family:Arial,Helvetica,sans-serif !important;"
        "font-size:11pt !important;line-height:1.35 !important;"
    )
    label_style = (
        "display:table-cell !important;visibility:visible !important;"
        "opacity:1 !important;position:static !important;vertical-align:top !important;"
        "width:78px !important;min-width:78px !important;height:auto !important;"
        "padding:2px 12px 2px 0 !important;overflow:visible !important;"
        "white-space:nowrap !important;color:#000 !important;background:#fff !important;"
        "font-family:Arial,Helvetica,sans-serif !important;font-size:11pt !important;"
        "font-weight:700 !important;line-height:1.35 !important;text-align:left !important;"
    )
    value_style = (
        "display:table-cell !important;visibility:visible !important;"
        "opacity:1 !important;position:static !important;vertical-align:top !important;"
        "width:auto !important;height:auto !important;padding:2px 0 !important;"
        "overflow:visible !important;white-space:normal !important;"
        "color:#000 !important;background:#fff !important;"
        "font-family:Arial,Helvetica,sans-serif !important;font-size:11pt !important;"
        "font-weight:400 !important;line-height:1.35 !important;text-align:left !important;"
    )

    inner_rows = []
    for label, value in rows:
        inner_rows.append(
            f'<div class="email-sorter-meta-row" style="{row_style}">'
            f'<span class="email-sorter-meta-label" style="{label_style}">'
            f"{html.escape(label)}:</span>"
            f'<span class="email-sorter-meta-value" style="{value_style}">'
            f"{html.escape(value)}</span>"
            "</div>"
        )

    metadata_rows = (
        '<div class="email-sorter-meta-rows" style="display:table !important;'
        'visibility:visible !important;opacity:1 !important;position:static !important;'
        'table-layout:auto !important;border-collapse:collapse !important;width:100% !important;'
        'max-width:900px !important;height:auto !important;overflow:visible !important;'
        'color:#000 !important;background:#fff !important;'
        'font-family:Arial,Helvetica,sans-serif !important;font-size:11pt !important;'
        'line-height:1.35 !important;text-align:left !important;">'
        + "".join(inner_rows)
        + "</div>"
    )

    title = (msg.header_title or "").strip()
    parts: list[str] = [
        '<div class="email-sorter-meta-header" style="display:block !important;'
        'visibility:visible !important;opacity:1 !important;position:static !important;'
        'float:none !important;clear:both !important;width:auto !important;max-width:none !important;'
        'height:auto !important;min-height:0 !important;margin:0 0 16px 0 !important;'
        'padding:0 !important;overflow:visible !important;color:#000 !important;'
        'background:#fff !important;font-family:Arial,Helvetica,sans-serif !important;'
        'font-size:11pt !important;font-weight:400 !important;line-height:1.35 !important;'
        'text-align:left !important;writing-mode:horizontal-tb !important;">',
    ]
    if title:
        parts.append(
            '<div class="email-sorter-meta-title" style="display:block !important;'
            'visibility:visible !important;opacity:1 !important;position:static !important;'
            'width:auto !important;height:auto !important;margin:0 !important;padding:0 0 8px 0 !important;'
            'overflow:visible !important;border:0 !important;border-bottom:2px solid #000 !important;'
            'color:#000 !important;background:#fff !important;'
            'font-family:Arial,Helvetica,sans-serif !important;font-size:14pt !important;'
            'font-weight:700 !important;line-height:1.25 !important;text-align:left !important;">'
            f"{html.escape(title)}</div>"
            '<div aria-hidden="true" style="display:block !important;height:8px !important;'
            'visibility:visible !important;background:#fff !important;"></div>'
        )
    parts.append(metadata_rows)
    parts.append("</div>")
    return "".join(parts)


def prepend_outlook_style_header(body_html: str, msg: EmailMessage) -> str:
    """Put an Outlook-like metadata block at the top of *body_html* (used for PDF print)."""
    header = _build_outlook_header_fragment(msg)
    m = re.search(r"(<body[^>]*>)", body_html, re.IGNORECASE | re.DOTALL)
    if m:
        pos = m.end()
        return body_html[:pos] + "\n" + header + "\n" + body_html[pos:]
    m_head = re.search(r"(</head\s*>)", body_html, re.IGNORECASE | re.DOTALL)
    if m_head:
        insert_at = m_head.end()
        return (
            body_html[:insert_at]
            + "\n<body>\n"
            + header
            + "\n"
            + body_html[insert_at:]
        )
    return (
        "<!DOCTYPE html>\n<html><head><meta charset=\"utf-8\" /></head><body>\n"
        + header
        + "\n"
        + body_html
        + "\n</body></html>\n"
    )


_CID_SRC_RE = re.compile(
    r"(?P<prefix>\bsrc\s*=\s*[\"'])\s*cid:(?P<cid>[^\"'>\s]+)(?P<suffix>[\"'])",
    re.IGNORECASE,
)


def _normalize_cid_token(token: str | None) -> str:
    raw = (token or "").strip().strip("<>").strip()
    if raw.lower().startswith("cid:"):
        raw = raw[4:]
    return raw.casefold()


def inline_cid_images(
    body_html: str,
    inline_images: dict[str, tuple[str, bytes]],
) -> tuple[str, int, int]:
    """Replace ``src="cid:..."`` references with ``data:`` URIs.

    Returns ``(updated_html, replaced_count, unresolved_count)``.
    """
    if not body_html or not inline_images:
        return body_html, 0, 0

    resolved_data_uris: dict[str, str] = {}
    for raw_cid, payload in inline_images.items():
        norm = _normalize_cid_token(raw_cid)
        if not norm:
            continue
        content_type, data = payload
        if not data:
            continue
        ctype = (content_type or "application/octet-stream").strip()
        encoded = base64.b64encode(data).decode("ascii")
        resolved_data_uris[norm] = f"data:{ctype};base64,{encoded}"

    if not resolved_data_uris:
        return body_html, 0, 0

    replaced = 0
    unresolved = 0

    def _replace(match: re.Match[str]) -> str:
        nonlocal replaced, unresolved
        token = _normalize_cid_token(match.group("cid"))
        data_uri = resolved_data_uris.get(token)
        if not data_uri:
            unresolved += 1
            return match.group(0)
        replaced += 1
        return f"{match.group('prefix')}{data_uri}{match.group('suffix')}"

    updated = _CID_SRC_RE.sub(_replace, body_html)
    return updated, replaced, unresolved


def extract_email(text: str) -> str:
    """'Display Name <user@example.com>' -> 'user@example.com'"""
    match = re.search(r"<([^<>]+)>", text)
    addr = match.group(1) if match else text
    return addr.replace("\ufeff", "").strip()


def extract_sender_name(text: str) -> str:
    """'"John Doe" <user@example.com>' -> 'John Doe'"""
    match = re.search(r'"([^"]+)"', text)
    name = match.group(1) if match else text
    return name.replace("\ufeff", "").strip()


def save_attachments(attachments: list[tuple[str, bytes]], save_dir: Path) -> list[Path]:
    """Write attachment tuples to *save_dir*, handling name collisions."""
    saved: list[Path] = []
    save_dir.mkdir(parents=True, exist_ok=True)
    for filename, data in attachments:
        target = save_dir / filename
        if target.exists():
            stem, suffix = target.stem, target.suffix
            n = 1
            while target.exists():
                target = save_dir / f"{stem}_{n}{suffix}"
                n += 1
        target.write_bytes(data)
        saved.append(target)
    return saved


def fetch_emails(
    *,
    mail_folder: str,
    attachments_dir: Path | None = None,
    azure_client_id: str,
    azure_tenant_id: str = "common",
    auth_flow: str = "interactive",
    token_cache_path: Path | None = None,
    force_full_graph_auth: bool = False,
    cancel_check=None,
    base_dir: Path | None = None,
):
    """Fetch every message in *mail_folder* using Microsoft Graph (see ``ms_graph_fetcher``)."""
    from . import ms_graph_fetcher

    return ms_graph_fetcher.fetch_emails(
        mail_folder=mail_folder,
        attachments_dir=attachments_dir,
        azure_client_id=azure_client_id,
        azure_tenant_id=azure_tenant_id,
        auth_flow=auth_flow,
        token_cache_path=token_cache_path,
        force_full_graph_auth=force_full_graph_auth,
        cancel_check=cancel_check,
        base_dir=base_dir,
    )
