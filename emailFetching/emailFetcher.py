"""IMAP email fetching — replaces the Power Automate email retrieval step.

Uses only Python standard-library modules (imaplib, email) so no extra
dependencies are needed.
"""

from __future__ import annotations

import email as email_lib
import imaplib
import re
from dataclasses import dataclass, field
from email.header import decode_header
from pathlib import Path


# ──────────────────────────────────────────────
# Data class returned per message
# ──────────────────────────────────────────────
@dataclass
class EmailMessage:
    """One fetched email with parsed fields."""

    from_raw: str
    subject: str
    body_html: str
    attachments: list[tuple[str, bytes]] = field(default_factory=list)

    @property
    def sender_email(self) -> str:
        return extract_email(self.from_raw)

    @property
    def sender_name(self) -> str:
        return extract_sender_name(self.from_raw)


# ──────────────────────────────────────────────
# From-header helpers (port of the inline PA scripts)
# ──────────────────────────────────────────────
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


# ──────────────────────────────────────────────
# Internal helpers
# ──────────────────────────────────────────────
def _decode_header_value(value: str) -> str:
    """Decode RFC 2047 encoded header values."""
    parts = decode_header(value)
    decoded: list[str] = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(part)
    return "".join(decoded)


def _get_html_body(msg: email_lib.message.Message) -> str:
    """Return the HTML body, falling back to plain text if no HTML part exists."""
    if msg.is_multipart():
        html_parts: list[str] = []
        text_parts: list[str] = []
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition", ""))
            if "attachment" in disposition:
                continue
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            charset = part.get_content_charset() or "utf-8"
            decoded = payload.decode(charset, errors="replace")
            if content_type == "text/html":
                html_parts.append(decoded)
            elif content_type == "text/plain":
                text_parts.append(decoded)
        return "\n".join(html_parts) if html_parts else "\n".join(text_parts)

    payload = msg.get_payload(decode=True)
    if payload is None:
        return ""
    charset = msg.get_content_charset() or "utf-8"
    return payload.decode(charset, errors="replace")


def _get_attachments(msg: email_lib.message.Message) -> list[tuple[str, bytes]]:
    """Return list of (filename, raw_bytes) for every attachment."""
    attachments: list[tuple[str, bytes]] = []
    if not msg.is_multipart():
        return attachments
    for part in msg.walk():
        disposition = str(part.get("Content-Disposition", ""))
        if "attachment" not in disposition:
            continue
        filename = part.get_filename()
        filename = _decode_header_value(filename) if filename else "unnamed_attachment"
        data = part.get_payload(decode=True)
        if data:
            attachments.append((filename, data))
    return attachments


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


# ──────────────────────────────────────────────
# Public entry point
# ──────────────────────────────────────────────
def fetch_emails(
    *,
    imap_server: str,
    port: int,
    username: str,
    password: str,
    mail_folder: str,
    use_ssl: bool = True,
    attachments_dir: Path | None = None,
) -> list[EmailMessage]:
    """Connect to an IMAP server and fetch every message in *mail_folder*.

    Attachments are saved to *attachments_dir* when provided.
    """
    conn: imaplib.IMAP4_SSL | imaplib.IMAP4
    if use_ssl:
        conn = imaplib.IMAP4_SSL(imap_server, port)
    else:
        conn = imaplib.IMAP4(imap_server, port)

    try:
        conn.login(username, password)
        status, _ = conn.select(mail_folder, readonly=True)
        if status != "OK":
            raise RuntimeError(f"Could not select folder '{mail_folder}': {status}")

        _, data = conn.search(None, "ALL")
        msg_nums = data[0].split()
        print(f"  Found {len(msg_nums)} email(s) in '{mail_folder}'")

        messages: list[EmailMessage] = []
        for num in msg_nums:
            _, msg_data = conn.fetch(num, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email_lib.message_from_bytes(raw_email)

            from_raw = _decode_header_value(msg.get("From", ""))
            subject = _decode_header_value(msg.get("Subject", ""))
            body_html = _get_html_body(msg)
            attachments = _get_attachments(msg)

            if attachments_dir and attachments:
                saved = save_attachments(attachments, attachments_dir)
                for p in saved:
                    print(f"    Saved attachment: {p.name}")

            messages.append(EmailMessage(
                from_raw=from_raw,
                subject=subject,
                body_html=body_html,
                attachments=attachments,
            ))

        return messages
    finally:
        try:
            conn.close()
        except Exception:
            pass
        conn.logout()
