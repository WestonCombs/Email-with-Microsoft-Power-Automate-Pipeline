"""Email message model and helpers — mail is fetched via Microsoft Graph (OAuth 2.0)."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path


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
    )
