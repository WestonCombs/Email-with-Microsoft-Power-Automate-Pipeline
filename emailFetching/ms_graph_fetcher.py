"""Fetch mail via Microsoft Graph API (OAuth 2.0 delegated) — replaces IMAP."""

from __future__ import annotations

import base64
import concurrent.futures
import json
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import msal

from .emailFetcher import EmailMessage, format_graph_datetime_local, save_attachments

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
# Silent MSAL refresh can block indefinitely; cap it so stale/expired sessions fail fast.
SILENT_AUTH_TIMEOUT_S = 30.0
# JSON list/detail calls — keep attachment downloads separate (can be large).
GRAPH_JSON_TIMEOUT_S = 30.0
# MSAL uses requests for authority/instance discovery; default timeout=None can hang forever.
MSAL_HTTP_TIMEOUT_S = 30.0


def _graph_get(url: str, access_token: str) -> dict:
    req = urllib.request.Request(
        url,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
        },
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=GRAPH_JSON_TIMEOUT_S) as resp:
            raw = resp.read()
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Graph HTTP {e.code} for {url}: {detail}") from e
    except urllib.error.URLError as e:
        if isinstance(e.reason, TimeoutError):
            raise RuntimeError(
                f"Graph request timed out after {GRAPH_JSON_TIMEOUT_S:.0f}s: {url}"
            ) from e
        raise
    if not raw:
        return {}
    return json.loads(raw.decode("utf-8"))


def _graph_get_all_pages(first_url: str, access_token: str) -> list[dict]:
    rows: list[dict] = []
    url: str | None = first_url
    while url:
        data = _graph_get(url, access_token)
        if isinstance(data, list):
            rows.extend(data)
            break
        rows.extend(data.get("value") or [])
        url = data.get("@odata.nextLink")
    return rows


def _acquire_token(
    *,
    client_id: str,
    tenant_id: str,
    cache_path: Path,
    auth_flow: str,
    force_interactive: bool = False,
) -> str:
    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        try:
            cache.deserialize(cache_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            pass

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=cache,
        timeout=MSAL_HTTP_TIMEOUT_S,
    )
    scopes = ["Mail.Read"]

    if not force_interactive:
        accounts = app.get_accounts()
        if accounts:
            result: dict | None
            executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
            try:
                fut = executor.submit(
                    app.acquire_token_silent, scopes, account=accounts[0]
                )
                result = fut.result(timeout=SILENT_AUTH_TIMEOUT_S)
            except concurrent.futures.TimeoutError:
                print(
                    f"WARNING: Silent Microsoft login/token refresh took longer than "
                    f"{SILENT_AUTH_TIMEOUT_S:.0f}s (often expired or stuck credentials). "
                    f"Removing {cache_path} so the next sign-in can rebuild the cache.\n"
                    "Opening interactive or device-code sign-in …",
                    file=sys.stderr,
                )
                try:
                    cache_path.unlink(missing_ok=True)
                except OSError:
                    pass
                cache = msal.SerializableTokenCache()
                app = msal.PublicClientApplication(
                    client_id,
                    authority=authority,
                    token_cache=cache,
                    timeout=MSAL_HTTP_TIMEOUT_S,
                )
                result = None
            finally:
                # Do not wait for a stuck acquire_token_silent worker (with-block would).
                executor.shutdown(wait=False)

            if result and result.get("access_token"):
                if cache.has_state_changed:
                    cache_path.parent.mkdir(parents=True, exist_ok=True)
                    cache_path.write_text(cache.serialize(), encoding="utf-8")
                return result["access_token"]

    auth_flow_norm = (auth_flow or "interactive").strip().lower()
    if auth_flow_norm == "device_code":
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(
                "Device code flow failed to start: "
                + json.dumps(flow, indent=2)
            )
        print(flow["message"], file=sys.stderr)
        result = app.acquire_token_by_device_flow(flow)
    else:
        # prompt=login: skip SSO for this request so the user can pick another
        # account and complete MFA (used when DEMO_MODE forces full sign-in).
        interactive_kw: dict = {"scopes": scopes}
        if force_interactive:
            interactive_kw["prompt"] = "login"
        result = app.acquire_token_interactive(**interactive_kw)

    if cache.has_state_changed:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(cache.serialize(), encoding="utf-8")

    if not result or not result.get("access_token"):
        err = result.get("error_description") or result.get("error") if result else "unknown"
        raise RuntimeError(f"Could not acquire Graph token: {err}")

    return result["access_token"]


def _normalize_well_known_folder(name: str) -> str | None:
    n = name.strip().lower()
    aliases = {
        "inbox": "inbox",
        "sent": "sentitems",
        "sent items": "sentitems",
        "sentitems": "sentitems",
        "drafts": "drafts",
        "deleted": "deleteditems",
        "deleted items": "deleteditems",
        "junk": "junkemail",
        "junk email": "junkemail",
        "archive": "archive",
        "outbox": "outbox",
    }
    return aliases.get(n)


def _folder_matches_display(folder: dict, wanted: str) -> bool:
    dn = (folder.get("displayName") or "").strip().casefold()
    return dn == wanted.casefold()


def _resolve_folder_id(access_token: str, mail_folder: str) -> str:
    name = mail_folder.strip()
    if not name:
        raise ValueError("mail_folder must be non-empty")

    well = _normalize_well_known_folder(name)
    if well:
        meta = _graph_get(f"{GRAPH_ROOT}/me/mailFolders/{well}", access_token)
        fid = meta.get("id")
        if not fid:
            raise RuntimeError(f"Could not resolve well-known folder '{well}': {meta}")
        return fid

    # Search top-level and nested child folders by display name
    to_scan: list[tuple[str | None, str]] = [(None, "")]
    seen: set[str] = set()

    while to_scan:
        parent_id, _ = to_scan.pop(0)
        if parent_id is None:
            url = f"{GRAPH_ROOT}/me/mailFolders?$top=200"
        else:
            if parent_id in seen:
                continue
            seen.add(parent_id)
            url = f"{GRAPH_ROOT}/me/mailFolders/{parent_id}/childFolders?$top=200"

        for folder in _graph_get_all_pages(url, access_token):
            fid = folder.get("id")
            if fid and _folder_matches_display(folder, name):
                return fid
            if fid:
                to_scan.append((fid, folder.get("displayName") or ""))

    raise RuntimeError(
        f"Could not find a mail folder with display name {name!r}. "
        "Use INBOX or the exact folder name as shown in Outlook."
    )


def _recipient_to_from_raw(from_obj: dict | None) -> str:
    if not from_obj:
        return ""
    ea = from_obj.get("emailAddress") or {}
    name = (ea.get("name") or "").strip()
    addr = (ea.get("address") or "").strip()
    if name:
        return f'"{name}" <{addr}>'
    return addr


def _message_body_html(body: dict | None) -> str:
    if not body:
        return ""
    content = body.get("content") or ""
    ctype = (body.get("contentType") or "text").lower()
    if ctype == "html":
        return content
    # Plain text — downstream expects HTML file; wrap minimally for consistency
    from html import escape

    return f"<pre>{escape(content)}</pre>"


def _download_attachment_bytes(
    access_token: str, message_id: str, attachment_id: str
) -> bytes:
    safe_mid = urllib.parse.quote(message_id, safe="")
    safe_aid = urllib.parse.quote(attachment_id, safe="")
    url = (
        f"{GRAPH_ROOT}/me/messages/{safe_mid}/attachments/{safe_aid}/$value"
    )
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            return resp.read()
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"Graph attachment download HTTP {e.code}: {detail}"
        ) from e


def _collect_attachments(
    access_token: str, message_id: str
) -> list[tuple[str, bytes]]:
    safe_mid = urllib.parse.quote(message_id, safe="")
    base = f"{GRAPH_ROOT}/me/messages/{safe_mid}/attachments?$top=100"
    out: list[tuple[str, bytes]] = []
    for att in _graph_get_all_pages(base, access_token):
        otype = att.get("@odata.type", "")
        name = att.get("name") or "unnamed_attachment"
        if "fileAttachment" in otype:
            b64 = att.get("contentBytes")
            if b64:
                try:
                    data = base64.b64decode(b64)
                except (ValueError, TypeError):
                    data = b""
            else:
                data = _download_attachment_bytes(
                    access_token, message_id, att["id"]
                )
            if data:
                out.append((name, data))
        # itemAttachment / referenceAttachment: skip (no simple bytes)
    return out


def fetch_emails(
    *,
    mail_folder: str,
    attachments_dir: Path | None = None,
    azure_client_id: str,
    azure_tenant_id: str = "common",
    auth_flow: str = "interactive",
    token_cache_path: Path | None = None,
    force_full_graph_auth: bool = False,
) -> list[EmailMessage]:
    """List every message in *mail_folder*, build :class:`EmailMessage` instances.

    Authentication is delegated (interactive browser or device code). Tokens are
    cached on disk for subsequent runs.

    If *force_full_graph_auth* is True (e.g. when ``DEMO_MODE=1`` in ``mainRunner``),
    silent token refresh is skipped and interactive login uses ``prompt=login`` so
    each run goes through full sign-in and account/MFA as required by Microsoft.
    """
    if not azure_client_id.strip():
        raise ValueError("azure_client_id is required")

    cache_path = token_cache_path or Path(".graph_token_cache.bin")
    token = _acquire_token(
        client_id=azure_client_id.strip(),
        tenant_id=(azure_tenant_id or "common").strip(),
        cache_path=cache_path,
        auth_flow=auth_flow,
        force_interactive=force_full_graph_auth,
    )

    folder_id = urllib.parse.quote(_resolve_folder_id(token, mail_folder), safe="")
    list_url = (
        f"{GRAPH_ROOT}/me/mailFolders/{folder_id}/messages"
        f"?$select=id,subject,from,hasAttachments"
        f"&$orderby=receivedDateTime%20asc"
        f"&$top=50"
    )
    summaries = _graph_get_all_pages(list_url, token)

    messages: list[EmailMessage] = []
    for summary in summaries:
        mid = summary.get("id")
        if not mid:
            continue
        safe_mid = urllib.parse.quote(mid, safe="")
        detail_url = (
            f"{GRAPH_ROOT}/me/messages/{safe_mid}"
            f"?$select=body,from,subject,hasAttachments,sentDateTime,receivedDateTime,toRecipients"
        )
        detail = _graph_get(detail_url, token)

        from_raw = _recipient_to_from_raw(detail.get("from"))
        subject = detail.get("subject") or ""
        body_html = _message_body_html(detail.get("body"))

        to_recs = detail.get("toRecipients") or []
        to_line = ", ".join(
            x for x in (_recipient_to_from_raw(r) for r in to_recs) if x
        )
        header_title = ""
        if to_recs:
            ea0 = (to_recs[0].get("emailAddress") or {})
            header_title = (ea0.get("name") or "").strip() or (
                ea0.get("address") or ""
            ).strip()

        sent_iso = detail.get("sentDateTime") or detail.get("receivedDateTime")
        sent_line = format_graph_datetime_local(sent_iso)

        attachments: list[tuple[str, bytes]] = []
        if detail.get("hasAttachments"):
            attachments = _collect_attachments(token, mid)

        if attachments_dir and attachments:
            saved = save_attachments(attachments, attachments_dir)
            for p in saved:
                print(f"    Saved attachment: {p.name}")

        messages.append(
            EmailMessage(
                from_raw=from_raw,
                subject=subject,
                body_html=body_html,
                attachments=attachments,
                to_line=to_line,
                sent_line=sent_line,
                header_title=header_title,
            )
        )

    return messages
