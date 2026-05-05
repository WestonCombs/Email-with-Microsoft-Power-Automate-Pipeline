"""Fetch mail via Microsoft Graph API (OAuth 2.0 delegated) — replaces IMAP."""

from __future__ import annotations

import base64
import concurrent.futures
import contextlib
import json
import socket
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import msal

from shared import runLogger as RL

from .emailFetcher import (
    EmailMessage,
    format_graph_datetime_local,
    inline_cid_images,
    save_attachments,
)
from .graph_browser_signin_hint import run_blocking_task_with_browser_signin_hint

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
AUTH_ROOT = "https://login.microsoftonline.com"
# Silent MSAL refresh can block indefinitely; cap it so stale/expired sessions fail fast.
SILENT_AUTH_TIMEOUT_S = 30.0
# JSON list/detail calls — keep attachment downloads separate (can be large).
GRAPH_JSON_TIMEOUT_S = 30.0
# MSAL uses requests for authority/instance discovery; default timeout=None can hang forever.
MSAL_HTTP_TIMEOUT_S = 30.0
# Guard MSAL setup/lookup steps that occasionally stall on some Windows installs.
MSAL_APP_INIT_TIMEOUT_S = 15.0
MSAL_GET_ACCOUNTS_TIMEOUT_S = 10.0
# Keep interactive/device sign-in from hanging forever with no visible progress.
INTERACTIVE_SIGNIN_TIMEOUT_S = 5 * 60.0
DEVICE_CODE_SIGNIN_TIMEOUT_S = 10 * 60.0
# Safety guard for malformed pagination loops.
GRAPH_MAX_PAGES = 5000
_NO_PROXY_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


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


def _graph_get_all_pages(
    first_url: str,
    access_token: str,
    *,
    context: str = "Graph list request",
) -> list[dict]:
    rows: list[dict] = []
    url: str | None = first_url
    page_count = 0
    seen_urls: set[str] = set()
    while url:
        if _cancel_check and _cancel_check():
            raise RuntimeError("Login/fetch cancelled by user")
        if url in seen_urls:
            raise RuntimeError(
                f"{context} entered a pagination loop (same nextLink repeated)."
            )
        seen_urls.add(url)
        page_count += 1
        if page_count > GRAPH_MAX_PAGES:
            raise RuntimeError(
                f"{context} exceeded {GRAPH_MAX_PAGES} pages; aborting to avoid hang."
            )
        data = _graph_get(url, access_token)
        if isinstance(data, list):
            rows.extend(data)
            break
        rows.extend(data.get("value") or [])
        url = data.get("@odata.nextLink")
    return rows


_cancel_check = None


@contextlib.contextmanager
def _prefer_ipv4_resolution():
    """Temporarily prefer IPv4 DNS answers to avoid IPv6 connect stalls."""
    orig_getaddrinfo = socket.getaddrinfo

    def _ipv4_first(host, port, family=0, type=0, proto=0, flags=0):
        infos = orig_getaddrinfo(host, port, family, type, proto, flags)
        v4 = [entry for entry in infos if entry[0] == socket.AF_INET]
        return v4 or infos

    socket.getaddrinfo = _ipv4_first
    try:
        yield
    finally:
        socket.getaddrinfo = orig_getaddrinfo


def _run_daemon_with_timeout(fn, *, timeout_s: float, label: str):
    state: dict[str, object] = {"done": False, "value": None, "error": None}
    lock = threading.Lock()
    done = threading.Event()

    def _worker() -> None:
        value = None
        error: BaseException | None = None
        try:
            value = fn()
        except BaseException as e:
            error = e
        with lock:
            state["done"] = True
            state["value"] = value
            state["error"] = error
        done.set()

    threading.Thread(target=_worker, daemon=True).start()
    if not done.wait(timeout_s):
        raise RuntimeError(f"{label} timed out after {timeout_s:.0f}s")
    with lock:
        error = state.get("error")
        value = state.get("value")
    if isinstance(error, BaseException):
        raise error
    return value


def _auth_post_form(url: str, form: dict[str, str], timeout_s: float = 30.0) -> tuple[int, dict]:
    body = urllib.parse.urlencode(form).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )

    def _do_post(open_fn) -> tuple[int, dict]:
        body = urllib.parse.urlencode(form).encode("utf-8")
        try:
            with open_fn(req, timeout=timeout_s) as resp:
                raw = resp.read().decode("utf-8", errors="replace")
                payload = json.loads(raw) if raw else {}
                return int(getattr(resp, "status", 200)), payload
        except urllib.error.HTTPError as e:
            raw = e.read().decode("utf-8", errors="replace")
            try:
                payload = json.loads(raw) if raw else {}
            except json.JSONDecodeError:
                payload = {"error_description": raw or str(e)}
            return int(e.code), payload
        except urllib.error.URLError as e:
            if isinstance(e.reason, TimeoutError):
                raise RuntimeError(
                    f"Microsoft auth request timed out after {timeout_s:.0f}s: {url}"
                ) from e
            raise RuntimeError(f"Microsoft auth request failed: {e}") from e

    def _run_attempt(open_fn, mode: str) -> tuple[int, dict]:
        return _run_daemon_with_timeout(
            lambda: _do_post(open_fn),
            timeout_s=timeout_s,
            label=f"{mode} auth request",
        )

    errors: list[str] = []
    attempts = [
        ("direct", _NO_PROXY_OPENER.open),
        ("system-proxy", urllib.request.urlopen),
    ]
    for mode, open_fn in attempts:
        try:
            return _run_attempt(open_fn, mode)
        except RuntimeError as e:
            errors.append(str(e))
            continue

    raise RuntimeError(
        f"Microsoft auth request failed for {url}. Attempts: {' | '.join(errors)}"
    )


def _acquire_token_device_code_fallback(
    *,
    client_id: str,
    tenant_id: str,
    cancel_check,
) -> str:
    scope = "offline_access Mail.Read"
    tenant = (tenant_id or "common").strip() or "common"
    device_url = f"{AUTH_ROOT}/{tenant}/oauth2/v2.0/devicecode"
    token_url = f"{AUTH_ROOT}/{tenant}/oauth2/v2.0/token"

    code_status, code_payload = _auth_post_form(
        device_url,
        {"client_id": client_id, "scope": scope},
        timeout_s=15.0,
    )
    if code_status >= 400:
        err = code_payload.get("error_description") or code_payload.get("error") or str(code_payload)
        raise RuntimeError(f"Fallback device-code auth failed to start: {err}")

    device_code = str(code_payload.get("device_code") or "").strip()
    if not device_code:
        raise RuntimeError(f"Fallback device-code auth returned no device_code: {code_payload}")

    msg = str(code_payload.get("message") or "").strip()
    if msg:
        print(msg, file=sys.stderr)
    print("  [Graph] Waiting for fallback device-code sign-in...", flush=True)
    RL.log("emailFetching", f"{RL.ts()}  stage=auth_device_code_wait_fallback")

    interval = max(2, int(code_payload.get("interval") or 5))
    expires_in = max(60, int(code_payload.get("expires_in") or 900))
    deadline = time.monotonic() + expires_in
    while time.monotonic() < deadline:
        if cancel_check and cancel_check():
            raise RuntimeError("Login cancelled by user during fallback device code sign-in")

        status, payload = _auth_post_form(
            token_url,
            {
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": client_id,
                "device_code": device_code,
            },
            timeout_s=15.0,
        )
        if status < 400 and payload.get("access_token"):
            return str(payload["access_token"])

        err = str(payload.get("error") or "").strip().lower()
        if err == "authorization_pending":
            time.sleep(interval)
            continue
        if err == "slow_down":
            interval = min(interval + 2, 15)
            time.sleep(interval)
            continue
        if err:
            desc = payload.get("error_description") or payload.get("error")
            raise RuntimeError(f"Fallback device-code sign-in failed: {desc}")
        raise RuntimeError(f"Fallback device-code sign-in failed: {payload}")

    raise RuntimeError(
        f"Fallback device-code sign-in timed out after {expires_in:.0f}s."
    )


def _call_with_timeout(fn, *, timeout_s: float, label: str):
    try:
        return _run_daemon_with_timeout(fn, timeout_s=timeout_s, label=label)
    except RuntimeError as e:
        raise RuntimeError(
            f"{label} timed out after {timeout_s:.0f}s. "
            "Try again, or switch auth flow to device_code in .env."
        ) from e


def _acquire_token(
    *,
    client_id: str,
    tenant_id: str,
    cache_path: Path,
    auth_flow: str,
    force_interactive: bool = False,
    base_dir: Path | None = None,
) -> str:
    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        try:
            cache.deserialize(cache_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            pass

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    interactive_cache_payload = ""

    def _new_public_client_app(cache_obj: msal.SerializableTokenCache):
        return msal.PublicClientApplication(
            client_id,
            authority=authority,
            token_cache=cache_obj,
            timeout=MSAL_HTTP_TIMEOUT_S,
            # Broker/WAM can hang on some machines; force browser/device flow instead.
            enable_broker_on_windows=False,
            # Skip authority instance discovery network roundtrip for known authority.
            instance_discovery=False,
        )

    RL.log("emailFetching", f"{RL.ts()}  stage=msal_app_init_start")
    try:
        app = _call_with_timeout(
            lambda: _new_public_client_app(cache),
            timeout_s=MSAL_APP_INIT_TIMEOUT_S,
            label="Microsoft auth client initialization",
        )
    except RuntimeError as e:
        RL.log("emailFetching", f"{RL.ts()}  stage=msal_app_init_timeout")
        print(
            "WARNING: Microsoft auth client initialization stalled. "
            "Switching to fallback device-code sign-in...",
            file=sys.stderr,
            flush=True,
        )
        return _acquire_token_device_code_fallback(
            client_id=client_id,
            tenant_id=tenant_id,
            cancel_check=_cancel_check,
        )
    RL.log("emailFetching", f"{RL.ts()}  stage=msal_app_init_done")
    scopes = ["Mail.Read"]

    if not force_interactive:
        if _cancel_check and _cancel_check():
            raise RuntimeError("Login cancelled by user")
        RL.log("emailFetching", f"{RL.ts()}  stage=silent_accounts_start")
        try:
            accounts = _call_with_timeout(
                app.get_accounts,
                timeout_s=MSAL_GET_ACCOUNTS_TIMEOUT_S,
                label="Reading cached Microsoft accounts",
            )
        except RuntimeError:
            print(
                "WARNING: Reading cached Microsoft accounts timed out. "
                "Skipping silent sign-in and opening interactive/device sign-in.",
                file=sys.stderr,
            )
            RL.log("emailFetching", f"{RL.ts()}  stage=silent_accounts_timeout")
            accounts = []
        RL.log(
            "emailFetching",
            f"{RL.ts()}  stage=silent_accounts_done  count={len(accounts)}",
        )
        if accounts:
            result: dict | None
            executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
            try:
                RL.log("emailFetching", f"{RL.ts()}  stage=silent_token_start")
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
                RL.log("emailFetching", f"{RL.ts()}  stage=silent_token_timeout")
                try:
                    cache_path.unlink(missing_ok=True)
                except OSError:
                    pass
                cache = msal.SerializableTokenCache()
                app = _call_with_timeout(
                    lambda: _new_public_client_app(cache),
                    timeout_s=MSAL_APP_INIT_TIMEOUT_S,
                    label="Microsoft auth client re-initialization",
                )
                result = None
            finally:
                # Do not wait for a stuck acquire_token_silent worker (with-block would).
                executor.shutdown(wait=False)
            RL.log(
                "emailFetching",
                f"{RL.ts()}  stage=silent_token_done  ok={1 if result and result.get('access_token') else 0}",
            )

            if result and result.get("access_token"):
                if cache.has_state_changed:
                    cache_path.parent.mkdir(parents=True, exist_ok=True)
                    cache_path.write_text(cache.serialize(), encoding="utf-8")
                return result["access_token"]

    auth_flow_norm = (auth_flow or "interactive").strip().lower()
    if auth_flow_norm == "device_code":
        if _cancel_check and _cancel_check():
            raise RuntimeError("Login cancelled by user")
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(
                "Device code flow failed to start: "
                + json.dumps(flow, indent=2)
            )
        print(flow["message"], file=sys.stderr)
        print(
            "  [Graph] Waiting for device-code sign-in to complete...",
            flush=True,
        )
        RL.log("emailFetching", f"{RL.ts()}  stage=auth_device_code_wait")
        result = run_blocking_task_with_browser_signin_hint(
            lambda: app.acquire_token_by_device_flow(flow),
            cancel_check=_cancel_check,
            cancel_message="Login cancelled by user during device code sign-in",
            timeout_seconds=DEVICE_CODE_SIGNIN_TIMEOUT_S,
            base_dir=base_dir,
        )
    else:
        # prompt=login: skip SSO for this request so the user can pick another
        # account and complete MFA (used when DEMO_MODE forces full sign-in).
        if _cancel_check and _cancel_check():
            raise RuntimeError("Login cancelled by user")
        interactive_kw: dict = {"scopes": scopes}
        if force_interactive:
            interactive_kw["prompt"] = "login"

        def _clear_cached_signin() -> None:
            try:
                cache_path.unlink(missing_ok=True)
            except OSError:
                pass

        def _interactive_attempt() -> dict:
            attempt_cache = msal.SerializableTokenCache()
            if cache_path.exists():
                try:
                    attempt_cache.deserialize(cache_path.read_text(encoding="utf-8"))
                except (OSError, ValueError):
                    pass
            attempt_app = _call_with_timeout(
                lambda: _new_public_client_app(attempt_cache),
                timeout_s=MSAL_APP_INIT_TIMEOUT_S,
                label="Microsoft interactive client initialization",
            )
            attempt_result = attempt_app.acquire_token_interactive(**interactive_kw)
            return {
                "result": attempt_result,
                "token_cache": (
                    attempt_cache.serialize()
                    if attempt_cache.has_state_changed
                    else ""
                ),
            }

        print(
            "  [Graph] Waiting for interactive Microsoft sign-in in your browser...",
            flush=True,
        )
        RL.log("emailFetching", f"{RL.ts()}  stage=auth_interactive_wait")
        wrapped_result = run_blocking_task_with_browser_signin_hint(
            _interactive_attempt,
            cancel_check=_cancel_check,
            cancel_message="Login cancelled by user during interactive sign-in",
            timeout_seconds=INTERACTIVE_SIGNIN_TIMEOUT_S,
            base_dir=base_dir,
            allow_retry=True,
            on_retry=_clear_cached_signin,
        )
        result = (
            wrapped_result.get("result")
            if isinstance(wrapped_result.get("result"), dict)
            else wrapped_result
        )
        interactive_cache_payload = str(wrapped_result.get("token_cache") or "")

    if not result or not result.get("access_token"):
        err = result.get("error_description") or result.get("error") if result else "unknown"
        raise RuntimeError(f"Could not acquire Graph token: {err}")

    if interactive_cache_payload:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(interactive_cache_payload, encoding="utf-8")
    elif cache.has_state_changed:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(cache.serialize(), encoding="utf-8")

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

        for folder in _graph_get_all_pages(
            url,
            access_token,
            context=f"Mail folder lookup for {name!r}",
        ):
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
) -> tuple[list[tuple[str, bytes]], dict[str, tuple[str, bytes]]]:
    safe_mid = urllib.parse.quote(message_id, safe="")
    base = f"{GRAPH_ROOT}/me/messages/{safe_mid}/attachments?$top=100"
    out: list[tuple[str, bytes]] = []
    inline_by_cid: dict[str, tuple[str, bytes]] = {}
    for att in _graph_get_all_pages(
        base,
        access_token,
        context=f"Attachment listing for message {message_id[:24]}",
    ):
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
                cid = str(att.get("contentId") or "").strip()
                if cid:
                    inline_by_cid[cid] = (
                        str(att.get("contentType") or "application/octet-stream"),
                        data,
                    )
        # itemAttachment / referenceAttachment: skip (no simple bytes)
    return out, inline_by_cid


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
) -> list[EmailMessage]:
    """List every message in *mail_folder*, build :class:`EmailMessage` instances.

    Authentication is delegated (interactive browser or device code). Tokens are
    cached on disk for subsequent runs; silent refresh from the cache is always
    tried before any browser or device-code prompt.

    If *force_full_graph_auth* is True (e.g. when ``DEMO_MODE=1`` in ``mainRunner``),
    silent token refresh is skipped and interactive login uses ``prompt=login`` so
    each run goes through full sign-in and account/MFA as required by Microsoft.

    *base_dir* enables the sign-in hint window's Cancel button to request a pipeline
    stop via the same cancel file as the launcher Stop control.
    """
    global _cancel_check
    _cancel_check = cancel_check
    try:
        if not azure_client_id.strip():
            raise ValueError("azure_client_id is required")

        with _prefer_ipv4_resolution():
            print("  [Graph] Acquiring Microsoft access token...", flush=True)
            RL.log("emailFetching", f"{RL.ts()}  stage=token_acquire_start")
            cache_path = token_cache_path or Path(".graph_token_cache.bin")
            token = _acquire_token(
                client_id=azure_client_id.strip(),
                tenant_id=(azure_tenant_id or "common").strip(),
                cache_path=cache_path,
                auth_flow=auth_flow,
                force_interactive=force_full_graph_auth,
                base_dir=base_dir,
            )
            print("  [Graph] Access token acquired.", flush=True)
            RL.log("emailFetching", f"{RL.ts()}  stage=token_acquire_done")

            print(f"  [Graph] Resolving folder {mail_folder!r}...", flush=True)
            RL.log("emailFetching", f"{RL.ts()}  stage=folder_resolve_start  folder={mail_folder!r}")
            folder_id_raw = _resolve_folder_id(token, mail_folder)
            folder_id = urllib.parse.quote(folder_id_raw, safe="")
            print("  [Graph] Folder resolved.", flush=True)
            RL.log("emailFetching", f"{RL.ts()}  stage=folder_resolve_done  folder={mail_folder!r}")
            list_url = (
                f"{GRAPH_ROOT}/me/mailFolders/{folder_id}/messages"
                f"?$select=id,subject,from,hasAttachments"
                f"&$orderby=receivedDateTime%20asc"
                f"&$top=50"
            )
            print("  [Graph] Listing message summaries...", flush=True)
            RL.log("emailFetching", f"{RL.ts()}  stage=message_list_start  folder={mail_folder!r}")
            summaries = _graph_get_all_pages(
                list_url,
                token,
                context=f"Message listing for folder {mail_folder!r}",
            )
            print(f"  [Graph] Message summaries found: {len(summaries)}", flush=True)
            RL.log(
                "emailFetching",
                f"{RL.ts()}  stage=message_list_done  folder={mail_folder!r}  count={len(summaries)}",
            )

            messages: list[EmailMessage] = []
            total_summaries = len(summaries)
            for idx, summary in enumerate(summaries, start=1):
                if _cancel_check and _cancel_check():
                    raise RuntimeError("Email fetch cancelled by user")
                if idx == 1 or idx == total_summaries or idx % 10 == 0:
                    print(
                        f"  [Graph] Loading message {idx}/{total_summaries}...",
                        flush=True,
                    )
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
                received_iso = detail.get("receivedDateTime") or detail.get("sentDateTime")
                sent_line = format_graph_datetime_local(sent_iso)

                attachments: list[tuple[str, bytes]] = []
                inline_cid_payloads: dict[str, tuple[str, bytes]] = {}
                if detail.get("hasAttachments"):
                    attachments, inline_cid_payloads = _collect_attachments(token, mid)

                if inline_cid_payloads:
                    body_html, cid_replaced, cid_unresolved = inline_cid_images(
                        body_html,
                        inline_cid_payloads,
                    )
                    if cid_replaced:
                        print(f"    Inlined {cid_replaced} CID image(s) for HTML/PDF rendering")
                        RL.log(
                            "emailFetching",
                            f"{RL.ts()}  message={mid}  inlined_cid_images={cid_replaced}",
                        )
                    if cid_unresolved:
                        print(
                            "    WARNING: "
                            f"{cid_unresolved} CID image reference(s) could not be inlined"
                        )
                        RL.log(
                            "emailFetching",
                            f"{RL.ts()}  WARNING: message={mid} unresolved_cid_images={cid_unresolved}",
                        )

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
                        sent_datetime_iso=str(sent_iso or "").strip(),
                        received_datetime_iso=str(received_iso or "").strip(),
                    )
                )
            return messages
    finally:
        _cancel_check = None
