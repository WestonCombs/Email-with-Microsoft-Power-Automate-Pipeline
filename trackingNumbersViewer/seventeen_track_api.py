"""Optional 17TRACK REST API (requires free API key from https://user.17track.net)."""

from __future__ import annotations

import json
import os
import sys
import time
import urllib.error
import urllib.request

# v2.4 endpoints (used by official docs / integrations)
_BASE = "https://api.17track.net/track/v2.4"
_REGISTER = f"{_BASE}/register"
_GET = f"{_BASE}/gettrackinfo"

_PYTHON_FILES = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PYTHON_FILES not in sys.path:
    sys.path.insert(0, _PYTHON_FILES)

from htmlHandler.carrier_urls import infer_carrier  # noqa: E402

# 17TRACK carrier keys (see https://res.17track.net/asset/carrier/info/apicarrier.all.json)
_CARRIER_17TRACK = {
    "UPS": 100002,
    "FedEx": 100003,
    "USPS": 21051,
    "DHL": 100001,  # DHL Express (common for intl numeric AWBs)
}


def _carrier_code_for_number(tracking_id: str) -> int:
    """Non-zero when we can infer carrier; 0 means “auto” (often fails on plain numerics)."""
    label = infer_carrier((tracking_id or "").strip())
    return int(_CARRIER_17TRACK.get(label, 0))


def api_key_from_env() -> str | None:
    k = (os.getenv("SEVENTEEN_TRACK_API_KEY") or os.getenv("17TRACK_API_KEY") or "").strip()
    return k or None


def post_track_v24(api_key: str, endpoint: str, body: list | dict, timeout: float = 25.0) -> dict:
    """POST JSON to v2.4 *endpoint* path (``register`` or ``gettrackinfo``)."""
    url = f"{_BASE}/{endpoint}"
    headers = {
        "Content-Type": "application/json; charset=utf-8",
        "17token": api_key,
    }

    def _post(url_full: str, body_obj: object) -> dict:
        data = json.dumps(body_obj).encode("utf-8")
        req = urllib.request.Request(url_full, data=data, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as e:
            raw = e.read().decode("utf-8", errors="replace")
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return {"_raw": raw, "parse_error": True}

    return _post(url, body)


def get_trackinfo_only(api_key: str, tracking_numbers: list[str], timeout: float = 25.0) -> dict:
    """POST gettrackinfo only (no register)."""
    nums = [n.strip() for n in tracking_numbers if n.strip()]
    if not nums:
        return {}
    payload = [{"number": n, "carrier": _carrier_code_for_number(n)} for n in nums]
    return post_track_v24(api_key, "gettrackinfo", payload, timeout=timeout)


def register_only(api_key: str, tracking_numbers: list[str], timeout: float = 25.0) -> dict:
    """POST register only."""
    nums = [n.strip() for n in tracking_numbers if n.strip()]
    if not nums:
        return {}
    payload = [{"number": n, "carrier": _carrier_code_for_number(n)} for n in nums]
    return post_track_v24(api_key, "register", payload, timeout=timeout)


def register_and_fetch(api_key: str, tracking_numbers: list[str], timeout: float = 25.0) -> dict:
    """Register numbers then fetch track info.

    17TRACK requires a successful ``/register`` before ``/gettrackinfo``. Plain numeric IDs
    (e.g. FedEx 12-digit) often fail with ``carrier: 0`` auto-detect; we pass explicit carrier
    codes from :func:`infer_carrier` / ``_CARRIER_17TRACK``.
    """
    nums = [n.strip() for n in tracking_numbers if n.strip()]
    payload_reg = [{"number": n, "carrier": _carrier_code_for_number(n)} for n in nums]
    if not payload_reg:
        return {}

    reg = post_track_v24(api_key, "register", payload_reg, timeout=timeout)
    # Server may need a moment after register before gettrackinfo accepts the number.
    time.sleep(2.0)

    get_resp = post_track_v24(api_key, "gettrackinfo", payload_reg, timeout=timeout)

    # If get says “register first” but register looked OK, retry once with a slightly longer wait.
    if _get_rejected_register_first(get_resp, nums):
        time.sleep(2.5)
        get_resp = post_track_v24(api_key, "gettrackinfo", payload_reg, timeout=timeout)

    # Attach register response for debugging when get only shows rejected.
    if isinstance(get_resp, dict) and isinstance(reg, dict):
        get_resp["_register_response"] = reg

    return get_resp


def _get_rejected_register_first(resp: dict, numbers: list[str]) -> bool:
    if not isinstance(resp, dict) or resp.get("parse_error"):
        return False
    data = resp.get("data")
    if not isinstance(data, dict):
        return False
    rej = data.get("rejected")
    if not isinstance(rej, list):
        return False
    want = {n.strip() for n in numbers}
    for item in rej:
        if not isinstance(item, dict):
            continue
        if str(item.get("number", "")).strip() not in want:
            continue
        err = item.get("error") or {}
        if isinstance(err, dict) and int(err.get("code", 0)) == -18019902:
            return True
    return False


def summarize_for_number(api_response: dict, number: str) -> str:
    """Human-readable one-line status from 17TRACK gettrackinfo response."""
    if not api_response or api_response.get("parse_error"):
        return (api_response or {}).get("_raw", "Invalid API response")[:500]

    code = api_response.get("code")
    if code is not None and int(code) != 0:
        msg = api_response.get("message") or api_response.get("msg") or api_response
        return str(msg)[:800]

    # Typical: {"data": {"accepted": [...], "rejected": [...]}}
    data = api_response.get("data")
    if isinstance(data, dict):
        for key in ("accepted", "registered", "pending", "delivered"):
            items = data.get(key)
            if isinstance(items, list):
                for item in items:
                    if isinstance(item, dict) and str(item.get("number", "")).strip() == number.strip():
                        return _one_item_summary(item)
        if isinstance(data.get("items"), list):
            for item in data["items"]:
                if isinstance(item, dict) and str(item.get("number", "")).strip() == number.strip():
                    return _one_item_summary(item)
    if isinstance(data, list):
        for item in data:
            if isinstance(item, dict) and str(item.get("number", "")).strip() == number.strip():
                return _one_item_summary(item)

    errs = api_response.get("data", {}).get("rejected") if isinstance(api_response.get("data"), dict) else None
    if isinstance(errs, list):
        for item in errs:
            if isinstance(item, dict) and str(item.get("number", "")).strip() == number.strip():
                e = item.get("error")
                if isinstance(e, dict):
                    return e.get("message") or str(e)
                return f"Rejected: {e}"

    reg = api_response.get("_register_response")
    if isinstance(reg, dict) and reg.get("data"):
        rj = reg["data"].get("rejected") if isinstance(reg["data"], dict) else None
        if isinstance(rj, list):
            for item in rj:
                if isinstance(item, dict) and str(item.get("number", "")).strip() == number.strip():
                    e = item.get("error")
                    if isinstance(e, dict):
                        return f"(register) {e.get('message', e)}"
                    return f"(register) {e}"

    return json.dumps(api_response, ensure_ascii=False)[:800]


def _one_item_summary(item: dict) -> str:
    track = item.get("track_info") or item.get("track") or {}
    if isinstance(track, dict):
        latest = track.get("latest_event") or track.get("latest_status") or {}
        if isinstance(latest, dict):
            st = latest.get("description") or latest.get("status") or latest.get("sub_status")
            if st:
                return str(st)
        shipping = track.get("shipping_info") or {}
        if isinstance(shipping, dict):
            stat = shipping.get("status") or shipping.get("status_description")
            if stat:
                return str(stat)
    return "Received — see raw JSON for details"
