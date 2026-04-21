"""Smart 17TRACK fetch: classify outcomes, skip redundant register, JSON cache, throttle."""

from __future__ import annotations

import json
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from collections.abc import Callable
from pathlib import Path
from typing import Any

_PYTHON_FILES = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PYTHON_FILES not in sys.path:
    sys.path.insert(0, _PYTHON_FILES)

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None  # type: ignore[assignment]

from htmlHandler.carrier_urls import infer_carrier  # noqa: E402
from trackingNumbersViewer import seventeen_track_api as api  # noqa: E402

# Already registered (17TRACK)
_CODE_ALREADY_REGISTERED = -18019901
NOTFOUND_RECHECK_DAYS = 14

# Minimum seconds between network gettrackinfo for same number (unless force_refresh)
MIN_FETCH_INTERVAL_SEC = float(os.getenv("SEVENTEEN_TRACK_MIN_INTERVAL_SEC", "3600"))

_CACHE_SUBDIR = "tracking_status_cache"


def _project_root() -> Path:
    if load_dotenv:
        load_dotenv(Path(_PYTHON_FILES) / ".env", override=False)
    base = os.getenv("BASE_DIR")
    if base:
        return Path(base).expanduser().resolve()
    return Path(_PYTHON_FILES).resolve().parent


def _cache_dir() -> Path:
    d = _project_root() / "email_contents" / _CACHE_SUBDIR
    d.mkdir(parents=True, exist_ok=True)
    return d


def _normalize_number(n: str) -> str:
    return re.sub(r"\s+", "", (n or "").strip())


def cache_path_for_number(tracking_id: str) -> Path:
    n = _normalize_number(tracking_id)
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", n)[:200] or "unknown"
    return _cache_dir() / f"{safe}.json"


def _parse_iso(ts: str | None) -> float | None:
    if not ts:
        return None
    try:
        s = ts.replace("Z", "+00:00")
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.timestamp()
    except Exception:
        return None


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse_purchase_datetime(value: object) -> datetime | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


def _notfound_cutoff_iso(purchase_datetime: object) -> str | None:
    dt = _parse_purchase_datetime(purchase_datetime)
    if dt is None:
        return None
    return (dt + timedelta(days=NOTFOUND_RECHECK_DAYS)).strftime("%Y-%m-%dT%H:%M:%SZ")


def tracking_is_greyed_out(tracking_id: str) -> bool:
    cached = load_cache(tracking_id)
    if not isinstance(cached, dict) or not bool(cached.get("notfound_give_up")):
        return False
    gr = cached.get("last_get_response")
    ti = extract_track_info(gr, tracking_id) if isinstance(gr, dict) else None
    if classify_outcome(ti) == "no_data":
        return False
    return True


def extract_track_info(get_resp: dict, number: str) -> dict[str, Any] | None:
    if not isinstance(get_resp, dict):
        return None
    data = get_resp.get("data")
    if not isinstance(data, dict):
        return None
    acc = data.get("accepted")
    if not isinstance(acc, list):
        return None
    want = _normalize_number(number)
    for item in acc:
        if not isinstance(item, dict):
            continue
        if _normalize_number(str(item.get("number", ""))) != want:
            continue
        ti = item.get("track_info")
        if isinstance(ti, dict):
            return ti
    return None


def extract_accepted_item(get_resp: dict, number: str) -> dict[str, Any] | None:
    """The ``accepted[]`` entry for *number* (carrier code lives here)."""
    if not isinstance(get_resp, dict):
        return None
    data = get_resp.get("data")
    if not isinstance(data, dict):
        return None
    acc = data.get("accepted")
    if not isinstance(acc, list):
        return None
    want = _normalize_number(number)
    for item in acc:
        if not isinstance(item, dict):
            continue
        if _normalize_number(str(item.get("number", ""))) != want:
            continue
        return item
    return None


# 17TRACK numeric carrier keys → short display (subset of api._CARRIER_17TRACK).
_CARRIER_CODE_DISPLAY: dict[int, str] = {
    int(v): k for k, v in api._CARRIER_17TRACK.items()
}


def _carrier_name_from_track_info_providers(track_info: dict[str, Any] | None) -> str | None:
    if not isinstance(track_info, dict):
        return None
    provs = track_info.get("tracking", {}).get("providers")
    if not isinstance(provs, list) or not provs:
        return None
    p0 = provs[0]
    if not isinstance(p0, dict):
        return None
    for key in ("provider", "Provider"):
        pr = p0.get(key)
        if isinstance(pr, dict):
            for nk in ("name", "alias", "Name", "alias_name"):
                v = pr.get(nk)
                if isinstance(v, str) and v.strip():
                    return v.strip()
    for nk in ("service_provider", "carrier_name", "provider_name", "Courier"):
        v = p0.get(nk)
        if isinstance(v, str) and v.strip():
            return v.strip()
    return None


def resolve_carrier_display(
    tracking_id: str,
    get_resp: dict[str, Any] | None,
    track_info: dict[str, Any] | None,
) -> str:
    """Human-readable carrier from 17TRACK ``gettrackinfo`` + ``track_info``; else pattern guess."""
    n = _normalize_number(tracking_id)
    item = extract_accepted_item(get_resp, n) if get_resp else None
    if item:
        raw_c = item.get("carrier")
        try:
            ic = int(raw_c) if raw_c is not None and str(raw_c).strip() != "" else None
        except (TypeError, ValueError):
            ic = None
        if ic and ic in _CARRIER_CODE_DISPLAY:
            return _CARRIER_CODE_DISPLAY[ic]

    pn = _carrier_name_from_track_info_providers(track_info)
    if pn:
        return pn

    return infer_carrier(tracking_id)


def carrier_display_for_number(tracking_id: str) -> str:
    """Carrier label for UI: prefers ``carrier_display`` on disk cache, else resolves from JSON."""
    c = load_cache(tracking_id)
    if isinstance(c, dict):
        cd = c.get("carrier_display")
        if isinstance(cd, str) and cd.strip():
            return cd.strip()
        gr = c.get("last_get_response")
        if isinstance(gr, dict):
            ti = extract_track_info(gr, tracking_id)
            return resolve_carrier_display(tracking_id, gr, ti)
    return resolve_carrier_display(tracking_id, None, None)


def recipient_location_line(track_info: dict[str, Any] | None) -> str:
    """Format ``shipping_info.recipient_address`` for display (not shipper / last-scan location)."""
    if not isinstance(track_info, dict):
        return ""
    si = track_info.get("shipping_info")
    if not isinstance(si, dict):
        return ""
    ra = si.get("recipient_address")
    if isinstance(ra, str) and ra.strip():
        return ra.strip()
    if isinstance(ra, dict):
        for k in ("formatted", "formated", "formatted_address"):
            v = ra.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
        city = str(ra.get("city") or "").strip()
        state = str(ra.get("state") or "").strip()
        country = str(ra.get("country") or "").strip()
        parts = [p for p in (city, state, country) if p]
        if parts:
            return ", ".join(parts)
    return ""


def _latest_event_location(track_info: dict[str, Any]) -> str:
    evs = track_info.get("tracking", {}).get("providers", [])
    if not isinstance(evs, list) or not evs:
        return ""
    latest = ""
    for p in evs:
        if not isinstance(p, dict):
            continue
        events = p.get("events", [])
        if not isinstance(events, list):
            continue
        for e in events:
            if not isinstance(e, dict):
                continue
            loc = str(e.get("location") or "").strip()
            if loc:
                latest = loc
    return latest


def classify_outcome(track_info: dict[str, Any] | None) -> str:
    """Return ``no_data`` | ``dead_invalid`` | ``terminal`` | ``active``.

    ``no_data`` means 17TRACK has not returned a ``track_info`` payload yet (including
    the UI label ``No data``). It must not use the NotFound two-week / give-up path.
    """
    if not track_info:
        return "no_data"
    latest_status = str(track_info.get("latest_status", {}).get("status", "") or "").lower()
    sub_status = str(track_info.get("latest_status", {}).get("sub_status", "") or "").lower()
    if latest_status in ("notfound", "not_found", "unknown") or sub_status in ("notfound",):
        return "dead_invalid"
    terminal_markers = (
        "delivered",
        "returned",
        "exception",
        "failed",
        "expired",
        "lost",
    )
    for m in terminal_markers:
        if m in latest_status or m in sub_status:
            return "terminal"
    return "active"


def _omit_redundant_sub_status(status: str, sub: str) -> str:
    """Drop sub-status noise such as ``Delivered_Other`` when the main status already says Delivered."""
    if not sub:
        return ""
    st = status.replace(" ", "_").strip()
    su = sub.replace(" ", "_").strip()
    if not st:
        return sub
    st_l, su_l = st.lower(), su.lower()
    if su_l == f"{st_l}_other":
        return ""
    if su_l.endswith("_other") and su_l.rsplit("_other", 1)[0] == st_l:
        return ""
    return sub


def build_quick_status_label(track_info: dict[str, Any] | None) -> str:
    if not track_info:
        return "No data"
    latest = track_info.get("latest_status") or {}
    if not isinstance(latest, dict):
        latest = {}
    status = str(latest.get("status") or "").strip()
    sub = str(latest.get("sub_status") or "").strip()
    sub = _omit_redundant_sub_status(status, sub)
    loc = recipient_location_line(track_info)
    parts = [p for p in (status, sub) if p]
    head = " — ".join(parts) if parts else "Unknown"
    if loc:
        return f"{head} ({loc})"
    return head


def _register_says_already_registered(reg_resp: dict) -> bool:
    if not isinstance(reg_resp, dict):
        return False
    data = reg_resp.get("data")
    if not isinstance(data, dict):
        return False
    rej = data.get("rejected")
    if not isinstance(rej, list):
        return False
    for item in rej:
        if not isinstance(item, dict):
            continue
        if int(item.get("error", {}).get("code", 0) or 0) == _CODE_ALREADY_REGISTERED:
            return True
    return False


def _get_needs_register(get_resp: dict, number: str) -> bool:
    if not isinstance(get_resp, dict):
        return True
    data = get_resp.get("data")
    if not isinstance(data, dict):
        return True
    acc = data.get("accepted")
    if isinstance(acc, list) and acc:
        for item in acc:
            if isinstance(item, dict) and _normalize_number(str(item.get("number", ""))) == _normalize_number(number):
                if item.get("track_info") is not None:
                    return False
    rej = data.get("rejected")
    if isinstance(rej, list):
        for item in rej:
            if not isinstance(item, dict):
                continue
            if _normalize_number(str(item.get("number", ""))) != _normalize_number(number):
                continue
            code = int(item.get("error", {}).get("code", 0) or 0)
            # -18019900 = register first (typical)
            if code in (-18019900, -18019902):
                return True
    return False


def _accepted_missing_track_info(get_resp: dict, number: str) -> bool:
    """True when the number is in ``accepted`` but ``track_info`` is still empty (carrier pending)."""
    if not isinstance(get_resp, dict):
        return False
    data = get_resp.get("data")
    if not isinstance(data, dict):
        return False
    acc = data.get("accepted")
    if not isinstance(acc, list):
        return False
    for item in acc:
        if not isinstance(item, dict):
            continue
        if _normalize_number(str(item.get("number", ""))) != _normalize_number(number):
            continue
        ti = item.get("track_info")
        if ti is None:
            return True
        if isinstance(ti, dict) and not ti:
            return True
        return False
    return False


def load_cache(tracking_id: str) -> dict[str, Any] | None:
    p = cache_path_for_number(tracking_id)
    if not p.is_file():
        return None
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def save_cache(tracking_id: str, payload: dict[str, Any]) -> None:
    p = cache_path_for_number(tracking_id)
    tmp = p.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
    tmp.replace(p)


def quick_status_from_cache(tracking_id: str) -> str | None:
    c = load_cache(tracking_id)
    if not c:
        return None
    return c.get("quick_status_label") or None


def fetch_tracking_smart(
    api_key: str,
    tracking_id: str,
    *,
    force_refresh: bool = False,
    purchase_datetime: object = None,
    timeout: float = 25.0,
) -> dict[str, Any]:
    """Fetch gettrackinfo with smart register skip and JSON cache.

    Returns a dict with keys: ``get_response``, ``track_info``, ``outcome``,
    ``quick_status_label``, ``from_cache`` (bool), ``cached`` (bool if served from disk without network).
    """
    n = _normalize_number(tracking_id)
    if not n:
        return {"error": "empty_number", "get_response": {}, "track_info": None}

    cached = load_cache(n)
    if cached and not force_refresh:
        last_ts = _parse_iso(cached.get("last_fetch_iso"))
        if last_ts is not None and (time.time() - last_ts) < MIN_FETCH_INTERVAL_SEC:
            gr = cached.get("last_get_response") or {}
            ti = extract_track_info(gr if isinstance(gr, dict) else {}, n)
            outcome = classify_outcome(ti)
            greyed = bool(cached.get("notfound_give_up"))
            if outcome == "no_data":
                greyed = False
            return {
                "get_response": gr,
                "track_info": ti,
                "outcome": outcome,
                "quick_status_label": cached.get("quick_status_label") or build_quick_status_label(ti),
                "carrier_display": cached.get("carrier_display")
                or resolve_carrier_display(n, gr if isinstance(gr, dict) else None, ti),
                "from_cache": True,
                "cached": True,
                "greyed_out": greyed,
            }

    get_resp = api.get_trackinfo_only(api_key, [n], timeout=timeout)
    ti = extract_track_info(get_resp, n)

    did_register = False
    if not ti and _get_needs_register(get_resp, n):
        did_register = True
        reg = api.register_only(api_key, [n], timeout=timeout)
        if _register_says_already_registered(reg):
            get_resp = api.get_trackinfo_only(api_key, [n], timeout=timeout)
        else:
            time.sleep(2.0)
            get_resp = api.get_trackinfo_only(api_key, [n], timeout=timeout)
        ti = extract_track_info(get_resp, n)

    # UPS often returns empty ``track_info`` on the first get after register; accepted-without-track
    # also happens while the carrier is still syncing — one delayed get usually fixes it.
    if not ti and (did_register or _accepted_missing_track_info(get_resp, n)):
        time.sleep(3.0)
        get_retry = api.get_trackinfo_only(api_key, [n], timeout=timeout)
        ti_retry = extract_track_info(get_retry, n)
        if ti_retry:
            get_resp, ti = get_retry, ti_retry

    outcome = classify_outcome(ti)
    label = build_quick_status_label(ti)
    carrier_disp = resolve_carrier_display(n, get_resp, ti)

    payload = {
        "number": n,
        "last_fetch_iso": _now_iso(),
        "outcome": outcome,
        "quick_status_label": label,
        "carrier_display": carrier_disp,
        "last_get_response": get_resp,
    }
    if outcome == "dead_invalid":
        payload["notfound_cutoff_iso"] = _notfound_cutoff_iso(purchase_datetime)
        cutoff_ts = _parse_iso(payload.get("notfound_cutoff_iso"))
        now_ts = time.time()
        if cutoff_ts is not None and now_ts > cutoff_ts:
            payload["notfound_final_check_done"] = True
            payload["notfound_give_up"] = True
        else:
            payload["notfound_final_check_done"] = False
            payload["notfound_give_up"] = False
    else:
        payload.pop("notfound_cutoff_iso", None)
        payload["notfound_final_check_done"] = False
        payload["notfound_give_up"] = False
    save_cache(n, payload)

    return {
        "get_response": get_resp,
        "track_info": ti,
        "outcome": outcome,
        "quick_status_label": label,
        "carrier_display": carrier_disp,
        "from_cache": False,
        "cached": False,
        "greyed_out": bool(payload.get("notfound_give_up")),
    }


def is_delivered(track_info: dict[str, Any] | None) -> bool:
    """True when latest status indicates a delivered shipment (word-boundary match)."""
    if not isinstance(track_info, dict):
        return False
    latest = track_info.get("latest_status") or {}
    if not isinstance(latest, dict):
        return False
    st = str(latest.get("status") or "")
    sub = str(latest.get("sub_status") or "")
    text = f"{st} {sub}"
    if re.search(r"\bnot[- ]?delivered\b|\bun[- ]?delivered\b", text, re.I):
        return False
    return bool(re.search(r"\bdelivered\b", text, re.I))


def iter_unique_tracking_ids(records: list[dict]) -> list[str]:
    """Deduped tracking IDs across JSON records (order preserved)."""
    seen: set[str] = set()
    out: list[str] = []
    for rec in records:
        raw = rec.get("tracking_numbers")
        if not isinstance(raw, list):
            continue
        for x in raw:
            if not isinstance(x, str):
                continue
            n = _normalize_number(x)
            if not n or n in seen:
                continue
            seen.add(n)
            out.append(n)
    return out


def _tracking_purchase_dates(records: list[dict]) -> dict[str, str]:
    out: dict[str, str] = {}
    for rec in records:
        cat = str(rec.get("email_category") or "").strip()
        if cat in ("POD", "Automation Hub"):
            continue
        purchase_datetime = rec.get("purchase_datetime")
        for raw in rec.get("tracking_numbers") or []:
            if not isinstance(raw, str):
                continue
            n = _normalize_number(raw)
            if not n:
                continue
            if n not in out and isinstance(purchase_datetime, str) and purchase_datetime.strip():
                out[n] = purchase_datetime.strip()
    return out


def prefetch_tracking_for_records(
    records: list[dict],
    *,
    on_progress: Callable[[int, int], None] | None = None,
    cancel_check: Callable[[], bool] | None = None,
) -> None:
    """Call 17TRACK for each unique ID before Excel: refresh non-terminal; throttle terminal cache.

    *on_progress* is invoked as ``on_progress(done, total)`` after each ID is handled
    (``done`` is 1-based index). *cancel_check* returns True to stop prefetch early.
    """
    key = api.api_key_from_env()
    if not key or not records:
        return
    purchase_dates = _tracking_purchase_dates(records)
    ids = iter_unique_tracking_ids(records)
    total = len(ids)
    if total <= 0:
        return

    if on_progress:
        on_progress(0, total)
    for idx, n in enumerate(ids, start=1):
        if cancel_check and cancel_check():
            break
        try:
            cached = load_cache(n)
            force = True
            if cached:
                gr = cached.get("last_get_response")
                ti = extract_track_info(gr, n) if isinstance(gr, dict) else None
                out = classify_outcome(ti)
                last_ts = _parse_iso(cached.get("last_fetch_iso"))
                if out == "dead_invalid":
                    cutoff_ts = _parse_iso(cached.get("notfound_cutoff_iso"))
                    if cutoff_ts is not None and time.time() > cutoff_ts and bool(cached.get("notfound_give_up")):
                        force = False
                    else:
                        force = True
                elif (
                    out == "terminal"
                    and last_ts is not None
                    and (time.time() - last_ts) < MIN_FETCH_INTERVAL_SEC
                ):
                    force = False
            fetch_tracking_smart(
                key,
                n,
                force_refresh=force,
                purchase_datetime=purchase_dates.get(n),
            )
        except Exception:
            # Network / API errors on one ID must not skip prefetch for the rest (e.g. UPS vs FedEx).
            continue
        if on_progress:
            on_progress(idx, total)


def shipping_summary_metrics(track_numbers: list[str]) -> tuple[int, int]:
    """(valid_count, delivered_count) from disk cache (run :func:`prefetch_tracking_for_records` first)."""
    valid = 0
    delivered = 0
    for raw in track_numbers:
        n = _normalize_number(raw)
        if not n:
            continue
        c = load_cache(n)
        if not c:
            continue
        gr = c.get("last_get_response")
        ti = extract_track_info(gr, n) if isinstance(gr, dict) else None
        out = classify_outcome(ti)
        if out == "dead_invalid":
            continue
        if ti is None and out != "no_data":
            continue
        valid += 1
        if is_delivered(ti):
            delivered += 1
    return valid, delivered


def format_shipping_summary_line(valid: int, delivered: int) -> str:
    if valid == 0:
        return "No status data"
    if delivered == valid:
        return "All Delivered"
    if delivered == 0:
        return "None Delivered"
    pct = int(round(100.0 * delivered / valid))
    return f"{pct}% Delivered"
