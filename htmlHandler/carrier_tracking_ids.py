"""Extract carrier-style tracking IDs from email text and from tracking URLs (incl. redirect finals).

IDs are normalized for deduplication; output order preserves first-seen across sources.
"""

from __future__ import annotations

import re
from urllib.parse import parse_qs, unquote, urlparse

# Query keys commonly used by retailers, Narvar, and carrier sites for package IDs.
_TRACK_QUERY_KEY_SUBSTRINGS = (
    "track",
    "trkn",
    "shipment",
    "package",
    "inquiry",
    "label",
    "tn",
    "waybill",
    "awb",
)

# Exact-ish keys (lowercase) often holding a single tracking token.
_TRACK_QUERY_KEYS_EXACT = frozenset({
    "n",
    "t",
    "id",
    "ids",
    "q",
    "tn",
    "trk",
    "trkn",
    "trknbr",
    "trackno",
    "tracknum",
    "tracknums",
    "tracknumber",
    "tracknumbers",
    "trackingnumber",
    "trackingnumbers",
    "tracking_number",
    "tracking_numbers",
    "tracking",
    "trackingid",
    "tracking_id",
    "shipid",
    "shipmentid",
    "packageid",
})

# Regex patterns applied to visible text (and to full URL strings). Order matters for priority.
_TEXT_PATTERNS: tuple[tuple[re.Pattern[str], str], ...] = (
    # UPS standard
    (re.compile(r"\b1Z[0-9A-Z]{16}\b", re.IGNORECASE), "ups"),
    # USPS: common 22- and 30-digit domestic formats (leading 9)
    (re.compile(r"\b9[0-9]{21}\b"), "usps22"),
    (re.compile(r"\b9[0-9]{29}\b"), "usps30"),
    # USPS / IMpb variants (20–24 digits starting with 92/93/94/95)
    (re.compile(r"\b(?:92|93|94|95)[0-9]{18,22}\b"), "usps20_24"),
    # FedEx: 12-digit (numeric) — conservative boundary
    (re.compile(r"(?<![0-9])([0-9]{12})(?![0-9])"), "fedex12"),
    # FedEx: 14- and 15-digit numeric (some services)
    (re.compile(r"(?<![0-9])([0-9]{14}|[0-9]{15})(?![0-9])"), "fedex1415"),
    # FedEx / display: groups of four digits (optional spaces)
    (re.compile(r"\b(?:[0-9]{4}\s){2}[0-9]{4}\s[0-9]{4}\b"), "fedex_spaced"),
    # DHL (numeric, 10–11 digits)
    (re.compile(r"(?<![0-9])([0-9]{10}|[0-9]{11})(?![0-9])"), "dhl_num"),
)


def _norm_key(s: str) -> str:
    return re.sub(r"[^0-9A-Z]", "", (s or "").upper())


def _canonical_display(s: str) -> str:
    t = (s or "").strip()
    if not t:
        return ""
    if re.match(r"^1Z[0-9A-Z]{16}$", t, re.I):
        return t.upper()
    # collapse internal whitespace for digit runs
    if re.fullmatch(r"[\d\s]+", t):
        return re.sub(r"\s+", "", t)
    return t


def _is_weak_numeric_only(s: str, kind: str) -> bool:
    """Reduce false positives for bare digit runs (order numbers, timestamps)."""
    if not s.isdigit():
        return False
    if len(set(s)) <= 1:
        return True
    if kind in ("fedex12", "dhl_num") and len(s) == 12:
        # 12-digit could be order id; still allow if we keep for URL-origin only — caller filters
        return False
    return False


def _valid_carrier_token(raw: str, *, source: str) -> bool:
    """Final plausibility check after regex match."""
    s = _canonical_display(raw)
    if not s:
        return False
    nk = _norm_key(s)
    if len(nk) < 8:
        return False
    if len(nk) > 34:
        return False
    # UPS
    if nk.startswith("1Z"):
        return bool(re.match(r"^1Z[0-9A-Z]{16}$", nk))
    # USPS-style long numerics
    if nk.isdigit():
        ln = len(nk)
        if nk.startswith("9") and ln in (22, 30):
            return True
        if nk[:2] in ("92", "93", "94", "95") and 20 <= ln <= 24:
            return True
        # FedEx 12–15
        if 12 <= ln <= 15 and source in ("fedex12", "fedex1415", "url", "fedex_spaced"):
            if _is_weak_numeric_only(nk, "fedex12"):
                return False
            return True
        # DHL 10–11 — only from URL or with strong context; allow text with hesitation
        if ln in (10, 11) and source == "url":
            return True
        if ln in (10, 11) and source == "dhl_num":
            return False
        return False
    return False


def extract_carrier_ids_from_text(text: str) -> list[str]:
    """Scan visible email text for carrier-style IDs (order of first appearance)."""
    if not (text or "").strip():
        return []
    out: list[str] = []
    seen: set[str] = set()
    for rx, kind in _TEXT_PATTERNS:
        for m in rx.finditer(text):
            raw = m.group(0) if m.lastindex is None else m.group(1)
            disp = _canonical_display(raw)
            if not _valid_carrier_token(disp, source=kind):
                continue
            key = _norm_key(disp)
            if key in seen:
                continue
            seen.add(key)
            out.append(disp)
    return out


def _query_values_might_be_tracking(key: str) -> bool:
    lk = key.lower()
    if lk in _TRACK_QUERY_KEYS_EXACT:
        return True
    for sub in _TRACK_QUERY_KEY_SUBSTRINGS:
        if sub in lk:
            return True
    return False


def _tokens_from_query(url: str) -> list[str]:
    out: list[str] = []
    try:
        q = urlparse(url).query
        if not q:
            return out
        for k, vals in parse_qs(q, keep_blank_values=False).items():
            if not _query_values_might_be_tracking(k):
                continue
            for v in vals:
                v = unquote((v or "").strip())
                if not v or len(v) > 80:
                    continue
                # Some sites pack multiple IDs comma-separated
                for piece in re.split(r"[\s,;]+", v):
                    piece = piece.strip(" '\"")
                    if len(piece) >= 8:
                        out.append(piece)
    except Exception:
        pass
    return out


def _tokens_from_path_and_string(url: str) -> list[str]:
    """Run the same text regexes against the raw URL (path + query often embed 1Z…)."""
    return extract_carrier_ids_from_text(url)


def extract_carrier_ids_from_url(url: str) -> list[str]:
    """Pull tracking tokens from a single URL (query params + embedded patterns in the string)."""
    if not (url or "").strip():
        return []
    merged: list[str] = []
    seen: set[str] = set()

    def add_many(candidates: list[str]) -> None:
        for c in candidates:
            disp = _canonical_display(c)
            if not disp:
                continue
            # URL-origin: allow plausible 12-digit FedEx
            ok = _valid_carrier_token(disp, source="url")
            if not ok and disp.isdigit() and len(disp) == 12:
                ok = True
            if not ok:
                continue
            key = _norm_key(disp)
            if key in seen:
                continue
            seen.add(key)
            merged.append(disp)

    add_many(_tokens_from_query(url))
    add_many(_tokens_from_path_and_string(url))
    return merged


def extract_carrier_ids_from_href_pairs(pairs: list[tuple[str, str]]) -> list[str]:
    """Collect IDs from each raw href and final redirect URL (document order, unique)."""
    out: list[str] = []
    seen: set[str] = set()
    for href, final in pairs:
        for u in (href, final):
            for tid in extract_carrier_ids_from_url(u):
                key = _norm_key(tid)
                if key in seen:
                    continue
                seen.add(key)
                out.append(tid)
    return out


def extract_carrier_ids_from_tracking_link_pairs(pairs: list[tuple[str, str]]) -> list[str]:
    """Collect IDs only from href/final pairs that qualify as shipment-tracking URLs.

    Uses the same classification rules as :func:`htmlHandler.tracking_hrefs.list_tracking_links_from_pairs`
    (either side classifies as tracking; chosen URL must be a usable browser URL). This is the set
    used to cross-check “regex/LLM” IDs against link-based discovery.
    """
    # Local import avoids circular import at module load.
    from htmlHandler.tracking_hrefs import is_absolute_browser_url, url_classifies_as_tracking

    out: list[str] = []
    seen: set[str] = set()
    for href, final in pairs:
        raw_ok = url_classifies_as_tracking(href)
        fin_ok = url_classifies_as_tracking(final)
        if not raw_ok and not fin_ok:
            continue
        chosen = final if fin_ok else href
        if not is_absolute_browser_url(chosen):
            continue
        for u in (href, final):
            for tid in extract_carrier_ids_from_url(u):
                key = _norm_key(tid)
                if key in seen:
                    continue
                seen.add(key)
                out.append(tid)
    return out


def merge_unique_tracking_ids(*sequences: list[str]) -> list[str]:
    """Concatenate sequences in order; keep first occurrence of each normalized ID."""
    out: list[str] = []
    seen: set[str] = set()
    for seq in sequences:
        for raw in seq:
            s = (raw or "").strip()
            if not s:
                continue
            disp = _canonical_display(s)
            key = _norm_key(disp)
            if not key or key in seen:
                continue
            seen.add(key)
            out.append(disp)
    return out


def normalize_openai_tracking_numbers(raw: object) -> list[str]:
    """Coerce OpenAI output to a list of non-empty strings (no dedupe)."""
    if raw is None:
        return []
    if isinstance(raw, str):
        s = raw.strip()
        return [s] if s else []
    if isinstance(raw, list):
        out: list[str] = []
        for x in raw:
            if isinstance(x, str) and x.strip():
                out.append(x.strip())
        return out
    return []
