"""Email href pipeline: extract from HTML → follow redirects → classify shipment tracking.

All logic lives here (regex scan, urllib redirects, domain/keyword heuristics).
"""

from __future__ import annotations

import os
import re
import threading
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from functools import partial
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse, urljoin

import runLogger as RL

_DEBUG_MODE: bool = os.getenv("DEBUG_MODE", "0").strip().lower() in ("1", "true", "yes")


def _dbg(msg: str) -> None:
    """Write one debug line to logs/debug_tracking_hrefs.txt — file only, no console."""
    if not _DEBUG_MODE:
        return
    try:
        RL.debug("tracking_hrefs", f"  {msg}")
    except Exception:
        pass

_SRC = "tracking_hrefs"

# --- JSON / Excel compatibility: same sentinel string as historical runs ---
MULTIPLE_TRACKING_LINKS = "multiple tracking links found"

_TRACKING_DOMAINS = [
    "narvar.com",
    "aftership.com",
    "trackingmore.com",
    "parcellab.com",
    "ups.com",
    "fedex.com",
    "usps.com",
    "dhl.com",
    "ontrac.com",
    "lasership.com",
]

_TRACKING_PATH_KEYWORDS = [
    "track",
    "tracking",
    "shipment",
    "orderstatus",
]

_EXCLUDE_KEYWORDS = [
    "login",
    "signin",
    "sign-in",
    "sign_in",
    "account",
    "password",
    "support",
    "help",
    "faq",
    "unsubscribe",
    "opt-out",
    "optout",
    "preferences",
    "manage-email",
    "marketing",
    "promo",
    "campaign",
    "newsletter",
]

_UTM_PARAMS = frozenset({
    "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content",
})

_HREF_PATTERN = re.compile(
    r'href\s*=\s*(?:"([^"]*?)"|\'([^\']*?)\')',
    re.IGNORECASE,
)

_DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

_REDIRECT_CODES = frozenset({301, 302, 303, 307, 308})

# Obvious non-HTTP schemes — never send to redirect resolver
_NON_WEB_PREFIXES = (
    "mailto:",
    "tel:",
    "sms:",
    "javascript:",
    "data:",
    "blob:",
    "about:",
)

# Process-wide: same href string reuses one redirect resolution across all emails in a run.
_final_url_cache: dict[tuple[str, int, float], str] = {}
_cache_lock = threading.Lock()


def _href_resolve_max_workers() -> int:
    raw = os.getenv("HREF_RESOLVE_MAX_WORKERS", "12").strip()
    try:
        n = int(raw)
    except ValueError:
        return 12
    return max(1, min(n, 64))


def normalize_href_for_http_fetch(href: str) -> tuple[str, bool]:
    """Return ``(url_for_requests, should_fetch)``.

    Only ``http``/``https`` (after normalization) are fetched. Protocol-relative
    ``//host/path`` becomes ``https://host/path``. Everything else (relative paths,
    ``mailto:``, ``javascript:``, fragments-only, unknown schemes) is not fetched.
    """
    raw = (href or "").strip()
    if not raw:
        return raw, False
    low = raw.lower()
    for p in _NON_WEB_PREFIXES:
        if low.startswith(p):
            return raw, False
    if raw.startswith("#"):
        return raw, False
    if raw.startswith("//"):
        return "https:" + raw, True
    parsed = urlparse(raw)
    sch = parsed.scheme.lower()
    if sch in ("http", "https"):
        return raw, True
    return raw, False


def summarize_hrefs_for_log(hrefs: list[str]) -> str:
    """One-line stats for pipeline logging."""
    n = len(hrefs)
    fetchable = sum(1 for h in hrefs if normalize_href_for_http_fetch(h)[1])
    return f"href_count={n}, http_fetchable={fetchable}, non_web_skipped={n - fetchable}"


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def _redirect_opener() -> urllib.request.OpenerDirector:
    return urllib.request.build_opener(_NoRedirect())


def _http_status_and_location(
    url: str,
    method: str,
    opener: urllib.request.OpenerDirector,
    timeout: float,
) -> tuple[int, str | None]:
    req = urllib.request.Request(
        url,
        method=method,
        headers={"User-Agent": _DEFAULT_USER_AGENT},
    )
    try:
        resp = opener.open(req, timeout=timeout)
        try:
            return resp.getcode(), resp.headers.get("Location")
        finally:
            resp.close()
    except urllib.error.HTTPError as e:
        loc = e.headers.get("Location") if e.headers else None
        return e.code, loc
    except Exception:
        return -1, None


def extract_hrefs_from_html(html: str) -> list[str]:
    """Return unique ``href`` values from raw HTML (double- or single-quoted)."""
    raw_occurrences = html.lower().count("href=")
    RL.trace(
        _SRC,
        f"extract_hrefs_from_html() — html len={len(html):,}, raw 'href=' count={raw_occurrences}",
    )
    seen: set[str] = set()
    out: list[str] = []
    for match in _HREF_PATTERN.finditer(html):
        value = (match.group(1) if match.group(1) is not None else match.group(2)).strip()
        if value and value not in seen:
            seen.add(value)
            out.append(value)
    RL.trace(_SRC, f"extract_hrefs_from_html() — {len(out)} unique hrefs")
    return out


def resolve_final_url(
    url: str,
    *,
    max_hops: int = 15,
    timeout: float = 12.0,
) -> str:
    """Follow HTTP redirects (multi-hop) and return the last URL in the chain.

    Only ``http``/``https`` targets are fetched (see :func:`normalize_href_for_http_fetch`).
    Other hrefs are returned unchanged. On failure, returns the last URL attempted.
    """
    raw = (url or "").strip()
    if not raw:
        return raw
    current, ok = normalize_href_for_http_fetch(raw)
    if not ok:
        _dbg(f"skip (non-http)  {raw[:120]}")
        return raw

    _dbg(f"resolve start    {current[:120]}")
    opener = _redirect_opener()
    for hop in range(1, max_hops + 1):
        code, loc = _http_status_and_location(current, "HEAD", opener, timeout)
        _dbg(f"  hop {hop}  HEAD {code}  loc={loc[:80] if loc else 'none'}")
        if code in _REDIRECT_CODES and loc:
            nxt = urljoin(current, loc.strip())
            _dbg(f"  hop {hop}  -> redirect {code} to {nxt[:120]}")
            current = nxt
            continue
        if code == 200:
            _dbg(f"  hop {hop}  HEAD 200 -> done: {current[:120]}")
            return current

        # HEAD failed or returned unexpected code — retry with GET
        code, loc = _http_status_and_location(current, "GET", opener, timeout)
        _dbg(f"  hop {hop}  GET  {code}  loc={loc[:80] if loc else 'none'}")
        if code in _REDIRECT_CODES and loc:
            nxt = urljoin(current, loc.strip())
            _dbg(f"  hop {hop}  -> redirect {code} to {nxt[:120]}")
            current = nxt
            continue
        if code == 200:
            _dbg(f"  hop {hop}  GET  200 -> done: {current[:120]}")
            return current
        _dbg(f"  hop {hop}  {code} (no further redirect) -> stopping at: {current[:120]}")
        return current

    _dbg(f"  max_hops={max_hops} reached -> final: {current[:120]}")
    return current


def resolve_final_url_cached(
    url: str,
    *,
    max_hops: int = 15,
    timeout: float = 12.0,
) -> str:
    """Like :func:`resolve_final_url`, but memoized per process for identical inputs.

    Reused across emails so repeated marketing/tracking URLs are not resolved twice.
    Thread-safe for parallel resolvers.
    """
    key = ((url or "").strip(), max_hops, timeout)
    with _cache_lock:
        hit = _final_url_cache.get(key)
    if hit is not None:
        return hit
    resolved = resolve_final_url(url, max_hops=max_hops, timeout=timeout)
    with _cache_lock:
        _final_url_cache[key] = resolved
    return resolved


def _href_to_final_pair(
    href: str,
    *,
    max_hops: int,
    timeout: float,
) -> tuple[str, str]:
    return href, resolve_final_url_cached(href, max_hops=max_hops, timeout=timeout)


def unique_final_urls(
    urls: list[str],
    *,
    max_workers: int | None = None,
    max_hops: int = 15,
    timeout: float = 12.0,
) -> list[str]:
    """Resolve each URL to its final destination; drop duplicate finals (first-seen order)."""
    if not urls:
        return []
    workers = max_workers if max_workers is not None else _href_resolve_max_workers()
    workers = min(workers, len(urls))
    pair_fn = partial(_href_to_final_pair, max_hops=max_hops, timeout=timeout)
    with ThreadPoolExecutor(max_workers=workers) as ex:
        finals = [f for _, f in ex.map(pair_fn, urls)]
    seen: set[str] = set()
    out: list[str] = []
    for final in finals:
        if final not in seen:
            seen.add(final)
            out.append(final)
    return out


def href_final_pairs(
    hrefs: list[str],
    *,
    max_workers: int | None = None,
    max_hops: int = 15,
    timeout: float = 12.0,
) -> list[tuple[str, str]]:
    """For each raw ``href``, the URL after the full redirect chain ``(href, final_url)``.

    Resolves distinct href strings in parallel (I/O bound). Results stay in the same order
    as *hrefs*. Repeated URLs across the whole program reuse :func:`resolve_final_url_cached`.
    """
    if not hrefs:
        return []
    workers = max_workers if max_workers is not None else _href_resolve_max_workers()
    workers = min(workers, len(hrefs))
    pair_fn = partial(_href_to_final_pair, max_hops=max_hops, timeout=timeout)
    with ThreadPoolExecutor(max_workers=workers) as ex:
        return list(ex.map(pair_fn, hrefs))


def list_tracking_links_from_pairs(
    pairs: list[tuple[str, str]],
) -> list[str]:
    """Distinct tracking destinations from pre-resolved ``(href, final)`` pairs (UTM-normalized dedupe)."""
    if not pairs:
        return []

    seen_final: set[str] = set()
    finals_unique: list[str] = []
    for _, final in pairs:
        if final not in seen_final:
            seen_final.add(final)
            finals_unique.append(final)

    _dbg(f"classify: {len(finals_unique)} unique final URLs to check")
    seen_norm: set[str] = set()
    tracking_finals: list[str] = []

    for final in finals_unique:
        verdict = url_classifies_as_tracking(final)
        _dbg(f"  {'TRACKING    ' if verdict else 'not-tracking'}  {final[:120]}")
        if not verdict:
            continue
        norm = _strip_utm_for_dedupe(final)
        if norm not in seen_norm:
            seen_norm.add(norm)
            tracking_finals.append(final)

    _dbg(f"list_tracking_links: {len(tracking_finals)} distinct tracking URL(s)")
    return tracking_finals


def pick_tracking_link_from_pairs(
    pairs: list[tuple[str, str]],
) -> str | None:
    """Same as :func:`pick_tracking_link`, but uses pre-resolved ``(href, final)`` pairs (one HTTP pass)."""
    tracking_finals = list_tracking_links_from_pairs(pairs)
    if len(tracking_finals) == 0:
        result = None
    elif len(tracking_finals) == 1:
        result = tracking_finals[0]
    else:
        result = MULTIPLE_TRACKING_LINKS
    _dbg(f"pick result: {result!r}")
    return result


def _strip_utm_for_dedupe(url: str) -> str:
    try:
        parsed = urlparse(url)
        params = parse_qs(parsed.query, keep_blank_values=False)
        filtered = {k: v for k, v in params.items() if k.lower() not in _UTM_PARAMS}
        clean_query = urlencode(filtered, doseq=True)
        return urlunparse(parsed._replace(query=clean_query, fragment=""))
    except Exception:
        return url


def url_classifies_as_tracking(url: str) -> bool:
    """Heuristic: does this URL string look like a shipment-tracking link (no HTTP calls)."""
    s = (url or "").strip()
    if s.startswith("//"):
        s = "https:" + s
    lower = s.lower()

    for kw in _EXCLUDE_KEYWORDS:
        if kw in lower:
            return False

    try:
        host = urlparse(s).hostname or ""
    except Exception:
        host = ""

    for domain in _TRACKING_DOMAINS:
        if host.endswith(domain):
            return True

    for kw in _TRACKING_PATH_KEYWORDS:
        if kw in lower:
            return True

    return False


def summarize_href_pairs(pairs: list[tuple[str, str]]) -> str:
    """Stats after resolution (for logs)."""
    n = len(pairs)
    diff = sum(1 for h, f in pairs if (h or "").strip() != (f or "").strip())
    return f"href_pairs={n}, final_differs_from_href={diff}"


def pick_tracking_link(hrefs: list[str]) -> str | None:
    """From raw hrefs: resolve redirect chains → unique final URLs → apply heuristics.

    Returns:
        - ``None`` — no tracking URL identified
        - A single ``https?://`` string — one distinct tracking destination
        - :data:`MULTIPLE_TRACKING_LINKS` — more than one distinct tracking destination
    """
    if not hrefs:
        return None
    return pick_tracking_link_from_pairs(href_final_pairs(hrefs))
