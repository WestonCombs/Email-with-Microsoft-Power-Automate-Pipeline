"""Email href pipeline: extract from HTML → follow redirects → classify shipment tracking.

Anchor ``href`` values are parsed with BeautifulSoup (no regex-based URL/href scanning).
"""

from __future__ import annotations

import os
import threading
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from functools import partial
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse, urljoin

from bs4 import BeautifulSoup

from shared import runLogger as RL

def _dbg(msg: str) -> None:
    """Write one debug line to logs/debug_tracking_hrefs.txt — file only, no console."""
    try:
        if not RL.is_debug():
            return
        RL.debug("tracking_hrefs", f"  {msg}")
    except Exception:
        pass


def _link_debug_enabled() -> bool:
    return RL.is_debug() or os.getenv("EMAIL_LINK_DEBUG", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )


def _link_debug(msg: str) -> None:
    """Write link parsing diagnostics to file logs (never console)."""
    try:
        if not _link_debug_enabled():
            return
        if RL.is_debug():
            RL.debug("tracking_hrefs", f"  {msg}")
            return
        # Explicit EMAIL_LINK_DEBUG without DEBUG_MODE still records diagnostics in file logs.
        RL.log("tracking_hrefs", f"{RL.ts()}  [link-debug] {msg}")
    except Exception:
        pass


_SRC = "tracking_hrefs"

# --- JSON / Excel compatibility: same sentinel string as historical runs ---
MULTIPLE_TRACKING_LINKS = "multiple tracking links found"

_TRACKING_DOMAINS = [
    "cta.nam.com",
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

_CARRIER_TRACKING_DOMAINS = frozenset({
    "ups.com",
    "fedex.com",
    "usps.com",
    "dhl.com",
    "ontrac.com",
    "lasership.com",
})

_TRACKING_PATH_KEYWORDS = [
    "track",
    "tracking",
    "shipment",
    "orderstatus",
    "delivery",
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

# Sorting only (not used for include/exclude classification) — longer matches rank higher.
_LINK_PRIORITY_SUBSTRINGS = (
    "track",
    "tracking",
    "shipment",
    "delivery",
    "orderstatus",
    "order/",
    "order?",
    "order&",
    "order=",
)


def _host_matches_domain(host: str, domain: str) -> bool:
    host = (host or "").strip(".").lower()
    domain = (domain or "").strip(".").lower()
    return host == domain or host.endswith("." + domain)


def _is_known_carrier_tracking_url(url: str) -> bool:
    """True for carrier tracking pages that are already useful browser targets."""
    s = clean_link(url)
    if s.startswith("//"):
        s = "https:" + s
    try:
        parsed = urlparse(s)
    except Exception:
        return False
    host = parsed.hostname or ""
    if not any(_host_matches_domain(host, d) for d in _CARRIER_TRACKING_DOMAINS):
        return False
    low = s.lower()
    return any(kw in low for kw in _TRACKING_PATH_KEYWORDS)


def clean_link(link: str) -> str:
    """Collapse broken/multiline email ``href`` values into one line."""
    s = (link or "").replace("\n", "").replace("\r", "").strip()
    # Strip BOM / zero-width chars (e.g. first href in a UTF-8 file with BOM).
    return s.lstrip("\ufeff\u200b\u200c\u200d\u2060").strip()


def is_absolute_browser_url(url: str) -> bool:
    """True if *url* is ``http(s)://`` with a host that can be opened in a browser.

    Drops stray tokens (e.g. query fragments saved without a scheme) that would
    never work with ``os.startfile`` / the default browser.
    """
    s = clean_link(url)
    if not s:
        return False
    if s.startswith("//"):
        s = "https:" + s
    try:
        p = urlparse(s)
    except Exception:
        return False
    if p.scheme not in ("http", "https"):
        return False
    host = (p.netloc or "").split("@")[-1].split(":")[0].strip().strip(".").lower()
    if not host:
        return False
    if host == "localhost":
        return True
    if host.startswith("["):
        return True
    if "." in host:
        return True
    # IPv4 without domain name
    octets = host.split(".")
    if len(octets) == 4 and all(o.isdigit() and 0 <= int(o) <= 255 for o in octets):
        return True
    return False


def _link_priority_score(url: str) -> int:
    low = url.lower()
    return sum(1 for s in _LINK_PRIORITY_SUBSTRINGS if s in low)


def extract_all_links(html: str) -> list[str]:
    """Return every distinct anchor ``href`` from *html* (document order), fully cleaned.

    Uses ``BeautifulSoup`` and ``<a href=…>`` only — no regex href scanning.
    """
    if not html:
        return []

    soup = BeautifulSoup(html, "html.parser")
    tags = soup.find_all("a", href=True)

    seen: set[str] = set()
    out: list[str] = []

    for tag in tags:
        raw = tag.get("href")
        if raw is None:
            continue
        link = clean_link(str(raw))
        if not link:
            continue

        low = link.lower()
        if low.startswith("http://") or low.startswith("https://"):
            if len(link) <= 50 and _link_priority_score(link) > 0:
                _link_debug(
                    f"SUSPICIOUS short tracking-like http(s) link "
                    f"(len={len(link)}): {link}"
                )
        elif low.startswith("http") and not (
            low.startswith("http://") or low.startswith("https://")
        ):
            _link_debug(f"SUSPICIOUS non-standard http scheme prefix: {link[:120]}")

        if link not in seen:
            seen.add(link)
            out.append(link)
            if _link_debug_enabled():
                _link_debug(f"Link Length: {len(link)}")
                _link_debug(f"Link Preview: {link[:120]}")

    return out


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
    """Return unique ``href`` values from raw HTML (BeautifulSoup ``<a href>`` only)."""
    raw_occurrences = html.lower().count("href=")
    RL.trace(
        _SRC,
        f"extract_hrefs_from_html() — html len={len(html):,}, raw 'href=' count={raw_occurrences}",
    )
    out = extract_all_links(html)
    RL.trace(_SRC, f"extract_hrefs_from_html() — {len(out)} unique hrefs")
    max_len = max((len(x) for x in out), default=0)
    if RL.is_debug():
        RL.debug(
            "tracking_hrefs",
            f"  extract_hrefs_from_html summary: unique={len(out)}, "
            f"max_href_len={max_len}, raw_href_token_count={raw_occurrences}",
        )
        if raw_occurrences > 0 and len(out) == 0:
            RL.debug(
                "tracking_hrefs",
                "  WARNING: HTML contains 'href=' substrings but 0 <a href> links were "
                "parsed (saved file may differ from inbox HTML, or hrefs are not on <a>).",
            )
        for i, h in enumerate(out[:25], 1):
            RL.debug(
                "tracking_hrefs",
                f"  href[{i}] len={len(h)} preview={h[:200]!r}",
            )
        if len(out) > 25:
            RL.debug(
                "tracking_hrefs",
                f"  … {len(out) - 25} more href(s) omitted from debug log (see JSON extracted_links)",
            )
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
    if _is_known_carrier_tracking_url(current):
        _dbg(f"resolve stop (already carrier tracking)  {current[:120]}")
        return current

    _dbg(f"resolve start    {current[:120]}")
    opener = _redirect_opener()
    for hop in range(1, max_hops + 1):
        code, loc = _http_status_and_location(current, "HEAD", opener, timeout)
        _dbg(f"  hop {hop}  HEAD {code}  loc={loc[:80] if loc else 'none'}")
        if code in _REDIRECT_CODES and loc:
            nxt = urljoin(current, loc.strip())
            _dbg(f"  hop {hop}  -> redirect {code} to {nxt[:120]}")
            current = nxt
            if _is_known_carrier_tracking_url(current):
                _dbg(f"  hop {hop}  carrier tracking target reached -> done: {current[:120]}")
                return current
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
            if _is_known_carrier_tracking_url(current):
                _dbg(f"  hop {hop}  carrier tracking target reached -> done: {current[:120]}")
                return current
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
    """Distinct tracking URLs from ``(href, final)`` pairs (UTM-normalized dedupe).

    Narvar-style CTAs often redirect to a generic retailer home URL that no longer
    matches tracking heuristics, while the *raw* ``href`` still does. If either
    raw or final classifies as tracking, we keep a link—preferring ``final`` when
    *both* classify, otherwise the side that still matches.
    """
    if not pairs:
        _dbg("list_tracking_links_from_pairs: empty pairs (no hrefs extracted or resolved)")
        return []

    _dbg(f"classify: {len(pairs)} (href, final) pair(s) in document order")
    seen_norm: set[str] = set()
    tracking_chosen: list[str] = []

    for href, final in pairs:
        raw_ok = url_classifies_as_tracking(href)
        fin_ok = url_classifies_as_tracking(final)
        if not raw_ok and not fin_ok:
            _dbg(f"  not-tracking  raw={href[:120]}  final={final[:120]}")
            continue
        chosen = final if fin_ok else href
        if raw_ok and not fin_ok:
            _dbg(
                "  TRACKING (raw only — final landing page not matched)  "
                f"final[:80]={final[:80]!r}"
            )
        else:
            _dbg(f"  {'TRACKING    ' if fin_ok else 'TRACKING(raw)'}  {chosen[:120]}")
        norm = _strip_utm_for_dedupe(chosen)
        if norm not in seen_norm:
            if not is_absolute_browser_url(chosen):
                _dbg(f"  skip (not a full http(s) URL): {chosen[:160]!r}")
                continue
            seen_norm.add(norm)
            tracking_chosen.append(chosen)

    _dbg(f"list_tracking_links: {len(tracking_chosen)} distinct tracking URL(s)")
    if len(tracking_chosen) > 1:
        indexed = list(enumerate(tracking_chosen))
        indexed.sort(key=lambda t: (-_link_priority_score(t[1]), t[0]))
        tracking_chosen = [t[1] for t in indexed]
    return tracking_chosen


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
