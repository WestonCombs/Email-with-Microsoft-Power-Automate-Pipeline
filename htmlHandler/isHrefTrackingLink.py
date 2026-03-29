"""Determine the tracking link from a list of href values using local logic.

No LLM or external API needed — keyword and domain matching only.
"""

from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

from htmlHandler.admin_log import trace

_SRC = "isHrefTrackingLink"

TRACKING_DOMAINS = [
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

TRACKING_PATH_KEYWORDS = [
    "track",
    "tracking",
    "shipment",
    "orderstatus",
]

EXCLUDE_KEYWORDS = [
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


def _normalize_url(url: str) -> str:
    """Strip UTM / analytics query params so near-duplicates collapse."""
    try:
        parsed = urlparse(url)
        params = parse_qs(parsed.query, keep_blank_values=False)
        filtered = {k: v for k, v in params.items() if k.lower() not in _UTM_PARAMS}
        clean_query = urlencode(filtered, doseq=True)
        return urlunparse(parsed._replace(query=clean_query, fragment=""))
    except Exception:
        return url


def _is_tracking_link(href: str) -> bool:
    """Return True if *href* looks like a shipment tracking URL."""
    lower = href.lower()

    for kw in EXCLUDE_KEYWORDS:
        if kw in lower:
            return False

    try:
        host = urlparse(href).hostname or ""
    except Exception:
        host = ""

    for domain in TRACKING_DOMAINS:
        if host.endswith(domain):
            return True

    for kw in TRACKING_PATH_KEYWORDS:
        if kw in lower:
            return True

    return False


def determine_tracking_link(hrefs: list[str]) -> str | None:
    """Evaluate a list of href strings and return:

    - A single URL string  — exactly one unique tracking link found
    - ``"multiple tracking links found"`` — more than one unique tracking link
    - ``None`` — no tracking links found
    """
    # Link-value logging disabled — tracking-link bug is resolved
    # preview = ", ".join(hrefs[:2]) if hrefs else "(empty list)"
    # trace(_SRC, f"determine_tracking_link() called — {len(hrefs)} hrefs, first 2: [{preview}]")

    seen_normalized: set[str] = set()
    unique_tracking: list[str] = []

    for href in hrefs:
        is_match = _is_tracking_link(href)
        # if is_match:
        #     trace(_SRC, f"  TRACKING match: {href}")
        if not is_match:
            continue
        normalized = _normalize_url(href)
        if normalized not in seen_normalized:
            seen_normalized.add(normalized)
            unique_tracking.append(href)

    if len(unique_tracking) == 0:
        result = None
    elif len(unique_tracking) == 1:
        result = unique_tracking[0]
    else:
        result = "multiple tracking links found"

    # trace(_SRC, f"determine_tracking_link() result — {result!r}")
    return result
