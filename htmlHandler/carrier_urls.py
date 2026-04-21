"""Infer carrier from a tracking ID and build public carrier tracking URLs (browser, not APIs)."""

from __future__ import annotations

import re
from urllib.parse import quote

# Import canonical form from sibling module
from htmlHandler.carrier_tracking_ids import _canonical_display, _norm_key


def infer_carrier(tracking_id: str) -> str:
    """Return a short label: UPS, USPS, FedEx, DHL, or Unknown."""
    p = _canonical_display((tracking_id or "").strip())
    if not p:
        return "Unknown"
    nk = _norm_key(p)

    if nk.startswith("1Z") and len(nk) == 18:
        return "UPS" if re.match(r"^1Z[0-9A-Z]{16}$", nk) else "Unknown"

    if nk.isdigit():
        ln = len(nk)
        if nk.startswith("9") and ln in (22, 30):
            return "USPS"
        if nk[:2] in ("92", "93", "94", "95") and 20 <= ln <= 24:
            return "USPS"
        if 12 <= ln <= 15:
            return "FedEx"
        if ln in (10, 11):
            return "DHL"

    return "Unknown"


def normalize_carrier_for_public_url(
    carrier_label: str | None,
    tracking_id: str = "",
) -> str | None:
    """
    Map 17TRACK / long carrier names to labels ``public_tracking_url`` understands
    (``UPS``, ``USPS``, ``FEDEX``, ``DHL``). Returns ``None`` to fall back to
    :func:`infer_carrier` for the tracking id.
    """
    if not carrier_label or not isinstance(carrier_label, str):
        if tracking_id:
            return infer_carrier(tracking_id)
        return None
    u = carrier_label.strip().upper()
    if not u:
        if tracking_id:
            return infer_carrier(tracking_id)
        return None
    if "FEDEX" in u or "FED EX" in u:
        return "FEDEX"
    if "UPS" in u or u.startswith("UNITED PARCEL"):
        return "UPS"
    if "USPS" in u or "U.S. POSTAL" in u or "UNITED STATES POSTAL" in u:
        return "USPS"
    if "DHL" in u:
        return "DHL"
    if tracking_id:
        return infer_carrier(tracking_id)
    return None


def public_tracking_url(tracking_id: str, carrier: str | None = None) -> str:
    """HTTPS URL to the carrier's (or fallback) public tracking page for this ID."""
    raw = _canonical_display((tracking_id or "").strip())
    if not raw:
        return ""
    c = (carrier or infer_carrier(raw)).upper()
    q = quote(raw, safe="")

    if c == "UPS":
        return f"https://www.ups.com/track?loc=en_US&tracknum={q}"
    if c == "USPS":
        return f"https://tools.usps.com/go/TrackConfirmAction?tLabels={q}"
    if c == "FEDEX":
        return f"https://www.fedex.com/fedextrack/?trknbr={q}"
    if c == "DHL":
        return f"https://www.dhl.com/en/express/tracking.html?AWB={q}"

    # Universal fallback (no carrier guess)
    return f"https://www.17track.net/en/track#nums={q}"
