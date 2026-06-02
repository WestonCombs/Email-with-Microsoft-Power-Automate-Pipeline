from __future__ import annotations

import re


_SUPPLIER_DOMAIN_HINTS: dict[str, str] = {
    "typology.com": "Typology",
    "typology.us": "Typology",
}

_SUPPLIER_ALIAS_DISPLAY: dict[str, str] = {
    "typology": "Typology",
    "typologycom": "Typology",
    "typology.com": "Typology",
    "typology paris": "Typology",
    "typologyparis": "Typology",
    "typology us": "Typology",
    "typologyus": "Typology",
}


def _clean_text(value: object) -> str:
    return str(value or "").replace("\ufeff", "").strip()


def _extract_email_domain(value: object) -> str:
    cleaned = _clean_text(value)
    if "@" not in cleaned:
        return ""
    return cleaned.rsplit("@", 1)[-1].strip().strip(">").casefold()


def _supplier_alias_key(value: object) -> str:
    cleaned = _clean_text(value)
    if not cleaned:
        return ""
    cleaned = re.sub(r"^mailto:", "", cleaned, flags=re.IGNORECASE).strip()
    email_match = re.search(r"<([^<>@\s]+@[^<>\s]+)>", cleaned)
    if email_match:
        cleaned = email_match.group(1)
    cleaned = re.sub(r"^https?://", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.split("/", 1)[0].split("?", 1)[0].strip().strip("<>")
    cleaned = cleaned.removeprefix("www.")
    folded = cleaned.casefold().replace("&", " and ")
    folded = re.sub(r"[^a-z0-9@.]+", " ", folded)
    folded = re.sub(r"\s+", " ", folded).strip()
    return folded if folded in _SUPPLIER_ALIAS_DISPLAY else re.sub(r"[^a-z0-9@]+", "", folded)


def normalize_supplier_display_name(value: object, sender_email: object = None) -> str | None:
    """Return a canonical supplier name for shared aliases, or ``None``."""
    cleaned = _clean_text(value)
    if not cleaned:
        return None

    domain = _extract_email_domain(cleaned) or cleaned.casefold()
    domain = re.sub(r"^https?://", "", domain, flags=re.IGNORECASE)
    domain = domain.split("/", 1)[0].split("?", 1)[0].strip().strip(">")
    domain = domain.removeprefix("www.")
    for suffix, display in _SUPPLIER_DOMAIN_HINTS.items():
        if domain == suffix or domain.endswith(f".{suffix}"):
            return display

    sender_domain = _extract_email_domain(sender_email)
    if sender_domain:
        for suffix, display in _SUPPLIER_DOMAIN_HINTS.items():
            if sender_domain == suffix or sender_domain.endswith(f".{suffix}"):
                if "@" in cleaned and _extract_email_domain(cleaned) == sender_domain:
                    return display

    key = _supplier_alias_key(cleaned)
    return _SUPPLIER_ALIAS_DISPLAY.get(key)


def normalized_supplier_vote_key(value: object) -> str:
    normalized = normalize_supplier_display_name(value)
    if normalized:
        return normalized.casefold()
    return ""
