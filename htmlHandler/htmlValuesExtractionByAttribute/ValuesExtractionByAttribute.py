"""Extract all values for a given HTML attribute using plain string scanning.

No LLM or external API needed — pure local program logic.
Handles both double-quoted and single-quoted attribute values.
"""

import re

_PATTERN_CACHE: dict[str, re.Pattern] = {}


def _get_pattern(attribute: str) -> re.Pattern:
    """Build and cache a regex for the given attribute name."""
    if attribute not in _PATTERN_CACHE:
        escaped = re.escape(attribute)
        _PATTERN_CACHE[attribute] = re.compile(
            rf'{escaped}\s*=\s*(?:"([^"]*?)"|\'([^\']*?)\')',
            re.IGNORECASE,
        )
    return _PATTERN_CACHE[attribute]


def extract_attribute_values(html: str, attribute: str) -> list[str]:
    """Return every unique value of *attribute* found in the raw HTML string.

    Example
    -------
    >>> extract_attribute_values('<a href="https://ups.com/track">Track</a>', "href")
    ['https://ups.com/track']
    """
    pattern = _get_pattern(attribute)
    seen: set[str] = set()
    values: list[str] = []
    for match in pattern.finditer(html):
        value = (match.group(1) if match.group(1) is not None else match.group(2)).strip()
        if value and value not in seen:
            seen.add(value)
            values.append(value)
    return values
