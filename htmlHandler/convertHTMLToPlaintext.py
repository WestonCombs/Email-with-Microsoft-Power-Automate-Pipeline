"""Convert raw HTML email content to clean plain text for LLM consumption.

Uses Python's built-in html.parser (via BeautifulSoup) to strip tags,
remove hidden elements, and convert data tables to readable markdown.
"""

import os
import re

from bs4 import BeautifulSoup

from shared import runLogger as RL

_SRC = "convertHTMLToPlaintext"
_OPENAI_MAX_CHARS_ENV = "openai_max_chars_per_prompting"
_DEFAULT_MAX_CHARS = 50_000


def _max_chars_from_env() -> int:
    raw = os.getenv(_OPENAI_MAX_CHARS_ENV)
    if raw is None or not str(raw).strip():
        return _DEFAULT_MAX_CHARS
    try:
        n = int(str(raw).strip(), 10)
    except ValueError:
        return _DEFAULT_MAX_CHARS
    return n if n > 0 else _DEFAULT_MAX_CHARS


def _remove_hidden_elements(soup: BeautifulSoup) -> int:
    """Remove elements that are visually hidden in the rendered email.
    Returns the count of removed elements."""
    removed = 0
    # Two-pass: collect targets first, then remove.
    # Decomposing a parent also destroys its children, clearing their attrs dict
    # to None.  If a child also had a style attribute it would still be queued
    # in a single-pass loop, and the next el.get("style") call would crash with
    # "'NoneType' object has no attribute 'get'".
    to_remove = []
    for el in soup.find_all(style=True):
        style = el.get("style") or ""
        if re.search(r"display\s*:\s*none", style, re.IGNORECASE) or \
           re.search(r"visibility\s*:\s*hidden", style, re.IGNORECASE):
            to_remove.append(el)
    for el in to_remove:
        if el.parent is not None:   # skip if already removed by an ancestor
            el.decompose()
            removed += 1
    return removed


def _convert_tables_to_markdown(soup: BeautifulSoup) -> int:
    """Convert multi-row, multi-column data tables to markdown so the LLM
    sees row/column relationships (e.g. item-price pairings) that are lost
    when tags are flattened to plain text.  Returns the count of converted tables."""
    converted = 0
    for table in reversed(soup.find_all("table")):
        rows: list[list[str]] = []
        for tr in table.find_all("tr"):
            cells = tr.find_all(["td", "th"])
            if cells:
                rows.append([
                    re.sub(r"\s+", " ", c.get_text(separator=" ")).strip().replace("|", "\\|")
                    for c in cells
                ])

        if len(rows) < 2 or all(len(r) < 2 for r in rows):
            continue

        max_cols = max(len(r) for r in rows)
        for r in rows:
            r.extend([""] * (max_cols - len(r)))

        lines: list[str] = []
        for i, row in enumerate(rows):
            lines.append("| " + " | ".join(row) + " |")
            if i == 0:
                lines.append("| " + " | ".join(["---"] * max_cols) + " |")

        table.replace_with("\n" + "\n".join(lines) + "\n")
        converted += 1
    return converted


def convert(html: str, max_chars: int | None = None) -> str:
    """Convert raw HTML to clean plain text, trimmed to *max_chars*.

    If *max_chars* is None, uses integer from env *openai_max_chars_per_prompting*
    (see python_files/.env), or 50_000 if unset or invalid.
    """
    if max_chars is None:
        max_chars = _max_chars_from_env()
    RL.trace(_SRC, f"convert() called — raw HTML len={len(html):,} chars")

    soup = BeautifulSoup(html, "html.parser")

    hidden_count = _remove_hidden_elements(soup)
    RL.trace(_SRC, f"removed {hidden_count} hidden elements")

    table_count = _convert_tables_to_markdown(soup)
    RL.trace(_SRC, f"converted {table_count} data tables to markdown")

    for tag in soup(["script", "style", "noscript", "svg", "meta", "head"]):
        tag.decompose()

    text = soup.get_text(separator="\n")
    lines = [line.strip() for line in text.splitlines()]
    lines = [line for line in lines if line]
    cleaned = "\n".join(lines)

    was_truncated = len(cleaned) > max_chars
    if was_truncated:
        cleaned = cleaned[:max_chars]

    RL.trace(
        _SRC,
        f"convert() result — plain text len={len(cleaned):,} chars, "
        f"lines={len(cleaned.splitlines())}, truncated={was_truncated}",
    )
    return cleaned
