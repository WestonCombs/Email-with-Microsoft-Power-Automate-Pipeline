"""Convert raw HTML email content to clean plain text for LLM consumption.

Uses Python's built-in html.parser (via BeautifulSoup) to strip tags,
remove hidden elements, and convert data tables to readable markdown.
"""

import re

from bs4 import BeautifulSoup


def _remove_hidden_elements(soup: BeautifulSoup) -> None:
    """Remove elements that are visually hidden in the rendered email."""
    for el in soup.find_all(style=True):
        style = el.get("style") or ""
        if re.search(r"display\s*:\s*none", style, re.IGNORECASE) or \
           re.search(r"visibility\s*:\s*hidden", style, re.IGNORECASE):
            el.decompose()


def _convert_tables_to_markdown(soup: BeautifulSoup) -> None:
    """Convert multi-row, multi-column data tables to markdown so the LLM
    sees row/column relationships (e.g. item-price pairings) that are lost
    when tags are flattened to plain text."""
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


def convert(html: str, max_chars: int = 50_000) -> str:
    """Convert raw HTML to clean plain text, trimmed to *max_chars*."""
    soup = BeautifulSoup(html, "html.parser")

    _remove_hidden_elements(soup)
    _convert_tables_to_markdown(soup)

    for tag in soup(["script", "style", "noscript", "svg", "meta", "head"]):
        tag.decompose()

    text = soup.get_text(separator="\n")
    lines = [line.strip() for line in text.splitlines()]
    lines = [line for line in lines if line]
    cleaned = "\n".join(lines)

    if len(cleaned) > max_chars:
        cleaned = cleaned[:max_chars]
    return cleaned
