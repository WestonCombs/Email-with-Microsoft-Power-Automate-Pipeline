import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path

from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openai import OpenAI, RateLimitError

# Load .env from python_files/ (one level up from this script's subfolder)
load_dotenv(Path(__file__).resolve().parent.parent / ".env")

# =========================
# CONFIG
# =========================
OPENAI_API_KEY_ENV = "OPENAI_API_KEY"
API_KEY = os.getenv(OPENAI_API_KEY_ENV)

MODEL = "gpt-4o-mini"
LOCAL_MODEL = "llama3"

VALID_CATEGORIES = [
    "Order Placed",
    "Order Received",
    "Order Confirmed",
    "Order Delayed",
    "Order Shipped",
    "Delivery Confirmation",
    "Purchased Gift Card",
    "Unknown",
]
CATEGORY_CONFIDENCE_THRESHOLD = 60

RATE_LIMIT_RETRY_WAIT = 3
RATE_LIMIT_MAX_RETRIES = 20
RATE_LIMIT_THROTTLE_THRESHOLD = 0.60   # trigger proactive cooldown at 60% used
RATE_LIMIT_COOLDOWN_CAP = 10           # max seconds to sleep per cooldown iteration

# Exit codes
EXIT_SUCCESS = 0
EXIT_ERROR = 1
EXIT_BAD_ARGS = 2

client = OpenAI(api_key=API_KEY)


# =========================
# UTILS
# =========================
def clean_text(text) -> str | None:
    if text is None:
        return None
    return str(text).replace("\ufeff", "").strip() or None


def infer_company_from_subject(subject: str | None) -> str | None:
    """Best-effort fallback when the extractor does not return a merchant name."""
    subject = clean_text(subject)
    if not subject:
        return None

    normalized = subject
    while True:
        updated = re.sub(r"^\s*(?:fw|fwd|re)\s*:\s*", "", normalized, flags=re.IGNORECASE)
        if updated == normalized:
            break
        normalized = updated

    patterns = [
        r"your\s+(.+?)\s+order(?:\b|:)",
        r"order\s+from\s+(.+?)(?:\b|:)",
        r"thanks\s+.+?\s+for\s+your\s+purchase\s+with\s+(.+?)(?:\b|!|\.|:)",
    ]

    for pattern in patterns:
        match = re.search(pattern, normalized, flags=re.IGNORECASE)
        if match:
            company = clean_text(match.group(1))
            if company:
                return company.strip(" -,:;.!?")

    return None


def strip_bom_from_argv(argv: list[str]) -> None:
    """Strip UTF-8 BOM (U+FEFF) from each argv entry in place (e.g. Power Automate)."""
    for i in range(len(argv)):
        if isinstance(argv[i], str):
            argv[i] = argv[i].replace("\ufeff", "")


# =========================
# HTML PARSING
# =========================
_TRACK_KEYWORD_RE = re.compile(r"track|shipment|delivery|carrier", re.IGNORECASE)
_ALNUM_RE = re.compile(r"^[A-Z0-9]+$", re.IGNORECASE)
_HAS_DIGIT_RE = re.compile(r"\d")


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
    when BeautifulSoup flattens everything to plain text."""
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


def parse_email_html(file_path: Path) -> tuple[str, list[dict]]:
    """Single-pass HTML parse returning (cleaned_text, all_links)."""
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        html = f.read()

    soup = BeautifulSoup(html, "html.parser")

    # Extract links before modifying the DOM
    seen: set[tuple[str, str]] = set()
    links: list[dict] = []
    for a_tag in soup.find_all("a", href=True):
        href = a_tag["href"].strip()
        if not href:
            continue
        anchor_text = a_tag.get_text(separator=" ").strip()
        key = (anchor_text, href)
        if key not in seen:
            seen.add(key)
            links.append({"anchor_text": anchor_text, "href": href})

    _remove_hidden_elements(soup)
    _convert_tables_to_markdown(soup)

    for tag in soup(["script", "style", "noscript", "svg", "meta", "head"]):
        tag.decompose()

    text = soup.get_text(separator="\n")
    lines = [line.strip() for line in text.splitlines()]
    lines = [line for line in lines if line]
    cleaned = "\n".join(lines)

    return cleaned, links


def _score_tracking_candidate(seg: str, near_track_keyword: bool) -> float:
    """Score how likely a URL/anchor segment is a tracking number (0.0–1.0).
    Pure-letter strings never reach here — they're filtered out upstream."""
    score = 0.5
    digit_ratio = sum(c.isdigit() for c in seg) / len(seg)
    score += digit_ratio * 0.2
    if 10 <= len(seg) <= 22:
        score += 0.15
    elif len(seg) > 30:
        score -= 0.2
    if near_track_keyword:
        score += 0.15
    return round(min(max(score, 0.0), 1.0), 2)


def extract_tracking_candidates(links: list[dict]) -> list[dict]:
    """Scan link URLs and anchor text for alphanumeric segments that contain
    digits (tracking numbers never contain real English words). Returns
    deduplicated candidates sorted by confidence, highest first."""
    seen: set[str] = set()
    candidates: list[dict] = []

    for link in links:
        href = link.get("href", "")
        anchor = link.get("anchor_text", "")
        has_kw = bool(_TRACK_KEYWORD_RE.search(href + " " + anchor))

        for seg in re.split(r"[/?&=\-_#]+", href):
            seg = seg.strip()
            if len(seg) < 7 or not _ALNUM_RE.match(seg) or not _HAS_DIGIT_RE.search(seg):
                continue
            key = seg.upper()
            if key in seen:
                continue
            seen.add(key)
            candidates.append({
                "value": seg,
                "confidence": _score_tracking_candidate(seg, has_kw),
                "source_url": href,
                "context": anchor or None,
            })

        anchor_clean = anchor.strip()
        if len(anchor_clean) >= 7 and _ALNUM_RE.match(anchor_clean) \
           and _HAS_DIGIT_RE.search(anchor_clean):
            key = anchor_clean.upper()
            if key not in seen:
                seen.add(key)
                candidates.append({
                    "value": anchor_clean,
                    "confidence": _score_tracking_candidate(anchor_clean, has_kw),
                    "source_url": href,
                    "context": "anchor text",
                })

    candidates.sort(key=lambda c: c["confidence"], reverse=True)
    return candidates


# =========================
# OPTIONAL: trim very long text
# =========================
def trim_text(text: str, max_chars: int = 50000) -> str:
    if len(text) <= max_chars:
        return text
    return text[:max_chars]


# =========================
# RATE-LIMIT HELPERS
# =========================
def _parse_reset_duration(reset_str: str) -> float:
    """Parse OpenAI reset-time strings like '6m0s', '2s', '200ms' into seconds."""
    if not reset_str:
        return 0.0
    total = 0.0
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*(ms|s|m|h)", reset_str):
        val, unit = float(m.group(1)), m.group(2)
        if unit == "ms":
            total += val / 1000
        elif unit == "s":
            total += val
        elif unit == "m":
            total += val * 60
        elif unit == "h":
            total += val * 3600
    return total


def _estimate_remaining(remain: int, limit: int, reset_secs: float, elapsed: float) -> int:
    """Estimate remaining quota after `elapsed` seconds of natural refill (linear approximation)."""
    if not limit or reset_secs <= 0:
        return remain
    if elapsed >= reset_secs:
        return limit  # full reset assumed
    return min(limit, remain + int((elapsed / reset_secs) * limit))


def _check_and_throttle(headers) -> float:
    """Inspect rate-limit headers; if usage >= threshold, sleep and re-check in a loop.

    After each sleep the function re-estimates current usage using elapsed time so it
    only proceeds once usage is confirmed back below the threshold.  Returns total
    seconds waited.
    """
    limit_req      = int(headers.get("x-ratelimit-limit-requests",     0))
    remain_req     = int(headers.get("x-ratelimit-remaining-requests", 0))
    limit_tok      = int(headers.get("x-ratelimit-limit-tokens",       0))
    remain_tok     = int(headers.get("x-ratelimit-remaining-tokens",   0))
    reset_req_secs = _parse_reset_duration(headers.get("x-ratelimit-reset-requests", ""))
    reset_tok_secs = _parse_reset_duration(headers.get("x-ratelimit-reset-tokens",   ""))

    threshold   = RATE_LIMIT_THROTTLE_THRESHOLD
    total_slept = 0.0
    start_mono  = time.monotonic()

    while True:
        elapsed = time.monotonic() - start_mono

        # Re-estimate remaining quota accounting for natural refill since the response arrived
        est_req = _estimate_remaining(remain_req, limit_req, reset_req_secs, elapsed)
        est_tok = _estimate_remaining(remain_tok, limit_tok, reset_tok_secs, elapsed)

        pct_req = (1 - est_req / limit_req) * 100 if limit_req else 0
        pct_tok = (1 - est_tok / limit_tok) * 100 if limit_tok else 0

        req_hot = bool(limit_req and (est_req / limit_req) <= (1 - threshold))
        tok_hot = bool(limit_tok and (est_tok / limit_tok) <= (1 - threshold))

        prefix = "  (re-check)" if total_slept > 0 else " "
        print(
            f"{prefix} Rate limit —"
            f" requests: {est_req}/{limit_req} ({pct_req:.0f}% used)"
            f"  |  tokens: {est_tok}/{limit_tok} ({pct_tok:.0f}% used)"
        )

        if not req_hot and not tok_hot:
            if total_slept > 0:
                print(f"  Usage confirmed below {threshold*100:.0f}% threshold — proceeding.")
            return total_slept

        # Build a label showing exactly which resource(s) triggered the cooldown
        triggered = []
        if req_hot:
            triggered.append(f"requests ({pct_req:.0f}% used)")
        if tok_hot:
            triggered.append(f"tokens ({pct_tok:.0f}% used)")
        trigger_label = " + ".join(triggered)

        # Sleep only as long as needed for the hot resource(s), capped for faster cycling
        rem_req_reset = max(0.0, reset_req_secs - elapsed) if req_hot else 0.0
        rem_tok_reset = max(0.0, reset_tok_secs - elapsed) if tok_hot else 0.0
        full_wait  = max(rem_req_reset, rem_tok_reset, 1.0)
        actual_wait = min(full_wait, RATE_LIMIT_COOLDOWN_CAP)

        print(
            f"  Throttling on {trigger_label} (>= {threshold*100:.0f}% threshold) — "
            f"waiting {actual_wait:.1f}s (full reset in {full_wait:.1f}s)..."
        )
        time.sleep(actual_wait)
        total_slept += actual_wait


# =========================
# PRIMARY EXTRACTION: OPENAI
# =========================
def extract_with_openai(
    text_only: str,
    candidates: list[dict],
    subject: str | None = None,
) -> dict:
    """Direct extraction of purchase details using OpenAI (primary path)."""
    candidates_json_str = json.dumps(candidates, indent=2) if candidates else "[]"
    subject_section = f"\nEMAIL SUBJECT: {subject}" if subject else ""

    prompt = f"""You are extracting structured purchase information from text that came from an HTML email.

Important rules:
1. Use ONLY the provided text, tracking candidates, and subject line as the source of truth.
2. Find the PURCHASE date/time, NOT the email received date or shipment/delivery date.
3. If a value is missing or unclear, return null for scalars and empty arrays for lists.
4. Distinguish ORDER NUMBER from TRACKING NUMBER — they are different things.
   An order number is the retailer's confirmation/order ID (e.g. "112-3456789-1234567").
   A tracking number is assigned by a shipping carrier for package delivery.
5. total_amount_paid should be the exact total paid as a number if possible.
6. tax_paid should be the tax dollar amount if present, or null if unknown.
7. purchase_datetime should be "YYYY-MM-DD HH:MM:SS" or "YYYY-MM-DD".
8. company should be the retailer/store/merchant, not the recipient's name.
9. order_number: check the subject line first, then the body. Strip any leading "#".
10. Do NOT guess. If something is not clearly present, use null or an empty array.
11. email_category: Classify this email into exactly ONE of these categories:
    - "Order Placed": Initial purchase email, typically no tracking info yet.
    - "Order Received": Merchant acknowledges receipt of the order.
    - "Order Confirmed": Explicit confirmation the order is being processed.
    - "Order Delayed": Notification that the order or shipment is delayed.
    - "Order Shipped": Shipment notification, usually includes tracking number(s).
    - "Delivery Confirmation": Package was delivered successfully.
    - "Purchased Gift Card": A gift card purchase confirmation or delivery (digital or physical).
    - "Unknown": Does not clearly fit any of the above.
    Contextual hints: "Order Placed" emails rarely have tracking numbers. "Order Shipped"
    emails almost always have tracking info. "Delivery Confirmation" emails mention delivery
    completion. Gift card emails mention gift card value, redemption codes, or gift card delivery.
    Use these signals to guide your choice.
12. email_category_confidence: Your confidence (0–100) in the chosen category.

TRACKING CANDIDATES:
Below are alphanumeric sequences pre-extracted from email link URLs and anchor
text. Each has a confidence score (0.0–1.0). They were selected because they
contain digits and are not English words — but some may be product IDs or
session tokens rather than real tracking numbers.
- Review each candidate against the email text to confirm or reject it.
- For tracking_numbers: include only confirmed tracking number strings.
- For tracking_links: include the source_url of any confirmed tracking candidate.
- If no candidates are listed, check the email text for tracking numbers instead.

{candidates_json_str}
{subject_section}

EMAIL TEXT:
{text_only}""".strip()

    api_kwargs = dict(
        model=MODEL,
        messages=[
            {
                "role": "developer",
                "content": (
                    "Extract purchase details from email text. "
                    "Return only valid structured JSON data."
                ),
            },
            {"role": "user", "content": prompt},
        ],
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "purchase_details",
                "schema": {
                    "type": "object",
                    "properties": {
                        "company": {
                            "type": ["string", "null"],
                            "description": "Retailer/store/company associated with the order email.",
                        },
                        "order_number": {
                            "type": ["string", "null"],
                            "description": "Retailer order/confirmation number. Strip leading '#'.",
                        },
                        "purchase_datetime": {
                            "type": ["string", "null"],
                            "description": "Purchase/order datetime, not shipment or delivery date.",
                        },
                        "total_amount_paid": {
                            "type": ["number", "null"],
                            "description": "Exact total amount paid.",
                        },
                        "tax_paid": {
                            "type": ["number", "null"],
                            "description": "Tax amount in dollars, or null if unknown.",
                        },
                        "tracking_numbers": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "List of shipment tracking numbers.",
                        },
                        "tracking_links": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "List of tracking URLs.",
                        },
                        "email_category": {
                            "type": "string",
                            "enum": [
                                "Order Placed",
                                "Order Received",
                                "Order Confirmed",
                                "Order Delayed",
                                "Order Shipped",
                                "Delivery Confirmation",
                                "Purchased Gift Card",
                                "Unknown",
                            ],
                            "description": "Category of this email.",
                        },
                        "email_category_confidence": {
                            "type": "number",
                            "description": "Confidence (0-100) in the assigned email_category.",
                        },
                    },
                    "required": [
                        "company",
                        "order_number",
                        "purchase_datetime",
                        "total_amount_paid",
                        "tax_paid",
                        "tracking_numbers",
                        "tracking_links",
                        "email_category",
                        "email_category_confidence",
                    ],
                    "additionalProperties": False,
                },
            },
        },
        temperature=0,
    )

    total_waited = 0.0
    for attempt in range(1, RATE_LIMIT_MAX_RETRIES + 1):
        try:
            raw = client.chat.completions.with_raw_response.create(**api_kwargs)
            response = raw.parse()
            total_waited += _check_and_throttle(raw.headers)
            break
        except RateLimitError:
            print(f"  Rate limit hit (attempt {attempt}/{RATE_LIMIT_MAX_RETRIES}) — waiting {RATE_LIMIT_RETRY_WAIT}s...")
            time.sleep(RATE_LIMIT_RETRY_WAIT)
            total_waited += RATE_LIMIT_RETRY_WAIT
    else:
        raise RuntimeError(
            f"OpenAI rate limit not cleared after {RATE_LIMIT_MAX_RETRIES} retries "
            f"({total_waited:.0f}s total wait)"
        )

    if total_waited > 0:
        print(f"  Total rate-limit wait: {total_waited:.1f}s")

    content = response.choices[0].message.content
    data = json.loads(content)

    if not isinstance(data.get("tracking_numbers"), list):
        data["tracking_numbers"] = []
    if not isinstance(data.get("tracking_links"), list):
        data["tracking_links"] = []

    return data


# =========================
# FALLBACK EXTRACTION: LOCAL LLM
# =========================
LOCAL_PROMPT_TEMPLATE = """You are extracting structured purchase/order information from an email.

RULES:
1. Return ONLY valid JSON. No markdown fences, no explanation, no extra text.
2. Use null for missing scalar fields and empty arrays for missing list fields.
3. Do NOT guess. If a value is not clearly present, use null or an empty array.
4. ORDER NUMBER = retailer confirmation ID. TRACKING NUMBER = shipping carrier ID. They are different.
5. Distinguish PURCHASE DATE from shipment/delivery dates.
6. purchase_datetime format: "YYYY-MM-DD HH:MM:SS" or "YYYY-MM-DD".
7. company is the retailer/store/merchant name.
8. order_number: check subject line first, then body. Strip any leading "#".
9. TRACKING CANDIDATES below were pre-extracted from email links. Confirm or reject each.
   Include confirmed values in tracking_numbers and their source_url in tracking_links.
   If no candidates, check the email text for tracking numbers.
10. email_category: Classify this email as exactly one of:
    "Order Placed", "Order Received", "Order Confirmed", "Order Delayed",
    "Order Shipped", "Delivery Confirmation", "Purchased Gift Card", or "Unknown".
    Hints: "Order Placed" emails rarely have tracking. "Order Shipped" usually has tracking.
    "Delivery Confirmation" mentions delivery completion. "Purchased Gift Card" involves
    gift card purchases, redemption codes, or gift card delivery.
11. email_category_confidence: Your confidence (0-100) in the chosen category.

Return JSON matching this exact schema:
{{
  "company": <string or null>,
  "order_number": <string or null>,
  "purchase_datetime": <string or null>,
  "total_amount_paid": <number or null>,
  "tax_paid": <number or null>,
  "tracking_numbers": [<string>, ...],
  "tracking_links": [<string>, ...],
  "email_category": <string>,
  "email_category_confidence": <number 0-100>
}}

EMAIL SUBJECT: {subject}

TRACKING CANDIDATES:
{candidates_json}

EMAIL TEXT:
{text}"""


def _extract_first_json_object(s: str) -> str | None:
    """Find the first top-level JSON object in a string (handles nested braces)."""
    start = s.find("{")
    if start == -1:
        return None
    depth = 0
    in_string = False
    escape = False
    for i in range(start, len(s)):
        c = s[i]
        if escape:
            escape = False
            continue
        if in_string:
            if c == "\\":
                escape = True
            elif c == '"':
                in_string = False
            continue
        if c == '"':
            in_string = True
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return s[start : i + 1]
    return None


def _strip_markdown_json_fence(s: str) -> str:
    m = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", s, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return s.strip()


def _parse_local_llm_json_text(cleaned: str) -> dict | None:
    """Try several strategies to extract a JSON dict; return dict or None."""
    candidates: list[str] = []
    t = cleaned.strip()
    if t:
        candidates.append(t)
    fenced = _strip_markdown_json_fence(cleaned)
    if fenced and fenced not in candidates:
        candidates.append(fenced.strip())

    for blob in candidates:
        try:
            out = json.loads(blob)
            if isinstance(out, dict):
                return out
        except json.JSONDecodeError:
            pass

    for blob in candidates:
        sub = _extract_first_json_object(blob)
        if sub:
            try:
                out = json.loads(sub)
                if isinstance(out, dict):
                    return out
            except json.JSONDecodeError:
                pass

    return None


def _normalize_local_result(raw: dict) -> dict:
    """Normalize local LLM output to the standard program schema.
    Handles flat keys, camelCase variants, and confidence-annotated wrappers."""
    def _scalar(*keys: str):
        for k in keys:
            v = raw.get(k)
            if isinstance(v, dict):
                v = v.get("value")
            if v is not None:
                return v
        return None

    def _list_field(*keys: str) -> list:
        for k in keys:
            entries = raw.get(k)
            if isinstance(entries, list):
                result = []
                for item in entries:
                    val = item.get("value") if isinstance(item, dict) else item
                    if val is not None and str(val).strip():
                        result.append(str(val).strip())
                return result
        return []

    return {
        "company": _scalar("company", "company_name"),
        "order_number": _scalar("order_number", "orderNumber"),
        "purchase_datetime": _scalar("purchase_datetime", "purchaseDate"),
        "total_amount_paid": _scalar("total_amount_paid", "totalPaid"),
        "tax_paid": _scalar("tax_paid", "taxPaid"),
        "tracking_numbers": _list_field("tracking_numbers", "trackingNumbers"),
        "tracking_links": _list_field("tracking_links", "trackingLinks"),
        "email_category": _scalar("email_category", "emailCategory", "category"),
        "email_category_confidence": _scalar(
            "email_category_confidence", "emailCategoryConfidence",
            "category_confidence",
        ),
    }


def extract_with_local_llm(
    text_only: str,
    candidates: list[dict],
    subject: str | None = None,
) -> dict:
    prompt = LOCAL_PROMPT_TEMPLATE.format(
        subject=subject or "(none)",
        candidates_json=json.dumps(candidates, indent=2),
        text=text_only,
    )

    try:
        result = subprocess.run(
            ["ollama", "run", LOCAL_MODEL, "--format", "json"],
            input=prompt,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=120,
        )
        raw_output = result.stdout.strip()
    except FileNotFoundError:
        print("  WARNING: ollama not found on PATH. Skipping local extraction.")
        return {"_error": "ollama_not_found"}
    except subprocess.TimeoutExpired:
        print("  WARNING: local LLM timed out.")
        return {"_error": "local_llm_timeout"}
    except Exception as e:
        print(f"  WARNING: local LLM subprocess failed: {e}")
        return {"_error": str(e)}

    if result.returncode != 0:
        print(f"  WARNING: ollama returned exit code {result.returncode}")
        stderr_snippet = (result.stderr or "")[:300]
        if stderr_snippet:
            print(f"  stderr: {stderr_snippet}")
        return {"_error": f"ollama_exit_{result.returncode}"}

    cleaned = raw_output.strip()
    cleaned = re.sub(r"<think>[\s\S]*?</think>", "", cleaned).strip()

    if cleaned.startswith("```"):
        lines = cleaned.splitlines()
        if len(lines) >= 2:
            lines = lines[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            cleaned = "\n".join(lines).strip()

    parsed = _parse_local_llm_json_text(cleaned)
    if parsed is None:
        snippet = (cleaned[:400] + "\u2026") if len(cleaned) > 400 else cleaned
        print("  WARNING: could not parse local LLM JSON output.")
        print(f"  (first ~400 chars of model output): {snippet!r}")
        return {"_error": "json_parse_failed", "_raw": raw_output[:800]}

    return _normalize_local_result(parsed)


# =========================
# ARGS
# =========================
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract structured purchase details from HTML emails."
    )
    parser.add_argument(
        "--base-dir",
        required=False,
        default=os.getenv("BASE_DIR"),
        dest="base_dir",
        help="Root project folder. Defaults to BASE_DIR in .env if not provided.",
    )
    parser.add_argument(
        "--file",
        default=None,
        help=(
            "Filename (or full path) of a single HTML file to process. "
            "A bare filename is resolved inside <base-dir>/email_contents/html/. "
            "When provided, only that file is processed and the result is appended "
            "to the output JSON."
        ),
    )
    parser.add_argument(
        "--subject",
        default=None,
        help="Email subject line. Passed to the LLM and embedded in the output JSON.",
    )
    parser.add_argument(
        "--sender-name",
        default=None,
        dest="sender_name",
        help="Display name of the sender. Embedded in the output JSON as-is.",
    )
    parser.add_argument(
        "--email",
        default=None,
        help="Sender email address. Embedded in the output JSON as-is.",
    )
    return parser.parse_args()


# =========================
# PROCESS ONE FILE
# =========================
def process_file(
    file_path: Path,
    subject: str | None,
    sender_name: str | None,
    email: str | None,
) -> dict:
    empty_extraction = {
        "company": None,
        "order_number": None,
        "purchase_datetime": None,
        "total_amount_paid": None,
        "tax_paid": None,
        "tracking_numbers": [],
        "tracking_links": [],
        "email_category": "Unknown",
        "email_category_confidence": 0,
    }

    try:
        print(f"Processing: {file_path}")

        text_only, all_links = parse_email_html(file_path)
        text_only = trim_text(text_only, max_chars=50000)
        candidates = extract_tracking_candidates(all_links)

        extracted = None

        # Primary: OpenAI direct extraction
        if API_KEY:
            print("  Running OpenAI extraction...")
            try:
                extracted = extract_with_openai(
                    text_only, candidates, subject=subject
                )
            except Exception as e:
                print(f"  WARNING: OpenAI extraction failed: {e}")
                extracted = None

        # Fallback: Local LLM (when no API key or OpenAI failed)
        if extracted is None:
            print("  Running local LLM extraction (fallback)...")
            local_data = extract_with_local_llm(
                text_only, candidates, subject=subject
            )
            if "_error" not in local_data:
                extracted = local_data
            else:
                print("  WARNING: local LLM extraction also failed.")
                extracted = empty_extraction

        if not clean_text(extracted.get("company")):
            extracted["company"] = infer_company_from_subject(subject)

        file_uri = "file:///" + str(file_path.resolve()).replace("\\", "/")

        raw_category = extracted.get("email_category", "Unknown")
        raw_confidence = extracted.get("email_category_confidence", 0)
        try:
            raw_confidence = float(raw_confidence)
        except (TypeError, ValueError):
            raw_confidence = 0.0

        if raw_category not in VALID_CATEGORIES:
            raw_category = "Unknown"
        if raw_confidence < CATEGORY_CONFIDENCE_THRESHOLD:
            raw_category = "Unknown"

        record: dict = {
            "source_file": clean_text(file_path),
            "source_file_link": file_uri,
            "subject": clean_text(subject),
            "sender_name": clean_text(sender_name),
            "email": clean_text(email),
            "company": clean_text(extracted.get("company")),
            "order_number": clean_text(extracted.get("order_number")),
            "purchase_datetime": clean_text(extracted.get("purchase_datetime")),
            "total_amount_paid": extracted.get("total_amount_paid"),
            "tax_paid": extracted.get("tax_paid"),
            "tracking_numbers": extracted.get("tracking_numbers", []),
            "tracking_links": extracted.get("tracking_links", []),
            "email_category": raw_category,
            "email_category_confidence": raw_confidence,
        }

        if not isinstance(record["tracking_numbers"], list):
            record["tracking_numbers"] = []
        if not isinstance(record["tracking_links"], list):
            record["tracking_links"] = []

        return record

    except Exception as e:
        return {
            "source_file": clean_text(file_path),
            "source_file_link": None,
            "subject": clean_text(subject),
            "sender_name": clean_text(sender_name),
            "email": clean_text(email),
            "error": clean_text(e),
            "company": None,
            "order_number": None,
            "purchase_datetime": None,
            "total_amount_paid": None,
            "tax_paid": None,
            "tracking_numbers": [],
            "tracking_links": [],
            "email_category": "Unknown",
            "email_category_confidence": 0,
        }


# =========================
# FILE RENAMING
# =========================
def _next_file_number(html_folder: Path) -> int:
    """Find the highest existing fileN.html number and return N+1."""
    pattern = re.compile(r"^file(\d+)\.html?$", re.IGNORECASE)
    max_n = 0
    for p in html_folder.iterdir():
        m = pattern.match(p.name)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def rename_single_file(file_path: Path, html_folder: Path) -> Path:
    """Rename an incoming HTML file to the next sequential fileN.html name."""
    n = _next_file_number(html_folder)
    new_name = f"file{n}.html"
    new_path = html_folder / new_name
    file_path.rename(new_path)
    print(f"  Renamed: {file_path.name} -> {new_name}")
    return new_path


def rename_html_files_sequential(html_folder: Path) -> list[Path]:
    """Rename all HTML files in the folder to file1.html, file2.html, etc.
    Returns the new paths in the same order."""
    html_files = sorted(
        [p for p in html_folder.iterdir() if p.suffix.lower() in (".html", ".htm")],
        key=lambda p: p.stat().st_mtime,
    )
    if not html_files:
        return []

    temp_names: list[tuple[Path, Path]] = []
    for i, fp in enumerate(html_files, start=1):
        tmp = html_folder / f"__tmp_rename_{i}__.html"
        fp.rename(tmp)
        temp_names.append((tmp, html_folder / f"file{i}.html"))

    new_paths: list[Path] = []
    for tmp, final in temp_names:
        tmp.rename(final)
        new_paths.append(final)
        print(f"  Renamed -> {final.name}")

    return new_paths


# =========================
# DEDUPLICATION
# =========================
def compute_file_hash(file_path: Path) -> str:
    """SHA-256 hex digest of a file's raw bytes."""
    h = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def _load_existing_results(output_path: Path) -> list[dict]:
    """Load results.json, returning [] if missing or unreadable."""
    if not output_path.exists():
        return []
    for enc in ("utf-8-sig", "utf-16", "utf-8", "latin-1"):
        try:
            with open(output_path, "r", encoding=enc) as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue
    return []


def _known_hashes(results: list[dict]) -> set[str]:
    """Collect all content_hash values from existing results."""
    return {r["content_hash"] for r in results if r.get("content_hash")}


# =========================
# MAIN
# =========================
def main():
    args = parse_args()

    if os.getenv("DEMO_MODE") == "1":
        args.email = "johndoe123@gmail.com"
        args.sender_name = "John Doe"
        print("DEMO MODE: sender overridden to John Doe <johndoe123@gmail.com>")

    if not args.base_dir:
        raise ValueError("BASE_DIR is not set. Add it to python_files/.env or pass --base-dir.")

    base = Path(args.base_dir)
    html_folder = base / "email_contents" / "html"
    output_path = base / "email_contents" / "json" / "results.json"

    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not API_KEY:
        print(
            f"WARNING: {OPENAI_API_KEY_ENV} is not set. "
            "OpenAI extraction will be skipped; using local LLM only."
        )

    # Single-file mode
    if args.file:
        candidate = Path(args.file)
        if not candidate.is_absolute() and candidate.parent == Path("."):
            candidate = html_folder / candidate
        file_path = candidate
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        results = _load_existing_results(output_path)
        hashes = _known_hashes(results)

        file_hash = compute_file_hash(file_path)
        if file_hash in hashes:
            print(f"SKIPPED (duplicate content): {file_path.name}")
            file_path.unlink()
            print(f"  Deleted temp file: {file_path.name}")
            for r in results:
                if r.get("content_hash") == file_hash:
                    r["duplicate_on_last_run"] = 1
                    break
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
            return

        file_path = rename_single_file(file_path, html_folder)

        entry = process_file(file_path, args.subject, args.sender_name, args.email)
        entry["content_hash"] = file_hash
        entry["duplicate_on_last_run"] = 0

        results.append(entry)

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        print(f"Appended result to: {output_path}")
        return

    # Batch mode: process entire html folder
    if not html_folder.exists():
        raise FileNotFoundError(f"HTML input folder not found: {html_folder}")

    print("Renaming HTML files to sequential format...")
    html_files = rename_html_files_sequential(html_folder)

    if not html_files:
        print(f"No HTML files found in: {html_folder}")
        return

    results = _load_existing_results(output_path)
    hashes = _known_hashes(results)
    new_count = 0

    for fp in html_files:
        file_hash = compute_file_hash(fp)
        if file_hash in hashes:
            print(f"  SKIPPED (duplicate content): {fp.name}")
            for r in results:
                if r.get("content_hash") == file_hash:
                    r["duplicate_on_last_run"] = 1
                    break
            continue
        entry = process_file(fp, args.subject, args.sender_name, args.email)
        entry["content_hash"] = file_hash
        entry["duplicate_on_last_run"] = 0
        results.append(entry)
        hashes.add(file_hash)
        new_count += 1

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    print(f"\nDone. {new_count} new + {len(results) - new_count} existing = {len(results)} total in: {output_path}")


class _Tee:
    """Writes to both an original stream and a log file simultaneously."""
    def __init__(self, log_path: Path, original_stream):
        self._file = open(log_path, "a", encoding="utf-8")
        self._original = original_stream
    def write(self, msg):
        self._original.write(msg)
        self._file.write(msg.replace("\ufeff", "") if isinstance(msg, str) else msg)
    def flush(self):
        self._original.flush()
        self._file.flush()
    def close(self):
        self._file.close()


if __name__ == "__main__":
    from datetime import datetime

    strip_bom_from_argv(sys.argv)

    _log_path = Path(os.getenv("BASE_DIR")) / "programFileOutput.txt"
    _tee = _Tee(_log_path, sys.stdout)
    sys.stdout = _tee
    sys.stderr = _Tee(_log_path, sys.stderr)

    _start_time = time.time()

    print(f"\n{'='*60}")
    print(f"Run started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Args: {sys.argv[1:]}")
    print(f"{'='*60}")

    _original_stdout = _tee._original
    _original_stderr = sys.stderr._original

    try:
        main()
        _elapsed = time.time() - _start_time
        print(f"Run finished successfully. Total operation time: {_elapsed:.2f}s")
    except SystemExit as e:
        _elapsed = time.time() - _start_time
        print(f"Total operation time: {_elapsed:.2f}s")
        if e.code == EXIT_BAD_ARGS:
            print("\nERROR: Invalid or missing arguments.")
            print("Set BASE_DIR and OPENAI_API_KEY in python_files/.env, or pass --base-dir as an argument.")
            print("Optional args: --file, --subject, --sender-name, --email")
        sys.stdout = _original_stdout
        sys.stderr = _original_stderr
        _tee.close()
        sys.exit(e.code)
    except Exception as e:
        _elapsed = time.time() - _start_time
        print(f"\nERROR: {e}")
        print(f"Total operation time: {_elapsed:.2f}s")
        sys.stdout = _original_stdout
        sys.stderr = _original_stderr
        _tee.close()
        sys.exit(1)

    sys.stdout = _original_stdout
    sys.stderr = _original_stderr
    _tee.close()
