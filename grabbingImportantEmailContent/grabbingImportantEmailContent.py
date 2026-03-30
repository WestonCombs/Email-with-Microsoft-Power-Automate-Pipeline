import argparse
import hashlib
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

# python_files/ — .env must load before htmlHandler (trace uses BASE_DIR from .env)
_PYTHON_FILES_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_PYTHON_FILES_DIR))

from dotenv import load_dotenv

load_dotenv(_PYTHON_FILES_DIR / ".env")

from openai import OpenAI, RateLimitError

from htmlHandler.admin_log import trace as _trace
from htmlHandler.convertHTMLToPlaintext import convert as html_to_plaintext
from htmlHandler.htmlValuesExtractionByAttribute.ValuesExtractionByAttribute import extract_attribute_values
from htmlHandler.isHrefTrackingLink import determine_tracking_link

# =========================
# CONFIG
# =========================
OPENAI_API_KEY_ENV = "OPENAI_API_KEY"
BASE_DIR_ENV = "BASE_DIR"
API_KEY = os.getenv(OPENAI_API_KEY_ENV)

MODEL = "gpt-4o-mini"

VALID_CATEGORIES = [
    "Delivery Shipped From Sender",
    "Delivery On The Way",
    "Delivery Arrived",
    "Order Received By Vendor",
    "Order Confirmed",
    "Gift Card Purchase",
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

OPENAI_USAGE_REL = Path("email_contents") / "openai usage"

# Set once per process run when main() initializes the flow usage log (CLI entry).
_flow_usage_log_path: Path | None = None


def _next_flow_usage_index(usage_dir: Path) -> int:
    """Next filename is <n>.txt where n is one greater than the highest existing N.txt."""
    if not usage_dir.exists():
        return 1
    max_n = 0
    pattern = re.compile(r"^(\d+)\.txt$", re.IGNORECASE)
    for p in usage_dir.iterdir():
        if p.is_file():
            m = pattern.match(p.name)
            if m:
                max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def init_flow_usage_log(base_dir: Path, flow_started_at: datetime) -> None:
    """Create a new numbered file for this flow; first line is when the flow started."""
    global _flow_usage_log_path
    usage_dir = base_dir / OPENAI_USAGE_REL
    usage_dir.mkdir(parents=True, exist_ok=True)
    n = _next_flow_usage_index(usage_dir)
    path = usage_dir / f"{n}.txt"
    header = f"Flow started: {flow_started_at.strftime('%Y-%m-%d %H:%M:%S')}\n\n"
    path.write_text(header, encoding="utf-8")
    _flow_usage_log_path = path


def _read_last_cumulative_tokens(log_path: Path) -> int:
    if not log_path.exists():
        return 0
    last = 0
    try:
        text = log_path.read_text(encoding="utf-8")
    except OSError:
        return 0
    for line in text.splitlines():
        m = re.search(r"cumulative_total=(\d+)", line)
        if m:
            last = int(m.group(1))
    return last


def append_openai_usage_log(
    *,
    prompt_tokens: int,
    completion_tokens: int,
    total_tokens: int,
) -> int:
    """Append one line for this flow's file; cumulative is for this flow only. Returns cumulative."""
    global _flow_usage_log_path
    if _flow_usage_log_path is None:
        return 0
    log_path = _flow_usage_log_path

    prev = _read_last_cumulative_tokens(log_path)
    cumulative = prev + total_tokens
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    line = (
        f"{ts} | "
        f"prompt_tokens={prompt_tokens} completion_tokens={completion_tokens} "
        f"total_tokens={total_tokens} | cumulative_total={cumulative}\n"
    )

    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line)
    return cumulative


# =========================
# UTILS
# =========================
def clean_text(text) -> str | None:
    if text is None:
        return None
    return str(text).replace("\ufeff", "").strip() or None


def read_email_html_file(file_path: Path) -> tuple[str, str]:
    """Read HTML saved by Outlook / Power Automate.

    Those tools often write **UTF-16 LE** (with or without BOM). Opening as UTF-8
    produces mojibake: no ``href=`` substring, so link extraction returns [].
    """
    raw = file_path.read_bytes()
    if not raw:
        return "", "empty"

    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig"), "utf-8-sig"
    if raw.startswith(b"\xff\xfe"):
        return raw.decode("utf-16-le"), "utf-16-le (BOM)"
    if raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16-be"), "utf-16-be (BOM)"

    try:
        utf8 = raw.decode("utf-8")
    except UnicodeDecodeError:
        utf8 = None

    if utf8 is not None:
        head = utf8[:400_000]
        if "href=" in head or "href =" in head.lower():
            return utf8, "utf-8"

    # UTF-16 LE without BOM: lots of NUL bytes in first KB of "ASCII" HTML
    sample = raw[: min(10_000, len(raw))]
    if sample.count(b"\x00") > 20 and len(raw) > 200:
        try:
            s16 = raw.decode("utf-16-le")
            if "href=" in s16 or "<a " in s16.lower():
                return s16, "utf-16-le (no BOM)"
        except (UnicodeDecodeError, UnicodeError):
            pass

    if utf8 is not None:
        return utf8, "utf-8 (may be mojibake if no href= found)"

    try:
        s16 = raw.decode("utf-16-le")
        return s16, "utf-16-le (fallback)"
    except (UnicodeDecodeError, UnicodeError):
        pass

    return raw.decode("utf-8", errors="replace"), "utf-8-replace"


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
# EXTRACTION: OPENAI
# =========================
def extract_with_openai(
    text_only: str,
    subject: str | None = None,
) -> dict:
    """Extract purchase details using OpenAI (plain text only, no HTML)."""
    subject_section = f"\nEMAIL SUBJECT: {subject}" if subject else ""

    prompt = f"""You are extracting structured purchase information from text that came from an HTML email.

Important rules:
1. Use ONLY the provided text and subject line as the source of truth.
2. Find the PURCHASE date/time, NOT the email received date or shipment/delivery date.
3. If a value is missing or unclear, return null.
4. order_number is the retailer's confirmation/order ID (e.g. "112-3456789-1234567").
5. total_amount_paid should be the exact total paid as a number if possible.
6. tax_paid should be the tax dollar amount if present, or null if unknown.
7. purchase_datetime should be "YYYY-MM-DD HH:MM:SS" or "YYYY-MM-DD".
8. company should be the retailer/store/merchant, not the recipient's name.
9. order_number: check the subject line first, then the body. Strip any leading "#".
10. Do NOT guess. If something is not clearly present, use null.
11. tracking_number: The shipping carrier's tracking number (e.g. "1Z999AA10123456784",
    "9400111899223456789012"). This is NOT the order number — it is the alphanumeric ID
    assigned by UPS, FedEx, USPS, DHL, or another carrier for package tracking.
    If not clearly present, return null.
12. email_category: Classify this email into exactly ONE of these categories:
    - "Delivery Shipped From Sender": The seller/merchant has shipped the package (often includes carrier handoff or tracking).
    - "Delivery On The Way": In-transit updates — package is moving toward the recipient (out for delivery, on the way, etc.).
    - "Delivery Arrived": Delivered — the package has arrived or delivery is complete.
    - "Order Received By Vendor": The store received the order (submitted/queued) but not yet fully confirmed or shipped.
    - "Order Confirmed": The merchant explicitly confirms the order is accepted and being processed.
    - "Gift Card Purchase": Gift card purchase, delivery, redemption code, or balance notification.
    - "Unknown": Does not clearly fit any of the above.
    Contextual hints: Shipped-from-sender and on-the-way emails often mention tracking or carrier movement;
    arrived emails stress delivery completion. Vendor-received vs confirmed depends on wording (received vs
    confirmed/processing). Gift card emails mention gift card value, codes, or digital delivery.
    Use these signals to guide your choice.
13. email_category_confidence: Your confidence (0–100) in the chosen category.
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
                        "tracking_number": {
                            "type": ["string", "null"],
                            "description": (
                                "Shipping carrier tracking number "
                                "(e.g. UPS, FedEx, USPS, DHL). Not the order number. "
                                "Null if not clearly present."
                            ),
                        },
                        "email_category": {
                            "type": "string",
                            "enum": [
                                "Delivery Shipped From Sender",
                                "Delivery On The Way",
                                "Delivery Arrived",
                                "Order Received By Vendor",
                                "Order Confirmed",
                                "Gift Card Purchase",
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
                        "tracking_number",
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

    usage = getattr(response, "usage", None)
    if usage is not None and _flow_usage_log_path is not None:
        pt = getattr(usage, "prompt_tokens", None) or 0
        ct = getattr(usage, "completion_tokens", None) or 0
        tt = getattr(usage, "total_tokens", None)
        if tt is None:
            tt = int(pt) + int(ct)
        try:
            cumulative = append_openai_usage_log(
                prompt_tokens=int(pt),
                completion_tokens=int(ct),
                total_tokens=int(tt),
            )
            print(
                f"  OpenAI tokens this request: {tt} (cumulative this flow: {cumulative})"
            )
        except OSError as e:
            print(f"  WARNING: Could not write OpenAI usage log: {e}")

    content = response.choices[0].message.content
    data = json.loads(content)
    return data


# =========================
# ARGS
# =========================
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract structured purchase details from HTML emails."
    )
    parser.add_argument(
        "--file",
        default=None,
        help=(
            "Filename (or full path) of a single HTML file to process. "
            "A bare filename is resolved under email_contents/html/ relative to BASE_DIR from .env. "
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
        "tracking_number": None,
        "email_category": "Unknown",
        "email_category_confidence": 0,
    }

    extracted_hrefs: list[str] = []

    try:
        _trace(
            "PIPELINE",
            "=" * 55,
        )
        _trace(
            "PIPELINE",
            f"START process_file() — file={file_path.name}, "
            f"subject={subject!r}, sender={sender_name!r}, email={email!r}",
        )
        print(f"Processing: {file_path}")

        # ── READ FILE FROM DISK (UTF-8 or UTF-16 from Outlook / Power Automate) ──
        raw_html, html_encoding = read_email_html_file(file_path)
        _trace(
            "PIPELINE STEP 0  [read file]",
            f"Decoded HTML as {html_encoding!r}",
        )

        href_count_raw = raw_html.lower().count('href=')
        anchor_count = raw_html.lower().count('<a ')
        _trace(
            "PIPELINE STEP 0  [read file]",
            f"Read {file_path.name} — {len(raw_html):,} chars, "
            f"{len(raw_html.splitlines())} lines, "
            f"raw 'href=' count={href_count_raw}, '<a ' count={anchor_count}",
            raw_html[:300],
        )
        if len(raw_html) < 500:
            _trace(
                "PIPELINE STEP 0  [read file]",
                "WARNING: HTML is suspiciously short! Full content below:",
                raw_html,
            )

        # ── STEP 1: HTML -> plain text (convertHTMLToPlaintext) ──
        _trace(
            "PIPELINE STEP 1  [html_to_plaintext]",
            f"Sending {len(raw_html):,} chars of raw HTML to convertHTMLToPlaintext.convert()",
        )
        text_only = html_to_plaintext(raw_html)
        _trace(
            "PIPELINE STEP 1  [html_to_plaintext]",
            f"Got back {len(text_only):,} chars of plain text ({len(text_only.splitlines())} lines)",
            text_only[:300],
        )

        # ── STEP 2: Extract hrefs (ValuesExtractionByAttribute) ──
        _trace(
            "PIPELINE STEP 2  [extract_attribute_values]",
            f"Sending {len(raw_html):,} chars of raw HTML to extract_attribute_values(html, 'href')",
        )
        extracted_hrefs = extract_attribute_values(raw_html, "href")
        _trace(
            "PIPELINE STEP 2  [extract_attribute_values]",
            f"Got back {len(extracted_hrefs)} unique hrefs",
        )
        # Individual href logging disabled — tracking-link bug is resolved
        if extracted_hrefs:
            pass
            # for i, h in enumerate(extracted_hrefs[:5]):
            #     _trace("PIPELINE STEP 2  [extract_attribute_values]", f"  href[{i}]: {h}")
            # if len(extracted_hrefs) > 5:
            #     _trace(
            #         "PIPELINE STEP 2  [extract_attribute_values]",
            #         f"  ... and {len(extracted_hrefs) - 5} more",
            #     )
        else:
            _trace(
                "PIPELINE STEP 2  [extract_attribute_values]",
                "WARNING: extracted_hrefs is EMPTY — this is the bug. "
                "The regex found 0 matches even though raw 'href=' count "
                f"was {href_count_raw}. Check the raw HTML sample above.",
            )

        # ── STEP 3: Determine tracking link (isHrefTrackingLink) ──
        _trace(
            "PIPELINE STEP 3  [determine_tracking_link]",
            f"Sending {len(extracted_hrefs)} hrefs to determine_tracking_link()",
        )
        tracking_link = determine_tracking_link(extracted_hrefs)
        _trace(
            "PIPELINE STEP 3  [determine_tracking_link]",
            f"Got back tracking_link={tracking_link!r}",
        )

        # ── STEP 4: OpenAI extraction ────────────────────────────
        extracted = None

        if API_KEY:
            _trace(
                "PIPELINE STEP 4  [OpenAI]",
                f"Sending {len(text_only):,} chars of plain text to extract_with_openai()",
            )
            print("  Running OpenAI extraction...")
            try:
                extracted = extract_with_openai(text_only, subject=subject)
                _trace(
                    "PIPELINE STEP 4  [OpenAI]",
                    f"OpenAI returned: {json.dumps(extracted, default=str)[:400]}",
                )
            except Exception as e:
                _trace("PIPELINE STEP 4  [OpenAI]", f"FAILED: {e}")
                print(f"  WARNING: OpenAI extraction failed: {e}")
                extracted = None
        else:
            _trace("PIPELINE STEP 4  [OpenAI]", "SKIPPED — no API key")

        if extracted is None:
            if not API_KEY:
                print("  WARNING: No OpenAI API key; using empty/default fields.")
            else:
                print("  WARNING: Using empty/default fields (OpenAI unavailable or failed).")
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
            "tracking_number": clean_text(extracted.get("tracking_number")),
            "tracking_link": tracking_link,
            "extracted_links": extracted_hrefs,
            "email_category": raw_category,
            "email_category_confidence": raw_confidence,
        }

        # ── FINAL: what's going into the JSON ────────────────────
        _trace(
            "PIPELINE FINAL",
            f"Record for JSON — extracted_links has {len(record['extracted_links'])} entries, "
            f"tracking_link={record['tracking_link']!r}, "
            f"company={record['company']!r}, "
            f"order_number={record['order_number']!r}, "
            f"category={record['email_category']!r}",
        )
        _trace("PIPELINE", f"END process_file() — {file_path.name}")
        _trace("PIPELINE", "=" * 55)

        return record

    except Exception as e:
        _trace("PIPELINE", f"EXCEPTION in process_file(): {e}")
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
            "tracking_number": None,
            "tracking_link": None,
            "extracted_links": extracted_hrefs,
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
def main(flow_started_at: datetime | None = None):
    args = parse_args()

    _trace(
        "MAIN",
        f"main() called — base_dir={os.getenv(BASE_DIR_ENV)!r}, file={args.file!r}, "
        f"subject={args.subject!r}, sender_name={args.sender_name!r}, "
        f"email={args.email!r}",
    )

    if os.getenv("DEMO_MODE") == "1":
        args.email = "johndoe123@gmail.com"
        args.sender_name = "John Doe"
        print("DEMO MODE: sender overridden to John Doe <johndoe123@gmail.com>")

    base_dir = os.getenv(BASE_DIR_ENV)
    if not base_dir:
        raise ValueError(
            f"{BASE_DIR_ENV} is not set. Add it to python_files/.env "
            f"(e.g. BASE_DIR=C:\\path\\to\\project)."
        )

    base = Path(base_dir).expanduser().resolve()
    html_folder = base / "email_contents" / "html"
    output_path = base / "email_contents" / "json" / "results.json"

    output_path.parent.mkdir(parents=True, exist_ok=True)

    started = flow_started_at or datetime.now()
    try:
        init_flow_usage_log(base, started)
    except OSError as e:
        print(f"WARNING: Could not create OpenAI usage log file: {e}")

    if not API_KEY:
        print(
            f"WARNING: {OPENAI_API_KEY_ENV} is not set. "
            "Structured extraction will be skipped (empty fields only)."
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
        _trace("MAIN", f"Renamed to {file_path.name}, file size = {file_path.stat().st_size:,} bytes")

        entry = process_file(file_path, args.subject, args.sender_name, args.email)
        entry["content_hash"] = file_hash
        entry["duplicate_on_last_run"] = 0

        _trace(
            "MAIN",
            f"Writing JSON — extracted_links={len(entry.get('extracted_links', []))} entries, "
            f"tracking_link={entry.get('tracking_link')!r}",
        )

        results.append(entry)

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        _trace("MAIN", f"JSON written to {output_path} — {len(results)} total records")
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
    strip_bom_from_argv(sys.argv)

    _base_for_log = os.getenv(BASE_DIR_ENV)
    if not _base_for_log:
        print(
            f"ERROR: {BASE_DIR_ENV} is not set. Set it in {_PYTHON_FILES_DIR / '.env'}",
            file=sys.stderr,
        )
        sys.exit(1)
    _log_path = Path(_base_for_log).expanduser().resolve() / "programFileOutput.txt"
    _tee = _Tee(_log_path, sys.stdout)
    sys.stdout = _tee
    sys.stderr = _Tee(_log_path, sys.stderr)

    _start_time = time.time()
    _flow_started_at = datetime.now()

    print(f"\n{'='*60}")
    print(f"Run started: {_flow_started_at.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Args: {sys.argv[1:]}")
    print(f"{'='*60}")

    _original_stdout = _tee._original
    _original_stderr = sys.stderr._original

    try:
        main(flow_started_at=_flow_started_at)
        _elapsed = time.time() - _start_time
        print(f"Run finished successfully. Total operation time: {_elapsed:.2f}s")
    except SystemExit as e:
        _elapsed = time.time() - _start_time
        print(f"Total operation time: {_elapsed:.2f}s")
        if e.code == EXIT_BAD_ARGS:
            print("\nERROR: Invalid or missing arguments.")
            print("Set BASE_DIR and OPENAI_API_KEY in python_files/.env.")
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
