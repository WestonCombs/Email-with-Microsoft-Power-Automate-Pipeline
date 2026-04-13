import argparse
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import time
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

# python_files/ — .env must load before htmlHandler (trace uses BASE_DIR from .env)
_PYTHON_FILES_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_PYTHON_FILES_DIR))
# Same folder as this script (sibling modules: isGiftCard, grabTrackingLinks, …)
sys.path.insert(0, str(Path(__file__).resolve().parent))

from dotenv import load_dotenv

load_dotenv(_PYTHON_FILES_DIR / ".env")

from openai import OpenAI, RateLimitError

from htmlHandler.convertHTMLToPlaintext import convert as html_to_plaintext
import time as _time

import runLogger as RL
from htmlHandler.tracking_hrefs import (
    extract_hrefs_from_html,
    href_final_pairs,
    list_tracking_links_from_pairs,
    normalize_href_for_http_fetch,
    url_classifies_as_tracking,
)

try:
    from xhtml2pdf import pisa
    _HAS_XHTML2PDF = True
except ImportError:
    _HAS_XHTML2PDF = False

from isGiftCard import UNKNOWN as IS_GIFT_CARD_UNKNOWN, is_gift_card, should_run_is_gift_card


def _openai_fields_log_line(extracted: dict) -> str:
    """Short field summary for logs (not full JSON / body text)."""
    keys = (
        "company",
        "order_number",
        "purchase_datetime",
        "email_category",
        "tracking_number",
        "total_amount_paid",
        "tax_paid",
    )
    parts: list[str] = []
    for k in keys:
        v = extracted.get(k)
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        s = re.sub(r"\s+", " ", s)
        if len(s) > 100:
            s = s[:97] + "..."
        parts.append(f"{k}={s!r}")
    return "OpenAI — " + (", ".join(parts) if parts else "no structured fields")


# =========================
# CONFIG
# =========================
OPENAI_API_KEY_ENV = "OPENAI_API_KEY"
BASE_DIR_ENV = "BASE_DIR"
API_KEY = os.getenv(OPENAI_API_KEY_ENV)

MODEL = "gpt-4o-mini"

VALID_CATEGORIES = [
    "Invoice",
    "Shipped",
    "Delivered",
    "Gift Card",
    "Unknown",
]
# Categories the LLM returns directly (Gift Card is set only after the invoice refine call).
LLM_EMAIL_CATEGORIES = frozenset({"Invoice", "Shipped", "Delivered", "Unknown"})
CATEGORY_CONFIDENCE_THRESHOLD = 60

RATE_LIMIT_RETRY_WAIT = 3
RATE_LIMIT_MAX_RETRIES = 20
RATE_LIMIT_THROTTLE_THRESHOLD = 0.60   # trigger proactive cooldown at 60% used
RATE_LIMIT_COOLDOWN_CAP = 10           # max seconds to sleep per cooldown iteration

# Exit codes
EXIT_SUCCESS = 0
EXIT_ERROR = 1
EXIT_BAD_ARGS = 2
# OpenAI returned 429 on every attempt until retries exhausted (quota / rate limit).
# mainRunner.py must use the same numeric code when handling the subprocess.
EXIT_OPENAI_RATE_LIMIT_FATAL = 3


class OpenAIRateLimitFatalError(Exception):
    """Raised when every OpenAI chat completion attempt failed with RateLimitError."""

    pass

client = OpenAI(api_key=API_KEY)

OPENAI_USAGE_REL = Path("logs") / "openai usage"

# gpt-4o-mini pricing (per token)
_COST_PER_INPUT_TOKEN = 0.15 / 1_000_000   # $0.15 per 1M input tokens
_COST_PER_OUTPUT_TOKEN = 0.60 / 1_000_000  # $0.60 per 1M output tokens

# Set once per process run when main() initializes the flow usage log (CLI entry).
_flow_usage_log_path: Path | None = None


def _write_tracking_log(
    file_name: str,
    subject: str | None,
    sender_name: str | None,
    sender_email: str | None,
    href_pairs: list[tuple[str, str]],
    tracking_links: list[str],
) -> None:
    """Write per-email href/tracking section to logs/tracking_hrefs.txt (accumulates)."""
    redirected = sum(1 for h, f in href_pairs if h.strip() != f.strip())
    tracking_count = sum(1 for _, f in href_pairs if url_classifies_as_tracking(f))
    ts = RL.ts()

    # Normal log: resolved finals with verdict
    lines: list[str] = [
        f"\n{'-' * 72}",
        f"{ts}  |  {file_name}",
        f'  subject: "{(subject or "")[:70]}"',
        f"  sender:  {sender_name} <{sender_email}>",
        f"  hrefs: {len(href_pairs)} unique  |  redirected: {redirected}  |  tracking candidates: {tracking_count}",
        "",
    ]
    if not href_pairs:
        lines.append("  (no hrefs extracted)")
    else:
        for i, (href, final) in enumerate(href_pairs, 1):
            verdict = "TRACKING    " if url_classifies_as_tracking(final) else "not-tracking"
            lines.append(f"  {i:3}. [{verdict}]  {final}")
    if not tracking_links:
        lines.append("\n  pick_tracking_link: none found")
    elif len(tracking_links) == 1:
        lines.append(f"\n  pick_tracking_link: {tracking_links[0]}")
    else:
        lines.append(f"\n  pick_tracking_link: {len(tracking_links)} distinct tracking URLs:")
        for i, u in enumerate(tracking_links, 1):
            lines.append(f"    {i}. {u}")
    lines.append("")
    RL.log("tracking_hrefs", "\n".join(lines))

    # Debug log: also shows original href when it differs from final
    debug_lines: list[str] = [
        f"\n{ts}  |  {file_name}  [debug]",
        "  href  →  final (redirect chain):",
    ]
    for i, (href, final) in enumerate(href_pairs, 1):
        verdict = "TRACKING    " if url_classifies_as_tracking(final) else "not-tracking"
        _, fetchable = normalize_href_for_http_fetch(href)
        skipped = "" if fetchable else " [non-http, not fetched]"
        if href.strip() != final.strip():
            debug_lines.append(f"  {i:3}. [{verdict}]  {href}")
            debug_lines.append(f"            →  {final}")
        else:
            debug_lines.append(f"  {i:3}. [{verdict}]  {final}{skipped}")
    debug_lines.append("")
    RL.debug("tracking_hrefs", "\n".join(debug_lines))


def _write_openai_log(
    file_name: str,
    subject: str | None,
    sender_name: str | None,
    sender_email: str | None,
    extracted: dict,
    final_category: str,
    confidence: int,
    gift_verdict,
    timings: dict,
) -> None:
    """Write per-email OpenAI extraction result to logs/openai_extraction.txt."""
    ts = RL.ts()
    gv = (
        "gift card" if gift_verdict is True
        else "items invoice" if gift_verdict is False
        else "n/a"
    )
    lines: list[str] = [
        f"\n{ts}  |  {file_name}",
        f'  "{(subject or "")[:70]}"  —  {sender_name} <{sender_email}>',
        f"  Category: {final_category} (conf={confidence}) | Gift card check: {gv}",
        f"  company={extracted.get('company') or 'n/a'}  |  "
        f"order={extracted.get('order_number') or 'n/a'}  |  "
        f"date={extracted.get('purchase_datetime') or 'n/a'}  |  "
        f"amount={extracted.get('total_amount_paid') or 'n/a'}",
        f"  tracking_number={extracted.get('tracking_number') or 'n/a'}",
        f"  OpenAI: {'ran' if timings.get('step5_ran') else 'skipped'}  "
        f"{timings.get('step5_s', 0.0):.2f}s  |  "
        f"GiftCard check: {'ran' if timings.get('step5b_ran') else 'skipped'}  "
        f"{timings.get('step5b_s', 0.0):.2f}s",
        "",
    ]
    RL.log("openai_extraction", "\n".join(lines))
    RL.debug("openai_extraction",
        f"\n{ts}  |  {file_name}  [debug]\n"
        f"  raw extracted: company={extracted.get('company')!r}, "
        f"order={extracted.get('order_number')!r}, "
        f"date={extracted.get('purchase_datetime')!r}, "
        f"amount={extracted.get('total_amount_paid')!r}, "
        f"tracking_number={extracted.get('tracking_number')!r}, "
        f"category_raw={extracted.get('email_category')!r}, "
        f"confidence_raw={extracted.get('email_category_confidence')!r}\n"
    )


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
    elapsed_secs: float = 0.0,
) -> int:
    """Append one line for this flow's file; cumulative is for this flow only. Returns cumulative."""
    global _flow_usage_log_path
    if _flow_usage_log_path is None:
        return 0
    log_path = _flow_usage_log_path

    prev = _read_last_cumulative_tokens(log_path)
    cumulative = prev + total_tokens
    cost = (prompt_tokens * _COST_PER_INPUT_TOKEN) + (completion_tokens * _COST_PER_OUTPUT_TOKEN)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    line = (
        f"{ts} | "
        f"prompt_tokens={prompt_tokens} completion_tokens={completion_tokens} "
        f"total_tokens={total_tokens} | "
        f"cumulative_total={cumulative} | "
        f"elapsed_secs={elapsed_secs:.2f} | "
        f"cost=${cost:.6f}\n"
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


def _find_browser() -> Path | None:
    """Find Edge or Chrome for headless PDF conversion (Edge ships with Win 10+)."""
    candidates = []
    for env_var in ("PROGRAMFILES(X86)", "PROGRAMFILES"):
        base = os.environ.get(env_var, "")
        if not base:
            continue
        candidates.append(Path(base) / "Microsoft" / "Edge" / "Application" / "msedge.exe")
        candidates.append(Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe")
    for p in candidates:
        if p.exists():
            return p
    return None


def convert_html_to_pdf(html_path: Path) -> Path:
    """Convert an HTML file to PDF. Tries Edge/Chrome headless first (perfect
    rendering), then xhtml2pdf as fallback. Returns the original path if all
    methods fail."""
    pdf_path = html_path.with_suffix(".pdf")
    file_uri = html_path.resolve().as_uri()

    # --- Strategy 1: Edge / Chrome headless (handles any email HTML) ---
    browser = _find_browser()
    if browser:
        try:
            result = subprocess.run(
                [
                    str(browser),
                    "--headless",
                    "--disable-gpu",
                    "--no-pdf-header-footer",
                    f"--print-to-pdf={pdf_path}",
                    file_uri,
                ],
                capture_output=True,
                timeout=30,
            )
            if pdf_path.exists() and pdf_path.stat().st_size > 0:
                html_path.unlink()
                print(f"  Converted to PDF: {pdf_path.name}")
                return pdf_path
        except Exception as e:
            print(f"  Browser PDF conversion failed: {e}")
            if pdf_path.exists():
                try:
                    pdf_path.unlink()
                except OSError:
                    pass

    # --- Strategy 2: xhtml2pdf (pure Python fallback) ---
    if _HAS_XHTML2PDF:
        try:
            html_content, _ = read_email_html_file(html_path)
            with open(pdf_path, "wb") as pdf_file:
                pisa.CreatePDF(
                    io.StringIO(html_content),
                    dest=pdf_file,
                    encoding="utf-8",
                )
            if pdf_path.exists() and pdf_path.stat().st_size > 0:
                html_path.unlink()
                print(f"  Converted to PDF (xhtml2pdf): {pdf_path.name}")
                return pdf_path
        except Exception as e:
            print(f"  xhtml2pdf conversion failed: {e}")

    # --- Both failed ---
    print(f"  WARNING: Could not convert {html_path.name} to PDF, keeping HTML")
    if pdf_path.exists():
        try:
            pdf_path.unlink()
        except OSError:
            pass
    return html_path


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


def _openai_rate_limit_debug(exc: RateLimitError) -> tuple[str, str]:
    """Build (short one-line summary, full multi-line block) for 429 responses."""
    lines: list[str] = [
        f"message: {exc.message}",
        f"status_code: {exc.status_code}",
    ]
    rid = getattr(exc, "request_id", None)
    if rid:
        lines.append(f"x-request-id: {rid}")

    body = getattr(exc, "body", None)
    if body is not None:
        try:
            lines.append("body:\n" + json.dumps(body, ensure_ascii=False, indent=2))
        except (TypeError, ValueError):
            lines.append(f"body: {body!r}")

    resp = getattr(exc, "response", None)
    if resp is not None:
        h = resp.headers
        for key in (
            "retry-after",
            "x-ratelimit-limit-requests",
            "x-ratelimit-remaining-requests",
            "x-ratelimit-reset-requests",
            "x-ratelimit-limit-tokens",
            "x-ratelimit-remaining-tokens",
            "x-ratelimit-reset-tokens",
        ):
            val = h.get(key)
            if val:
                lines.append(f"{key}: {val}")

    full = "\n".join(lines)

    short_parts: list[str] = []
    if isinstance(body, dict):
        err = body.get("error")
        if isinstance(err, dict):
            for k in ("code", "type"):
                v = err.get(k)
                if v:
                    short_parts.append(f"{k}={v}")
            em = err.get("message")
            if isinstance(em, str) and em.strip():
                short_parts.append(em.strip()[:200])
    if resp is not None:
        ra = resp.headers.get("retry-after")
        if ra:
            short_parts.append(f"retry-after={ra}s")
    if rid:
        short_parts.append(f"request_id={rid}")
    short = "  |  ".join(short_parts) if short_parts else exc.message[:240]
    return short, full


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
def _sanitize_for_api(text: str) -> str:
    """Strip control characters (except newline/tab) that break JSON serialization."""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)


def _chat_completion_json_parsed(api_kwargs: dict) -> dict:
    """Run chat completion with JSON response; rate-limit retry and usage logging."""
    total_waited = 0.0
    call_start = time.monotonic()
    for attempt in range(1, RATE_LIMIT_MAX_RETRIES + 1):
        try:
            raw_resp = client.chat.completions.with_raw_response.create(**api_kwargs)
            response = raw_resp.parse()
            total_waited += _check_and_throttle(raw_resp.headers)
            break
        except RateLimitError as e:
            print(f"  Rate limit hit (attempt {attempt}/{RATE_LIMIT_MAX_RETRIES}) — waiting {RATE_LIMIT_RETRY_WAIT}s...")
            if RL.is_debug():
                short, full = _openai_rate_limit_debug(e)
                print(f"  {short}")
                RL.debug(
                    "grabbingImportantEmailContent",
                    f"OpenAI RateLimitError (attempt {attempt}/{RATE_LIMIT_MAX_RETRIES}):\n{full}\n",
                )
            time.sleep(RATE_LIMIT_RETRY_WAIT)
            total_waited += RATE_LIMIT_RETRY_WAIT
    else:
        raise OpenAIRateLimitFatalError(
            f"OpenAI rate limit not cleared after {RATE_LIMIT_MAX_RETRIES} retries "
            f"({total_waited:.0f}s total wait)"
        )
    call_elapsed = time.monotonic() - call_start

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
                elapsed_secs=call_elapsed,
            )
            print(
                f"  OpenAI tokens this request: {tt} (cumulative this flow: {cumulative})"
            )
        except OSError as e:
            print(f"  WARNING: Could not write OpenAI usage log: {e}")

    content = response.choices[0].message.content
    return json.loads(content)


def resolve_base_email_category(extracted: dict) -> tuple[str, float]:
    """Map first-pass LLM output to category (Invoice / Shipped / Delivered / Unknown).

    Gift Card is not produced here; it is applied later only after is_gift_card().
    """
    raw = extracted.get("email_category", "Unknown")
    if raw not in LLM_EMAIL_CATEGORIES:
        raw = "Unknown"
    try:
        conf = float(extracted.get("email_category_confidence", 0))
    except (TypeError, ValueError):
        conf = 0.0

    if conf < CATEGORY_CONFIDENCE_THRESHOLD:
        return ("Unknown", conf)

    return (raw, conf)


def extract_with_openai(
    text_only: str,
    subject: str | None = None,
) -> dict:
    """Extract purchase details using OpenAI (plain text only, no HTML)."""
    text_only = _sanitize_for_api(text_only)
    if subject:
        subject = _sanitize_for_api(subject)
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
12. email_category: Classify into exactly ONE of these categories (do NOT use "Gift Card" here):
    - "Invoice": Order placed, confirmed, or receipt for a purchase (merchandise, services, or a gift card purchase — any order/receipt email).
    - "Shipped": The package has been shipped or is in transit (carrier handoff, tracking, out for delivery).
    - "Delivered": The package has arrived — delivery is complete.
    - "Unknown": Does not clearly fit Invoice, Shipped, or Delivered.
13. email_category_confidence: Your confidence (0–100) in email_category.
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
                                "Invoice",
                                "Shipped",
                                "Delivered",
                                "Unknown",
                            ],
                            "description": (
                                "Invoice = any order confirmation or purchase receipt. "
                                "Do not use Gift Card here."
                            ),
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

    return _chat_completion_json_parsed(api_kwargs)


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
            "A bare filename is resolved under email_contents/pdf/ relative to BASE_DIR from .env "
            "(e.g. file1.html). When provided, only that file is processed and the result is appended "
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
    """Run the full extraction pipeline for one HTML email file.

    Returns the record dict. Always includes a ``_timings`` key that ``main()``
    pops before writing to JSON.
    """
    t_overall = _time.perf_counter()

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
    timings: dict = {}

    try:
        # ── STEP 1: Read HTML from disk ──────────────────────────
        t = _time.perf_counter()
        print("  » Reading email HTML...", end=" ", flush=True)
        raw_html, html_encoding = read_email_html_file(file_path)
        href_count_raw = raw_html.lower().count('href=')
        timings["step1_s"] = round(_time.perf_counter() - t, 3)
        timings["html_chars"] = len(raw_html)
        print(f"done  ({timings['step1_s']:.2f}s, {len(raw_html):,} chars)")
        if len(raw_html) < 500:
            print(f"  WARNING: HTML is very short ({len(raw_html)} chars) — check source email.")
        RL.debug("grabbingImportantEmailContent",
            f"  [step1] {file_path.name}: {len(raw_html):,} chars ({html_encoding}), "
            f"href_tokens={href_count_raw}, anchor_tags={raw_html.lower().count('<a ')}"
        )
        RL.debug("htmlHandler",
            f"  {RL.ts()}  {file_path.name}: html_in={len(raw_html):,} chars"
        )

        # ── STEP 2: HTML → plain text ────────────────────────────
        t = _time.perf_counter()
        print("  » Converting to plain text...", end=" ", flush=True)
        text_only = html_to_plaintext(raw_html)
        timings["step2_s"] = round(_time.perf_counter() - t, 3)
        timings["text_chars"] = len(text_only)
        print(f"done  ({timings['step2_s']:.2f}s, {len(text_only):,} chars)")
        RL.debug("grabbingImportantEmailContent",
            f"  [step2] plaintext: {len(text_only):,} chars, {len(text_only.splitlines())} lines"
        )
        RL.debug("htmlHandler",
            f"         text_out={len(text_only):,} chars, "
            f"truncated={'yes' if len(text_only) >= 50_000 else 'no'}"
        )
        RL.log("htmlHandler",
            f"{RL.ts()}  {file_path.name}:  "
            f"html={len(raw_html):,}  text={len(text_only):,}  "
            f"ratio={len(text_only)/max(len(raw_html),1):.0%}"
        )

        # ── STEP 3: Extract hrefs ────────────────────────────────
        t = _time.perf_counter()
        print("  » Extracting hrefs...", end=" ", flush=True)
        extracted_hrefs = extract_hrefs_from_html(raw_html)
        fetchable = sum(1 for h in extracted_hrefs if normalize_href_for_http_fetch(h)[1])
        timings["step3_s"] = round(_time.perf_counter() - t, 3)
        timings["hrefs_found"] = len(extracted_hrefs)
        timings["hrefs_fetchable"] = fetchable
        print(f"done  ({timings['step3_s']:.2f}s, {len(extracted_hrefs)} unique, {fetchable} http/https)")
        if not extracted_hrefs:
            print(f"  WARNING: 0 hrefs extracted (HTML had {href_count_raw} href= tokens)")
        RL.debug("grabbingImportantEmailContent",
            f"  [step3] hrefs: found={len(extracted_hrefs)}, http/https={fetchable}, "
            f"non-web={len(extracted_hrefs)-fetchable}"
        )

        # ── STEP 4: Resolve redirects + pick tracking link ───────
        # Uses htmlHandler.tracking_hrefs: href_final_pairs → list_tracking_links_from_pairs
        # (BeautifulSoup hrefs, redirect resolution, tracking heuristics, dedupe).
        t = _time.perf_counter()
        print("  » Resolving redirects & classifying tracking links...", end=" ", flush=True)
        href_pairs = href_final_pairs(extracted_hrefs)
        tracking_links = list_tracking_links_from_pairs(href_pairs)
        # Primary URL for legacy single-field consumers; full list is ``tracking_links``.
        tracking_link = tracking_links[0] if tracking_links else None
        redirected = sum(1 for h, f in href_pairs if h.strip() != f.strip())
        tracking_cands = sum(1 for _, f in href_pairs if url_classifies_as_tracking(f))
        timings["step4_s"] = round(_time.perf_counter() - t, 3)
        timings["hrefs_redirected"] = redirected
        timings["tracking_candidates"] = tracking_cands
        n_track = len(tracking_links)
        timings["tracking_result"] = (
            "none" if n_track == 0 else ("multiple" if n_track > 1 else "single")
        )
        ts_label = (
            "none found" if n_track == 0
            else (f"{n_track} tracking links" if n_track > 1 else "1 link found")
        )
        print(f"done  ({timings['step4_s']:.2f}s, {redirected} redirected, {ts_label})")
        _write_tracking_log(file_path.name, subject, sender_name, email, href_pairs, tracking_links)
        RL.debug("grabbingImportantEmailContent",
            f"  [step4] tracking: candidates={tracking_cands}, count={n_track}, links={tracking_links!r}"
        )

        # ── STEP 5: OpenAI extraction ────────────────────────────
        extracted = None
        t5 = _time.perf_counter()
        timings["step5_ran"] = bool(API_KEY)

        if API_KEY:
            print("  » OpenAI extraction...", end=" ", flush=True)
            try:
                extracted = extract_with_openai(text_only, subject=subject)
                timings["step5_s"] = round(_time.perf_counter() - t5, 3)
                print(f"done  ({timings['step5_s']:.2f}s)")
                RL.debug("grabbingImportantEmailContent",
                    f"  [step5] {_openai_fields_log_line(extracted)}"
                )
            except OpenAIRateLimitFatalError:
                timings["step5_s"] = round(_time.perf_counter() - t5, 3)
                print(f"FAILED  ({timings['step5_s']:.2f}s)")
                raise
            except Exception as e:
                timings["step5_s"] = round(_time.perf_counter() - t5, 3)
                print(f"FAILED  ({timings['step5_s']:.2f}s, {e})")
                extracted = None
        else:
            timings["step5_s"] = 0.0
            print("  » OpenAI extraction...  skipped (no API key set)")

        if extracted is None:
            if not API_KEY:
                print("  WARNING: No API key — empty fields only.")
            extracted = empty_extraction

        if not clean_text(extracted.get("company")):
            extracted["company"] = infer_company_from_subject(subject)

        file_uri = "file:///" + str(file_path.resolve()).replace("\\", "/")
        final_category, raw_confidence = resolve_base_email_category(extracted)

        # ── STEP 5b: Gift card check ─────────────────────────────
        gift_verdict: bool | int | None = None
        t5b = _time.perf_counter()
        timings["step5b_ran"] = False

        if API_KEY and should_run_is_gift_card(extracted):
            timings["step5b_ran"] = True
            print("  » Gift card check...", end=" ", flush=True)
            try:
                gift_verdict = is_gift_card(text_only, subject=subject)
                timings["step5b_s"] = round(_time.perf_counter() - t5b, 3)
                gv_label = (
                    "gift card" if gift_verdict is True
                    else ("items invoice" if gift_verdict is False else "inconclusive")
                )
                print(f"done  ({timings['step5b_s']:.2f}s, {gv_label})")
            except OpenAIRateLimitFatalError:
                timings["step5b_s"] = round(_time.perf_counter() - t5b, 3)
                print(f"FAILED  ({timings['step5b_s']:.2f}s)")
                raise
            except Exception as e:
                timings["step5b_s"] = round(_time.perf_counter() - t5b, 3)
                print(f"FAILED  ({timings['step5b_s']:.2f}s, {e})")
                gift_verdict = None
        else:
            timings["step5b_s"] = 0.0

        if gift_verdict is not None:
            if gift_verdict is True:
                final_category = "Gift Card"
            elif gift_verdict is False:
                final_category = "Invoice"
            elif gift_verdict == IS_GIFT_CARD_UNKNOWN:
                final_category = "Invoice"

        if final_category not in VALID_CATEGORIES:
            final_category = "Unknown"

        timings["total_s"] = round(_time.perf_counter() - t_overall, 3)
        timings["category"] = final_category
        timings["category_confidence"] = raw_confidence

        # Write openai_extraction log
        _write_openai_log(
            file_path.name, subject, sender_name, email,
            extracted, final_category, raw_confidence, gift_verdict, timings,
        )

        # Write one-line summary to grabbingImportantEmailContent log
        RL.log("grabbingImportantEmailContent",
            f"{RL.ts()}  {file_path.name}  |  "
            f"\"{(subject or '')[:50]}\"  |  "
            f"{final_category} (conf={raw_confidence})  |  "
            f"order={extracted.get('order_number') or 'n/a'}  |  "
            f"total={timings['total_s']:.2f}s"
        )
        RL.debug("grabbingImportantEmailContent",
            f"  [final] category={final_category}, confidence={raw_confidence}, "
            f"company={extracted.get('company')!r}, order={extracted.get('order_number')!r}, "
            f"amount={extracted.get('total_amount_paid')!r}, "
            f"tracking_links={tracking_links!r}\n"
        )

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
            "tracking_links": tracking_links,
            "extracted_links": extracted_hrefs,
            "email_category": final_category,
            "email_category_confidence": raw_confidence,
            "_timings": timings,
        }
        for idx, u in enumerate(tracking_links, 1):
            record[f"trackingNumber{idx}"] = u
        return record

    except Exception as e:
        timings["total_s"] = round(_time.perf_counter() - t_overall, 3)
        timings["error"] = str(e)
        print(f"  ERROR in pipeline: {e}")
        RL.log("grabbingImportantEmailContent",
            f"{RL.ts()}  {file_path.name}  |  ERROR: {e}"
        )
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
            "tracking_links": [],
            "extracted_links": extracted_hrefs,
            "email_category": "Unknown",
            "email_category_confidence": 0,
            "_timings": timings,
        }


# =========================
# FILE RENAMING
# =========================
_SEQUENTIAL_HTML = re.compile(r"^file(\d+)\.(?:html|htm)$", re.IGNORECASE)


def _next_file_number(html_folder: Path) -> int:
    """Find the highest existing fileN.html/.pdf number (recursive) and return N+1."""
    pattern = re.compile(r"^file(\d+)\.(?:html?|pdf)$", re.IGNORECASE)
    max_n = 0
    for p in html_folder.rglob("file*"):
        if p.is_file():
            m = pattern.match(p.name)
            if m:
                max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def rename_single_file(file_path: Path, html_folder: Path) -> Path:
    """Assign ``fileN.html`` to a drop-in HTML file.

    If the basename is already ``fileN.html`` / ``fileN.htm`` (e.g. from mainRunner),
    return it unchanged. Otherwise rename to the next free ``fileN.html``.
    """
    if _SEQUENTIAL_HTML.match(file_path.name):
        return file_path
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
# CONVENTION FILE NAMING
# =========================
_CATEGORY_SUFFIX_MAP = {
    "Invoice":   "INVOICE",
    "Shipped":   "SHIPPED",
    "Delivered": "DELIVERED",
    "Gift Card": None,
}


def _sanitize_for_filename(name: str) -> str:
    """Remove characters invalid in Windows filenames and collapse whitespace."""
    sanitized = re.sub(r'[<>:"/\\|?*]', '', name)
    sanitized = re.sub(r'\s+', ' ', sanitized).strip('. ')
    return sanitized or "Unknown"


def _extract_date_str(purchase_datetime: str | None) -> str:
    """Pull a YYYY-MM-DD date from the extracted purchase_datetime, or today's date."""
    if purchase_datetime:
        m = re.match(r"(\d{4}-\d{2}-\d{2})", purchase_datetime.strip())
        if m:
            return m.group(1)
    return datetime.now().strftime("%Y-%m-%d")


def build_convention_filename(record: dict, extension: str = ".pdf") -> str:
    """Build a filename following the mom's naming convention.

    Invoice    → DOC <store> <YYYY-MM-DD> INVOICE
    Shipped    → DOC <store> <YYYY-MM-DD> SHIPPED
    Delivered  → DOC <store> <YYYY-MM-DD> DELIVERED
    Gift Card  → <store> <YYYY-MM-DD>
    Unknown    → DOC <store> <YYYY-MM-DD>
    """
    category = record.get("email_category", "Unknown")
    store = _sanitize_for_filename(record.get("company") or "Unknown")
    date_str = _extract_date_str(record.get("purchase_datetime"))

    suffix = _CATEGORY_SUFFIX_MAP.get(category)

    if category == "Gift Card":
        name = f"{store} {date_str}"
    elif suffix:
        name = f"DOC {store} {date_str} {suffix}"
    else:
        name = f"DOC {store} {date_str}"

    return name + extension


def rename_to_convention(file_path: Path, record: dict, target_folder: Path) -> Path:
    """Rename *file_path* to the convention name inside *target_folder*.
    Appends (2), (3), … if the name already exists."""
    ext = file_path.suffix
    base_name = build_convention_filename(record, extension="")
    new_path = target_folder / f"{base_name}{ext}"

    counter = 2
    while new_path.exists():
        new_path = target_folder / f"{base_name} ({counter}){ext}"
        counter += 1

    file_path.rename(new_path)
    print(f"  Convention rename: {file_path.name} -> {new_path.name}")
    return new_path


def rebuild_email_html_archive_folder(html_dir: Path) -> None:
    """Empty *html_dir* (create if missing). Call once per program run before saving archived HTML."""
    html_dir.mkdir(parents=True, exist_ok=True)
    for child in list(html_dir.iterdir()):
        try:
            if child.is_file():
                child.unlink()
            elif child.is_dir():
                shutil.rmtree(child)
        except OSError as e:
            print(f"WARNING: could not remove {child}: {e}")


def archive_html_before_pdf(source_html: Path, record: dict, html_folder: Path) -> Path:
    """Copy *source_html* into *html_folder* using the same convention basename as the PDF (``.html``).

    Must run **before** :func:`convert_html_to_pdf`, which may delete *source_html*.
    """
    html_folder.mkdir(parents=True, exist_ok=True)
    ext = ".html"
    base_name = build_convention_filename(record, extension="")
    new_path = html_folder / f"{base_name}{ext}"
    counter = 2
    while new_path.exists():
        new_path = html_folder / f"{base_name} ({counter}){ext}"
        counter += 1
    shutil.copy2(source_html, new_path)
    print(f"  Archived HTML: {new_path.name}")
    return new_path


# =========================
# COMPANY NAME CONSENSUS (per order_number)
# =========================


def _normalized_order_key(record: dict) -> str:
    """Stable order id for grouping rows (same as Excel / JSON order_number)."""
    v = record.get("order_number")
    if v is None:
        return ""
    s = str(v).replace("\ufeff", "").strip()
    return s


def _company_vote_key(company: str | None) -> str:
    """Normalize company for plurality voting (case, spacing, & vs 'and')."""
    c = clean_text(company)
    if not c:
        return ""
    s = c.casefold()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("&", " and ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def unify_company_names_by_order(results: list[dict]) -> None:
    """For each order_number shared by 2+ rows, set ``company`` to the plurality winner.

    Voting uses a normalized key so variants like 'Bath & Body Works' and
    'bath and body works' count together; the displayed string is the most
    frequent *original* spelling among rows in the winning bucket.
    """
    groups: dict[str, list[dict]] = defaultdict(list)
    for r in results:
        ok = _normalized_order_key(r)
        if ok:
            groups[ok].append(r)

    for order_key, group in groups.items():
        if len(group) < 2:
            continue

        key_votes: Counter[str] = Counter()
        originals_by_vote_key: dict[str, list[str]] = defaultdict(list)

        for r in group:
            raw = clean_text(r.get("company"))
            if not raw:
                continue
            vk = _company_vote_key(raw)
            if not vk:
                continue
            key_votes[vk] += 1
            originals_by_vote_key[vk].append(raw)

        if not key_votes:
            continue

        winning_vote_key = sorted(
            key_votes.items(),
            key=lambda kv: (
                -kv[1],
                -max((len(x) for x in originals_by_vote_key[kv[0]]), default=0),
                kv[0],
            ),
        )[0][0]

        origs = originals_by_vote_key[winning_vote_key]
        oc = Counter(origs)
        winner_display = sorted(
            oc.items(),
            key=lambda kv: (-kv[1], -len(kv[0]), kv[0]),
        )[0][0]

        before_vals = [clean_text(r.get("company")) for r in group]
        for r in group:
            r["company"] = winner_display

        if any(b != winner_display for b in before_vals):
            print(
                f"  Company consensus (order {order_key}): {winner_display!r} "
                f"— {key_votes[winning_vote_key]} vote(s) for winning label, "
                f"{len(group)} row(s) updated"
            )


def rename_assets_to_match_record(
    record: dict, pdf_folder: Path, html_folder: Path
) -> None:
    """Rename PDF and archived HTML so basenames match :func:`build_convention_filename`."""
    src = record.get("source_file")
    if not src:
        return
    old_pdf = Path(src)
    if not old_pdf.is_absolute():
        old_pdf = pdf_folder / old_pdf.name
    if not old_pdf.is_file():
        return

    ext = old_pdf.suffix
    want_name = build_convention_filename(record, ext)
    if old_pdf.name == want_name:
        return

    new_pdf = pdf_folder / want_name
    counter = 2
    while new_pdf.exists() and new_pdf.resolve() != old_pdf.resolve():
        stem = Path(want_name).stem
        new_pdf = pdf_folder / f"{stem} ({counter}){ext}"
        counter += 1

    old_stem = old_pdf.stem
    old_pdf.rename(new_pdf)

    old_html = html_folder / f"{old_stem}.html"
    if old_html.is_file():
        new_html = html_folder / f"{new_pdf.stem}.html"
        c2 = 2
        while new_html.exists() and new_html.resolve() != old_html.resolve():
            new_html = html_folder / f"{new_pdf.stem} ({c2}).html"
            c2 += 1
        old_html.rename(new_html)

    record["source_file"] = clean_text(new_pdf)
    record["source_file_link"] = (
        "file:///" + str(new_pdf.resolve()).replace("\\", "/")
    )


def apply_order_company_consensus_and_sync(
    results: list[dict], base_dir: Path
) -> None:
    """Update ``company`` by order-number plurality and rename PDF/HTML to match."""
    pdf_folder = base_dir / "email_contents" / "pdf"
    html_folder = base_dir / "email_contents" / "html"
    unify_company_names_by_order(results)
    for r in results:
        try:
            rename_assets_to_match_record(r, pdf_folder, html_folder)
        except OSError as e:
            print(f"  WARNING: could not sync filenames for {r.get('source_file')}: {e}")


def _write_results_with_consensus(
    output_path: Path, results: list[dict], base: Path
) -> None:
    apply_order_company_consensus_and_sync(results, base)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)


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


def _print_openai_fatal_banner() -> None:
    """Console message when OpenAI retries are exhausted (subprocess — no GUI here)."""
    print(
        "\n"
        + "=" * 60
        + "\n"
        "FATAL: OpenAI API — could not complete after all automatic retries.\n"
        "\n"
        "A moderator must fix the OPENAI_API_KEY and the OpenAI account it belongs to\n"
        "(billing, quota, and key/project scope at platform.openai.com).\n"
        "\n"
        "This run stops here. Remaining emails were not processed.\n"
        + "=" * 60
        + "\n",
        file=sys.stderr,
    )


# =========================
# MAIN
# =========================
def main(flow_started_at: datetime | None = None):
    args = parse_args()

    RL.trace(
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
    pdf_folder = base / "email_contents" / "pdf"
    html_archive_folder = base / "email_contents" / "html"
    output_path = base / "email_contents" / "json" / "results.json"

    output_path.parent.mkdir(parents=True, exist_ok=True)

    started = flow_started_at or datetime.now()
    ext_log = os.getenv("OPENAI_USAGE_LOG_PATH")
    if ext_log:
        global _flow_usage_log_path
        _flow_usage_log_path = Path(ext_log)
    else:
        try:
            init_flow_usage_log(base, started)
        except OSError as e:
            print(f"WARNING: Could not create OpenAI usage log file: {e}")

    if not API_KEY:
        print(
            f"WARNING: {OPENAI_API_KEY_ENV} is not set. "
            "Structured extraction will be skipped (empty fields only)."
        )

    _timing_buffer_path: Path | None = None
    _buf_raw = os.getenv("TIMING_BUFFER_PATH", "").strip()
    if _buf_raw:
        _timing_buffer_path = Path(_buf_raw)

    # Single-file mode
    if args.file:
        candidate = Path(args.file)
        if not candidate.is_absolute() and candidate.parent == Path("."):
            candidate = pdf_folder / candidate
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
            _write_results_with_consensus(output_path, results, base)
            return

        file_path = rename_single_file(file_path, pdf_folder)

        entry = process_file(file_path, args.subject, args.sender_name, args.email)
        timings = entry.pop("_timings", {})
        if _timing_buffer_path:
            RL.write_timing_entry(_timing_buffer_path, {
                "file": file_path.name,
                "subject": args.subject,
                "is_duplicate": False,
                **timings,
            })

        entry["content_hash"] = file_hash
        entry["duplicate_on_last_run"] = 0

        archive_html_before_pdf(file_path, entry, html_archive_folder)
        pdf_path = convert_html_to_pdf(file_path)
        pdf_path = rename_to_convention(pdf_path, entry, pdf_folder)
        file_uri = "file:///" + str(pdf_path.resolve()).replace("\\", "/")
        entry["source_file"] = clean_text(pdf_path)
        entry["source_file_link"] = file_uri

        results.append(entry)

        _write_results_with_consensus(output_path, results, base)

        print(f"  Result saved to: {output_path} ({len(results)} total records)")
        return

    # Batch mode: process entire pdf folder
    if not pdf_folder.exists():
        raise FileNotFoundError(f"PDF input folder not found: {pdf_folder}")

    print("Renaming HTML files to sequential format...")
    html_files = rename_html_files_sequential(pdf_folder)

    if not html_files:
        print(f"No HTML files found in: {pdf_folder}")
        return

    results = _load_existing_results(output_path)
    hashes = _known_hashes(results)
    new_count = 0

    rebuild_email_html_archive_folder(html_archive_folder)

    for fp in html_files:
        file_hash = compute_file_hash(fp)
        if file_hash in hashes:
            print(f"  » Duplicate — already in results, skipped: {fp.name}")
            if _timing_buffer_path:
                RL.write_timing_entry(_timing_buffer_path, {
                    "file": fp.name, "subject": args.subject, "is_duplicate": True, "total_s": 0.0
                })
            for r in results:
                if r.get("content_hash") == file_hash:
                    r["duplicate_on_last_run"] = 1
                    break
            continue
        entry = process_file(fp, args.subject, args.sender_name, args.email)
        timings = entry.pop("_timings", {})
        if _timing_buffer_path:
            RL.write_timing_entry(_timing_buffer_path, {
                "file": fp.name, "subject": args.subject, "is_duplicate": False, **timings
            })

        entry["content_hash"] = file_hash
        entry["duplicate_on_last_run"] = 0

        archive_html_before_pdf(fp, entry, html_archive_folder)
        pdf_path = convert_html_to_pdf(fp)
        pdf_path = rename_to_convention(pdf_path, entry, pdf_folder)
        file_uri = "file:///" + str(pdf_path.resolve()).replace("\\", "/")
        entry["source_file"] = clean_text(pdf_path)
        entry["source_file_link"] = file_uri

        results.append(entry)
        hashes.add(file_hash)
        new_count += 1

    _write_results_with_consensus(output_path, results, base)

    print(f"\n  Done. {new_count} new + {len(results) - new_count} existing = {len(results)} total records → {output_path}")


if __name__ == "__main__":
    strip_bom_from_argv(sys.argv)

    _base_for_log = os.getenv(BASE_DIR_ENV)
    if not _base_for_log:
        print(
            f"ERROR: {BASE_DIR_ENV} is not set. Set it in {_PYTHON_FILES_DIR / '.env'}",
            file=sys.stderr,
        )
        sys.exit(1)

    _start_time = time.time()
    _flow_started_at = datetime.now()

    print(f"\n{'='*60}")
    print(f"Run started: {_flow_started_at.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Args: {sys.argv[1:]}")
    print(f"{'='*60}")

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
        sys.exit(e.code)
    except OpenAIRateLimitFatalError as e:
        _elapsed = time.time() - _start_time
        _print_openai_fatal_banner()
        print(f"Detail: {e}", file=sys.stderr)
        print(f"Total operation time: {_elapsed:.2f}s")
        sys.exit(EXIT_OPENAI_RATE_LIMIT_FATAL)
    except Exception as e:
        _elapsed = time.time() - _start_time
        print(f"\nERROR: {e}")
        print(f"Total operation time: {_elapsed:.2f}s")
        sys.exit(1)
