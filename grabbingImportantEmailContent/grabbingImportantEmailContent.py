import argparse
import os
import json
import sys
from pathlib import Path

from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openai import OpenAI

_PYTHON_FILES = Path(__file__).resolve().parent.parent
if str(_PYTHON_FILES) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES))
from version import APP_VERSION

# Load .env from python_files/ (one level up from this script's subfolder)
load_dotenv(Path(__file__).resolve().parent.parent / ".env")

# =========================
# CONFIG
# =========================
OPENAI_API_KEY_ENV = "OPENAI_API_KEY"
API_KEY = os.getenv(OPENAI_API_KEY_ENV)

MODEL = "gpt-4o-mini"

client = OpenAI(api_key=API_KEY)


# =========================
# UTILS
# =========================
def clean_text(text) -> str | None:
    if text is None:
        return None
    return str(text).replace("\ufeff", "").strip() or None


# =========================
# HTML -> TEXT
# =========================
def extract_text_from_html(file_path: Path) -> str:
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        html = f.read()

    soup = BeautifulSoup(html, "html.parser")

    # Remove non-visible / noisy elements
    for tag in soup(["script", "style", "noscript", "svg", "meta", "head"]):
        tag.decompose()

    text = soup.get_text(separator="\n")

    # Clean whitespace
    lines = [line.strip() for line in text.splitlines()]
    lines = [line for line in lines if line]
    cleaned = "\n".join(lines)

    return cleaned


# =========================
# OPTIONAL: trim very long text
# =========================
def trim_text(text: str, max_chars: int = 50000) -> str:
    if len(text) <= max_chars:
        return text
    return text[:max_chars]


# =========================
# LLM EXTRACTION
# =========================
def extract_purchase_details(text_only: str, source_file: str, subject: str | None = None) -> dict:
    subject_section = f"\nEMAIL SUBJECT: {subject}" if subject else ""

    prompt = f"""
You are extracting structured purchase information from text that came from an HTML email/document.

Important rules:
1. Use ONLY the provided text and subject line.
2. Find the PURCHASE date/time, NOT the email received date.
3. If a value is missing or unclear, return null.
4. tracking_number should be the shipment tracking number only.
5. total_amount_paid should be the exact total paid as a number if possible.
6. tax_paid should be:
   - true if tax was charged
   - false if the text clearly shows no tax
   - null if unknown
7. purchase_datetime should be the exact purchase/order datetime if present.
   Format it as an ISO-like string when possible, such as:
   "2026-03-24 15:42:00"
   If only a date is known, return that date string like:
   "2026-03-24"
8. company_name should be the name of the company or retailer the order was placed with
   (e.g. "Amazon", "Walmart", "Best Buy"). Use the subject line as a hint if the body
   does not make it obvious. Return null if truly unknown.
9. order_number is the retailer's order/confirmation number (e.g. "#815419007417", "112-3456789-1234567").
   ALWAYS check the subject line first — it very commonly contains the order number (e.g.
   "Your order #12345 has shipped", "Order confirmation #98765"). Strip any leading "#" symbol.
   Fall back to the email body if not found in the subject. Return null only if truly absent.
10. Do not guess.

Source file: {source_file}{subject_section}

TEXT:
{text_only}
""".strip()

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "developer",
                "content": "Extract purchase details from email/document text and return only valid structured data."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "purchase_details",
                "schema": {
                    "type": "object",
                    "properties": {
                        "tracking_number": {
                            "type": ["string", "null"],
                            "description": "Shipment tracking number only, if present."
                        },
                        "total_amount_paid": {
                            "type": ["number", "null"],
                            "description": "Exact total amount paid."
                        },
                        "tax_paid": {
                            "type": ["boolean", "null"],
                            "description": "Whether tax was paid."
                        },
                        "purchase_datetime": {
                            "type": ["string", "null"],
                            "description": "Actual purchase/order datetime, not email date."
                        },
                        "company_name": {
                            "type": ["string", "null"],
                            "description": "Name of the company or retailer the order was placed with."
                        },
                        "order_number": {
                            "type": ["string", "null"],
                            "description": "Retailer order/confirmation number. Check subject line first, then body. Strip any leading '#'."
                        }
                    },
                    "required": [
                        "tracking_number",
                        "total_amount_paid",
                        "tax_paid",
                        "purchase_datetime",
                        "company_name",
                        "order_number"
                    ],
                    "additionalProperties": False
                }
            }
        },
        temperature=0
    )

    content = response.choices[0].message.content
    data = json.loads(content)
    return data


# =========================
# ARGS
# =========================
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=f"Email Sorter v{APP_VERSION}: extract structured purchase details from HTML emails."
    )
    parser.add_argument(
        "--base-dir",
        required=False,
        default=os.getenv("BASE_DIR"),
        dest="base_dir",
        help="Root project folder. Defaults to BASE_DIR in .env if not provided."
    )
    parser.add_argument(
        "--file",
        default=None,
        help=(
            "Filename (or full path) of a single HTML file to process. "
            "A bare filename is resolved inside <base-dir>/email_contents/html/. "
            "When provided, only that file is processed and the result is appended "
            "to the output JSON."
        )
    )
    parser.add_argument(
        "--subject",
        default=None,
        help="Email subject line. Passed to the LLM and embedded in the output JSON."
    )
    parser.add_argument(
        "--sender-name",
        default=None,
        dest="sender_name",
        help="Display name of the sender. Embedded in the output JSON as-is."
    )
    parser.add_argument(
        "--email",
        default=None,
        help="Sender email address. Embedded in the output JSON as-is."
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
    try:
        print(f"Processing: {file_path}")
        text_only = extract_text_from_html(file_path)
        text_only = trim_text(text_only, max_chars=50000)
        extracted = extract_purchase_details(text_only, str(file_path), subject=subject)
        file_uri = "file:///" + str(file_path.resolve()).replace("\\", "/")
        return {
            "source_file": clean_text(file_path),
            "source_file_link": file_uri,
            "subject": clean_text(subject),
            "sender_name": clean_text(sender_name),
            "email": clean_text(email),
            "company_name": clean_text(extracted.get("company_name")),
            "order_number": clean_text(extracted.get("order_number")),
            "tracking_number": clean_text(extracted.get("tracking_number")),
            "total_amount_paid": extracted.get("total_amount_paid"),
            "tax_paid": extracted.get("tax_paid"),
            "purchase_datetime": clean_text(extracted.get("purchase_datetime"))
        }
    except Exception as e:
        return {
            "source_file": clean_text(file_path),
            "source_file_link": None,
            "subject": clean_text(subject),
            "sender_name": clean_text(sender_name),
            "email": clean_text(email),
            "error": clean_text(e),
            "company_name": None,
            "order_number": None,
            "tracking_number": None,
            "total_amount_paid": None,
            "tax_paid": None,
            "purchase_datetime": None
        }


# =========================
# MAIN
# =========================
def main():
    args = parse_args()

    if not API_KEY:
        raise ValueError(f"{OPENAI_API_KEY_ENV} is not set in python_files/.env or environment.")

    if not args.base_dir:
        raise ValueError("BASE_DIR is not set. Add it to python_files/.env or pass --base-dir.")

    base = Path(args.base_dir)
    html_folder = base / "email_contents" / "html"
    output_path = base / "email_contents" / "json" / "results.json"

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Single-file mode
    if args.file:
        candidate = Path(args.file)
        # Resolve bare filenames relative to the html folder
        if not candidate.is_absolute() and candidate.parent == Path("."):
            candidate = html_folder / candidate
        file_path = candidate
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        entry = process_file(file_path, args.subject, args.sender_name, args.email)

        # Load existing results and append
        results = []
        if output_path.exists():
            for enc in ("utf-8-sig", "utf-16", "utf-8", "latin-1"):
                try:
                    with open(output_path, "r", encoding=enc) as f:
                        results = json.load(f)
                    break
                except (UnicodeDecodeError, json.JSONDecodeError):
                    continue

        results.append(entry)

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        print(f"Appended result to: {output_path}")
        return

    # Batch mode: process entire html folder
    if not html_folder.exists():
        raise FileNotFoundError(f"HTML input folder not found: {html_folder}")

    html_files = list(html_folder.rglob("*.html")) + list(html_folder.rglob("*.htm"))

    if not html_files:
        print(f"No HTML files found in: {INPUT_FOLDER}")
        return

    results = [process_file(fp, args.subject, args.sender_name, args.email) for fp in html_files]

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    print(f"\nDone. Wrote {len(results)} result(s) to: {output_path}")


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
    import sys
    from datetime import datetime

    _log_path = Path(__file__).resolve().parent / "run_log.txt"
    _tee = _Tee(_log_path, sys.stdout)
    sys.stdout = _tee
    sys.stderr = _Tee(_log_path, sys.stderr)

    print(f"\n{'='*60}")
    print(f"Email Sorter v{APP_VERSION}")
    print(f"Run started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Args: {sys.argv[1:]}")
    print(f"{'='*60}")

    _original_stdout = _tee._original
    _original_stderr = sys.stderr._original

    try:
        main()
        print("Run finished successfully.")
    except SystemExit as e:
        if e.code == 2:
            print("\nERROR: Invalid or missing arguments.")
            print("Set BASE_DIR and OPENAI_API_KEY in python_files/.env, or pass --base-dir as an argument.")
            print("Optional args: --file, --subject, --sender-name, --email")
        sys.stdout = _original_stdout
        sys.stderr = _original_stderr
        _tee.close()
        sys.exit(e.code)
    except Exception as e:
        print(f"\nERROR: {e}")
        sys.stdout = _original_stdout
        sys.stderr = _original_stderr
        _tee.close()
        sys.exit(1)

    sys.stdout = _original_stdout
    sys.stderr = _original_stderr
    _tee.close()