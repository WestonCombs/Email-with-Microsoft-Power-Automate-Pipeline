import json
import os
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

BASE_DIR = os.getenv("BASE_DIR")
if not BASE_DIR:
    raise ValueError("BASE_DIR is not set in python_files/.env")

INPUT_FILE  = Path(BASE_DIR) / "email_contents" / "json" / "results.json"
OUTPUT_FILE = INPUT_FILE

with open(INPUT_FILE, "r", encoding="utf-8") as f:
    data = json.load(f)

_DATE_FORMATS = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]

def _parse_datetime(value):
    """Return a datetime for sorting; nulls/unparseable values sort to the bottom."""
    if not value:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(value.strip(), fmt)
        except ValueError:
            continue
    return None

# Primary sort: order_number (nulls last); secondary sort: purchase_datetime (nulls last)
data.sort(key=lambda x: (
    x.get("order_number") is None,
    str(x.get("order_number") or ""),
    _parse_datetime(x.get("purchase_datetime")) is None,
    _parse_datetime(x.get("purchase_datetime")) or datetime.min,
))

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2, ensure_ascii=False)

print(f"Sorted {len(data)} records by order_number then purchase_datetime and saved to:\n{OUTPUT_FILE}")
