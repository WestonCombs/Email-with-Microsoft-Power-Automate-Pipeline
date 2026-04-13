import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import runLogger as RL

_base_dir_raw = os.getenv("BASE_DIR")
if not _base_dir_raw:
    raise ValueError("BASE_DIR is not set in python_files/.env")

PROJECT_ROOT = Path(_base_dir_raw).expanduser().resolve()
INPUT_FILE  = PROJECT_ROOT / "email_contents" / "json" / "results.json"
OUTPUT_FILE = INPUT_FILE

_DATE_FORMATS = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]


def _parse_datetime(value):
    """Return a datetime for sorting; nulls/unparseable values sort to the bottom."""
    if not value or not isinstance(value, str):
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(value.strip(), fmt)
        except ValueError:
            continue
    return None


def main():
    t = time.perf_counter()
    with open(INPUT_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    try:
        from grabbingImportantEmailContent.grabbingImportantEmailContent import (
            apply_order_company_consensus_and_sync,
        )

        apply_order_company_consensus_and_sync(data, PROJECT_ROOT)
    except Exception as e:
        print(f"  WARNING: order-level company consensus skipped: {e}")

    data.sort(key=lambda x: (
        x.get("order_number") is None,
        str(x.get("order_number") or ""),
        _parse_datetime(x.get("purchase_datetime")) is None,
        _parse_datetime(x.get("purchase_datetime")) or datetime.min,
    ))

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    elapsed = time.perf_counter() - t
    print(f"  Sorted {len(data)} records  ({elapsed:.2f}s)")
    RL.log("sortJSONByOrderNumber",
        f"{RL.ts()}  sorted {len(data)} records in {elapsed:.2f}s  →  {OUTPUT_FILE}"
    )


if __name__ == "__main__":
    print(f"\n{'='*60}")
    print(f"[sortJSONByOrderNumber] Run started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}")

    try:
        main()
        print("Sort finished successfully.")
    except Exception as e:
        print(f"\nERROR: {e}")
        sys.exit(1)
