import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

BASE_DIR = os.getenv("BASE_DIR")
if not BASE_DIR:
    raise ValueError("BASE_DIR is not set in python_files/.env")

JSON_PATH  = str(Path(BASE_DIR) / "email_contents" / "json" / "results.json")
EXCEL_PATH = str(Path(BASE_DIR) / "email_contents" / "orders.xlsx")
LOG_PATH   = Path(BASE_DIR) / "programFileOutput.txt"

# Fixed columns that always appear (in order).
FIXED_COLUMNS = [
    "email_category",
    "purchase_datetime",
    "order_number",
    "sender_name",
    "company",
    "email",
    "total_amount_paid",
    "tax_paid",
    "tracking_number",
    "duplicate_on_last_run",
]

FIXED_HEADERS = {
    "email_category":         "Email Category",
    "purchase_datetime":      "Purchase Date",
    "order_number":           "Order Number",
    "sender_name":            "Sender Name",
    "company":                "Company",
    "email":                  "Email",
    "total_amount_paid":      "Total Paid",
    "tax_paid":               "Tax Paid",
    "tracking_number":        "Tracking Number",
    "duplicate_on_last_run":  "Duplicate On Last Run",
}

# The final column is always source_file_link.
LINK_COLUMN = "source_file_link"
LINK_HEADER = "View Email"

HYPERLINK_FONT = Font(name="Calibri", color="0563C1", underline="single")
HEADER_FILL    = PatternFill("solid", fgColor="2F5597")
HEADER_FONT    = Font(bold=True, color="FFFFFF", name="Calibri")
CELL_FONT      = Font(name="Calibri")
CENTER_ALIGN   = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN     = Alignment(horizontal="left",   vertical="center")


def load_json(path: str) -> list[dict]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def clean_value(value):
    """Return None for JSON null and for the literal string 'null'."""
    if value is None:
        return None
    if isinstance(value, str) and value.strip().lower() == "null":
        return None
    return value


def infer_company_from_subject(subject) -> str | None:
    subject = clean_value(subject)
    if not isinstance(subject, str):
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
            company = clean_value(match.group(1))
            if company:
                return company.strip(" -,:;.!?")

    return None


def get_company_value(record: dict):
    explicit_company = clean_value(record.get("company"))
    if explicit_company:
        return explicit_company

    return infer_company_from_subject(record.get("subject"))


def get_tracking_number_value(record: dict):
    tracking_numbers = record.get("tracking_numbers", [])
    if not isinstance(tracking_numbers, list):
        return None

    cleaned_numbers = [
        str(value).strip()
        for value in tracking_numbers
        if value is not None and str(value).strip()
    ]
    return ", ".join(cleaned_numbers) if cleaned_numbers else None


def _build_column_order() -> tuple[list[str], list[str]]:
    keys: list[str] = list(FIXED_COLUMNS)
    labels: list[str] = [FIXED_HEADERS[c] for c in FIXED_COLUMNS]

    keys.append(LINK_COLUMN)
    labels.append(LINK_HEADER)

    return keys, labels


def _record_to_row(record: dict, column_keys: list[str]) -> list:
    """Convert a single JSON record to a flat row list matching column_keys."""
    row: list = []
    for key in column_keys:
        if key == "company":
            row.append(get_company_value(record))
        elif key == "tracking_number":
            row.append(get_tracking_number_value(record))
        else:
            row.append(clean_value(record.get(key)))
    return row


def style_header_row(ws, col_count: int):
    for col_idx in range(1, col_count + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER_ALIGN


def set_column_widths(ws, column_keys: list[str]):
    width_map = {
        "email_category":        22,
        "purchase_datetime":     22,
        "order_number":          18,
        "sender_name":           20,
        "company":               24,
        "email":                 30,
        "total_amount_paid":     14,
        "tax_paid":              12,
        "tracking_number":       24,
        "duplicate_on_last_run": 24,
        "source_file_link":      14,
    }
    for col_idx, key in enumerate(column_keys, start=1):
        if key in width_map:
            w = width_map[key]
        else:
            w = 16
        ws.column_dimensions[get_column_letter(col_idx)].width = w


def apply_hyperlink_column(ws, col_idx: int, start_row: int):
    """Replace file URI values in a column with clickable 'Open' hyperlink cells."""
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
        cell = row[0]
        uri = cell.value
        if uri and isinstance(uri, str) and uri.startswith("file:///"):
            cell.value = "Open"
            cell.hyperlink = uri
            cell.font = HYPERLINK_FONT
            cell.alignment = CENTER_ALIGN


def apply_cell_styles(ws, start_row: int):
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row):
        for cell in row:
            cell.font = CELL_FONT
            cell.alignment = LEFT_ALIGN


def build_workbook(records: list[dict]) -> Workbook:
    column_keys, header_labels = _build_column_order()

    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"

    ws.append(header_labels)
    style_header_row(ws, len(column_keys))

    for record in records:
        ws.append(_record_to_row(record, column_keys))

    apply_cell_styles(ws, start_row=2)
    set_column_widths(ws, column_keys)

    link_col_idx = column_keys.index(LINK_COLUMN) + 1
    apply_hyperlink_column(ws, link_col_idx, start_row=2)

    ws.freeze_panes = "B2"
    return wb


def append_to_workbook(path: str, records: list[dict]):
    wb = load_workbook(path)
    ws = wb.active

    existing_headers = [
        ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)
    ]

    desired_keys, desired_labels = _build_column_order()

    if desired_labels != existing_headers:
        for col_idx, label in enumerate(desired_labels, start=1):
            cell = ws.cell(row=1, column=col_idx, value=label)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = CENTER_ALIGN
        for col_idx in range(len(desired_labels) + 1, ws.max_column + 1):
            ws.cell(row=1, column=col_idx, value=None)
        set_column_widths(ws, desired_keys)

    next_row = ws.max_row + 1

    for offset, record in enumerate(records):
        row_data = _record_to_row(record, desired_keys)
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=next_row + offset, column=col_idx, value=value)
            cell.font = CELL_FONT
            cell.alignment = LEFT_ALIGN

    link_col_idx = desired_keys.index(LINK_COLUMN) + 1
    apply_hyperlink_column(ws, link_col_idx, start_row=next_row)
    ws.freeze_panes = "B2"

    wb.save(path)
    print(f"Appended {len(records)} row(s) to '{path}'.")


def reset_duplicate_flags(json_path: str):
    """Reset all duplicate_on_last_run flags to 0 in the JSON file."""
    with open(json_path, "r", encoding="utf-8") as f:
        records = json.load(f)

    for record in records:
        record["duplicate_on_last_run"] = 0

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)

    print(f"Reset duplicate_on_last_run to 0 for {len(records)} record(s).")


def main():
    records = load_json(JSON_PATH)

    wb = build_workbook(records)
    wb.save(EXCEL_PATH)
    print(f"Created '{EXCEL_PATH}' with {len(records)} row(s).")

    reset_duplicate_flags(JSON_PATH)


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
    _tee = _Tee(LOG_PATH, sys.stdout)
    sys.stdout = _tee
    sys.stderr = _Tee(LOG_PATH, sys.stderr)
    _original_stdout = _tee._original
    _original_stderr = sys.stderr._original

    print(f"\n{'='*60}")
    print(f"[createExcelDocument] Run started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}")

    try:
        main()
        print("Excel creation finished successfully.")
    except Exception as e:
        print(f"\nERROR: {e}")
        sys.stdout = _original_stdout
        sys.stderr = _original_stderr
        _tee.close()
        sys.exit(1)

    sys.stdout = _original_stdout
    sys.stderr = _original_stderr
    _tee.close()
