import json
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook

_PYTHON_FILES = Path(__file__).resolve().parent.parent
if str(_PYTHON_FILES) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES))
from version import APP_VERSION
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

BASE_DIR = os.getenv("BASE_DIR")
if not BASE_DIR:
    raise ValueError("BASE_DIR is not set in python_files/.env")

JSON_PATH  = str(Path(BASE_DIR) / "email_contents" / "json" / "results.json")
EXCEL_PATH = str(Path(BASE_DIR) / "email_contents" / "orders.xlsx")

COLUMN_ORDER = [
    "purchase_datetime",
    "order_number",
    "sender_name",
    "company_name",
    "email",
    "total_amount_paid",
    "tax_paid",
    "tracking_number",
    "source_file_link",
]

HEADERS = {
    "purchase_datetime":   "Purchase Date",
    "order_number":        "Order Number",
    "sender_name":         "Sender Name",
    "company_name":        "Company",
    "email":               "Email",
    "total_amount_paid":   "Total Paid",
    "tax_paid":            "Tax Paid",
    "tracking_number":     "Tracking Number",
    "source_file_link":    "View Email",
}

HYPERLINK_FONT = Font(name="Calibri", color="0563C1", underline="single")

HEADER_ROW = [HEADERS[col] for col in COLUMN_ORDER]

HEADER_FILL  = PatternFill("solid", fgColor="2F5597")
HEADER_FONT  = Font(bold=True, color="FFFFFF", name="Calibri")
CELL_FONT    = Font(name="Calibri")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN   = Alignment(horizontal="left",   vertical="center")


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


def style_header_row(ws, col_count: int):
    for col_idx in range(1, col_count + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.alignment = CENTER_ALIGN


def set_column_widths(ws):
    column_widths = {
        1:  22,  # Purchase Date
        2:  18,  # Order Number
        3:  20,  # Sender Name
        4:  20,  # Company
        5:  30,  # Email
        6:  14,  # Total Paid
        7:  12,  # Tax Paid
        8:  24,  # Tracking Number
        9:  14,  # View Email
    }
    for col_idx, width in column_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width


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


def build_workbook(records: list[dict]) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"

    ws.append(HEADER_ROW)
    style_header_row(ws, len(COLUMN_ORDER))

    for record in records:
        row = [clean_value(record.get(col)) for col in COLUMN_ORDER]
        ws.append(row)

    apply_cell_styles(ws, start_row=2)
    set_column_widths(ws)
    link_col_idx = COLUMN_ORDER.index("source_file_link") + 1
    apply_hyperlink_column(ws, link_col_idx, start_row=2)
    ws.freeze_panes = "B2"
    return wb


def append_to_workbook(path: str, records: list[dict]):
    wb = load_workbook(path)
    ws = wb.active

    existing_headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    if existing_headers[0] != HEADERS[COLUMN_ORDER[0]]:
        print("Warning: existing file headers don't match expected format. Appending anyway.")

    next_row = ws.max_row + 1

    for offset, record in enumerate(records):
        row = [clean_value(record.get(col)) for col in COLUMN_ORDER]
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=next_row + offset, column=col_idx, value=value)
            cell.font = CELL_FONT
            cell.alignment = LEFT_ALIGN

    link_col_idx = COLUMN_ORDER.index("source_file_link") + 1
    apply_hyperlink_column(ws, link_col_idx, start_row=next_row)
    wb.save(path)
    print(f"Appended {len(records)} row(s) to '{path}'.")


def apply_cell_styles(ws, start_row: int):
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row):
        for cell in row:
            cell.font = CELL_FONT
            cell.alignment = LEFT_ALIGN


def main():
    print(f"Email Sorter v{APP_VERSION}")
    records = load_json(JSON_PATH)

    if os.path.exists(EXCEL_PATH):
        append_to_workbook(EXCEL_PATH, records)
    else:
        wb = build_workbook(records)
        wb.save(EXCEL_PATH)
        print(f"Created '{EXCEL_PATH}' with {len(records)} row(s).")


if __name__ == "__main__":
    main()
