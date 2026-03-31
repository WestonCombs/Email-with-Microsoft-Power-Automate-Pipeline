import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

_base_dir_raw = os.getenv("BASE_DIR")
if not _base_dir_raw:
    raise ValueError("BASE_DIR is not set in python_files/.env")

PROJECT_ROOT = Path(_base_dir_raw).expanduser().resolve()
JSON_PATH  = str(PROJECT_ROOT / "email_contents" / "json" / "results.json")
EXCEL_PATH = str(PROJECT_ROOT / "email_contents" / "orders.xlsx")
LOG_PATH   = PROJECT_ROOT / "programFileOutput.txt"

# Column layout is split into three logical sections.
# "source_file_link" appears at the end of each section as a "View Email" hyperlink.
# A hairline vertical border is drawn after each section boundary.
COLUMN_ORDER = [
    # ── Section 1: order identity ──
    "email_category",
    "order_number",
    "purchase_datetime",
    "company",
    "email",
    "source_file_link",      # View Email (1) — section boundary
    # ── Section 2: financials ──
    "total_amount_paid",
    "tax_paid",
    "source_file_link",      # View Email (2) — section boundary
    # ── Section 3: shipping / misc ──
    "tracking_number",
    "tracking_link",
    "duplicate_on_last_run",
    "source_file_link",      # View Email (3)
]

COLUMN_HEADERS = {
    "email_category":        "Category",
    "order_number":          "Order Number",
    "purchase_datetime":     "Purchase Date",
    "company":               "Company",
    "email":                 "Email",
    "total_amount_paid":     "Total Paid",
    "tax_paid":              "Tax Paid",
    "tracking_number":       "Tracking Number",
    "tracking_link":         "Tracking Link",
    "duplicate_on_last_run": "Duplicate On Last Run",
    "source_file_link":      "View Email",
}

HYPERLINK_FONT = Font(name="Calibri", color="0563C1", underline="single")
HEADER_FILL    = PatternFill("solid", fgColor="2F5597")
HEADER_FONT    = Font(bold=True, color="FFFFFF", name="Calibri")
CELL_FONT      = Font(name="Calibri")
CENTER_ALIGN   = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN     = Alignment(horizontal="left",   vertical="center")

CATEGORY_FILLS = {
    "Invoice":   PatternFill("solid", fgColor="E3F2FD"),  # light sky
    "Shipped":   PatternFill("solid", fgColor="FFF3E0"),  # light peach
    "Delivered": PatternFill("solid", fgColor="E8F5E9"),  # light mint
    "Gift Card": PatternFill("solid", fgColor="FFF9C4"),  # light yellow
    "Unknown":   PatternFill("solid", fgColor="F3E5F5"),  # light purple
}


HAIR_SIDE  = Side(style="hair", color="000000")
THICK_SIDE = Side(style="medium", color="000000")


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


def _build_column_order() -> tuple[list[str], list[str]]:
    keys = list(COLUMN_ORDER)
    labels = [COLUMN_HEADERS[k] for k in keys]
    return keys, labels


def _record_to_row(record: dict, column_keys: list[str]) -> list:
    """Convert a single JSON record to a flat row list matching column_keys."""
    row: list = []
    for key in column_keys:
        if key == "company":
            row.append(get_company_value(record))
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
        "order_number":          18,
        "purchase_datetime":     22,
        "company":               24,
        "email":                 30,
        "total_amount_paid":     14,
        "tax_paid":              12,
        "tracking_number":       28,
        "tracking_link":         40,
        "duplicate_on_last_run": 24,
        "source_file_link":      14,
    }
    for col_idx, key in enumerate(column_keys, start=1):
        w = width_map.get(key, 16)
        ws.column_dimensions[get_column_letter(col_idx)].width = w


def _col_indices(column_keys: list[str], key: str) -> list[int]:
    """Return 1-based column indices for every occurrence of *key* in *column_keys*."""
    return [i + 1 for i, k in enumerate(column_keys) if k == key]


def apply_hyperlink_columns(ws, col_indices: list[int], start_row: int):
    """Replace file URI values with clickable 'Open' hyperlinks in every View Email column."""
    for col_idx in col_indices:
        for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row,
                                min_col=col_idx, max_col=col_idx):
            cell = row[0]
            uri = cell.value
            if uri and isinstance(uri, str) and uri.startswith("file:///"):
                cell.value = "Open"
                cell.hyperlink = uri
                cell.font = HYPERLINK_FONT
                cell.alignment = CENTER_ALIGN


def _merge_border(cell, top=None, bottom=None, left=None, right=None):
    """Merge new sides into a cell's existing border without overwriting unset sides."""
    old = cell.border
    cell.border = Border(
        top=top if top is not None else old.top,
        bottom=bottom if bottom is not None else old.bottom,
        left=left if left is not None else old.left,
        right=right if right is not None else old.right,
    )


def apply_order_group_borders(ws, records: list[dict], start_row: int):
    """Draw a bold bottom border on the last row of each order_number group."""
    if not records:
        return
    for i, record in enumerate(records):
        current = record.get("order_number")
        nxt = records[i + 1].get("order_number") if i + 1 < len(records) else None
        if current != nxt:
            row_idx = start_row + i
            for col_idx in range(1, ws.max_column + 1):
                _merge_border(ws.cell(row=row_idx, column=col_idx),
                              bottom=THICK_SIDE)


def apply_cell_styles(ws, start_row: int):
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row):
        for cell in row:
            cell.font = CELL_FONT
            cell.alignment = LEFT_ALIGN


def apply_category_colors(ws, start_row: int, category_col: int):
    """Shade every cell in a row with a pastel fill based on its email_category."""
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row):
        category = ws.cell(row=row[0].row, column=category_col).value
        fill = CATEGORY_FILLS.get(category)
        if fill:
            for cell in row:
                cell.fill = fill


def apply_row_borders(ws, start_row: int):
    """Add a hairline bottom border on every row from the header through the last data row."""
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            _merge_border(cell, bottom=HAIR_SIDE)


def apply_section_dividers(ws, boundary_cols: list[int]):
    """Draw a hairline right border on each section-boundary column (all rows)."""
    for col_idx in boundary_cols:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                min_col=col_idx, max_col=col_idx):
            _merge_border(row[0], right=HAIR_SIDE)


def apply_table_outline(ws):
    """Draw a hairline border around the entire table (header + data)."""
    last_row = ws.max_row
    last_col = ws.max_column

    for col_idx in range(1, last_col + 1):
        _merge_border(ws.cell(row=1, column=col_idx), top=HAIR_SIDE)
        _merge_border(ws.cell(row=last_row, column=col_idx), bottom=HAIR_SIDE)

    for row_idx in range(1, last_row + 1):
        _merge_border(ws.cell(row=row_idx, column=1), left=HAIR_SIDE)
        _merge_border(ws.cell(row=row_idx, column=last_col), right=HAIR_SIDE)


def apply_header_border(ws):
    """Draw a hairline top and bottom border on the header row (horizontal only)."""
    last_col = ws.max_column
    for col_idx in range(1, last_col + 1):
        _merge_border(ws.cell(row=1, column=col_idx),
                      top=HAIR_SIDE, bottom=HAIR_SIDE)


def build_workbook(records: list[dict]) -> Workbook:
    column_keys, header_labels = _build_column_order()

    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"

    ws.append(header_labels)
    style_header_row(ws, len(column_keys))

    for record in records:
        ws.append(_record_to_row(record, column_keys))

    link_cols = _col_indices(column_keys, "source_file_link")
    section_boundary_cols = link_cols[:-1]

    apply_cell_styles(ws, start_row=2)
    set_column_widths(ws, column_keys)

    category_col_idx = column_keys.index("email_category") + 1
    apply_category_colors(ws, start_row=2, category_col=category_col_idx)

    apply_row_borders(ws, start_row=2)
    apply_header_border(ws)
    apply_section_dividers(ws, section_boundary_cols)
    apply_table_outline(ws)
    apply_order_group_borders(ws, records, start_row=2)

    apply_hyperlink_columns(ws, link_cols, start_row=2)

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

    link_cols = _col_indices(desired_keys, "source_file_link")
    section_boundary_cols = link_cols[:-1]

    category_col_idx = desired_keys.index("email_category") + 1
    apply_category_colors(ws, start_row=next_row, category_col=category_col_idx)

    apply_row_borders(ws, start_row=next_row)
    apply_header_border(ws)
    apply_section_dividers(ws, section_boundary_cols)
    apply_table_outline(ws)
    apply_order_group_borders(ws, records, start_row=next_row)

    apply_hyperlink_columns(ws, link_cols, start_row=next_row)

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
