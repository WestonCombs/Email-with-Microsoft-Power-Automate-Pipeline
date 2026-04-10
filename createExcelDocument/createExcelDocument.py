import importlib.util
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

_PYTHON_FILES_DIR = Path(__file__).resolve().parent.parent
if str(_PYTHON_FILES_DIR) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES_DIR))

from dotenv import load_dotenv
from htmlHandler.tracking_hrefs import MULTIPLE_TRACKING_LINKS
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

_base_dir_raw = os.getenv("BASE_DIR")
if not _base_dir_raw:
    raise ValueError("BASE_DIR is not set in python_files/.env")

PROJECT_ROOT = Path(_base_dir_raw).expanduser().resolve()
JSON_PATH  = str(PROJECT_ROOT / "email_contents" / "json" / "results.json")

# Hidden on Orders; VBA reads this before ThisWorkbook.Path (works if Path is empty).
CLIPBOARD_INI_CELL = "AA1"
# Hidden: plain-text file:/// URI per row. Copy Path cells use internal # links only
# (Excel hyperlink events cannot cancel file:// navigation).
COPY_PATH_URI_COL = 28  # column AB — keep equal to COL_FILE_URI in macro_template.py
# Hidden full tracking URLs: column 29 + slot-1. Must match VBA COL_TRACK_URI_* in macro_template.py
TRACKING_URI_COL_START = 29  # AC
# One hidden column per visible tracking slot (Track 1 … Track N).
TRACKING_LINK_VISIBLE_SLOTS = 15
TRACKING_URI_SLOT_COUNT = TRACKING_LINK_VISIBLE_SLOTS  # AC through (28+N)

# Optional .env overrides (absolute or relative paths expand from user / cwd).
def _optional_path(env_name: str, default: Path) -> Path:
    raw = os.getenv(env_name)
    if raw:
        return Path(raw).expanduser().resolve()
    return default


ORDERS_TEMPLATE_PATH = _optional_path(
    "EXCEL_TEMPLATE_PATH", _PYTHON_FILES_DIR / "orders_template.xlsm"
)
# VBA reads absolute path from cell AA1; ini can live outside email_contents.
CLIPBOARD_LAUNCH_INI_PATH = _optional_path(
    "EXCEL_CLIPBOARD_INI_PATH", _PYTHON_FILES_DIR / "excel_clipboard_launch.ini"
)
EXCEL_PATH = str(
    _optional_path("EXCEL_OUTPUT_PATH", PROJECT_ROOT / "email_contents" / "orders.xlsx")
)

# Launched from Excel VBA (VIEWER= in excel_clipboard_launch.ini) for the "View Link List" column.
TRACKING_VIEWER_SCRIPT = _PYTHON_FILES_DIR / "trackingLinkViewer" / "tracking_link_viewer.py"


def _resolve_excel_output_path(using_template: bool) -> str:
    """Use .xlsx unless we loaded a real .xlsm template (VBA). Plain openpyxl .xlsm saves break Excel."""
    p = Path(EXCEL_PATH)
    if using_template:
        if p.suffix.lower() != ".xlsm":
            out = str(p.with_suffix(".xlsm"))
            print(
                f"Note: Macro template in use - saving to '{out}' (.xlsm) instead of '{p.name}'."
            )
            return out
        return str(p)
    if p.suffix.lower() == ".xlsm":
        out = str(p.with_suffix(".xlsx"))
        print(
            f"Note: No orders_template.xlsm — saving to '{out}' (.xlsx). "
            "Saving plain openpyxl output as .xlsm is invalid and Excel will not open it."
        )
        return out
    return str(p)


def _macro_template_module():
    """Load sibling macro_template.py without treating this folder as a package."""
    mpath = Path(__file__).resolve().parent / "macro_template.py"
    spec = importlib.util.spec_from_file_location("_email_sorter_macro_template", mpath)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load macro helper: {mpath}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# Column layout: order identity, then Copy Path + View Email (adjacent, right after Email),
# then financials and shipping. Copy Path uses in-sheet # links; hidden column AB holds the
# file URI for VBA. View Email is the file "Open" hyperlink.
# Hairline vertical borders after View Email and after Tax Paid (section breaks).
SECTION_BOUNDARY_KEYS = ("source_file_link", "tax_paid")

# Tracking URLs: hidden columns AC–AQ for VBA + ``View Link List`` launches ``trackingLinkViewer``.
_COLUMN_ORDER_PREFIX = [
    "email_category",
    "order_number",
    "purchase_datetime",
    "company",
    "email",
    "copy_file_path",
    "source_file_link",
    "total_amount_paid",
    "tax_paid",
]
_COLUMN_ORDER_SUFFIX = ["open_tracking_list"]

COLUMN_ORDER = _COLUMN_ORDER_PREFIX + _COLUMN_ORDER_SUFFIX

COLUMN_HEADERS = {
    "email_category":        "Category",
    "order_number":          "Order Number",
    "purchase_datetime":     "Purchase Date",
    "company":               "Company",
    "email":                 "Email",
    "total_amount_paid":     "Total Paid",
    "tax_paid":              "Tax Paid",
    "open_tracking_list":    "View Link List",
    "copy_file_path":        "Copy Path",
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
        elif key == "copy_file_path":
            row.append(clean_value(record.get("source_file_link")))
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
        "open_tracking_list":    18,
        "copy_file_path":        12,
        "source_file_link":      14,
    }
    for col_idx, key in enumerate(column_keys, start=1):
        w = width_map.get(key, 16)
        ws.column_dimensions[get_column_letter(col_idx)].width = w


def _col_indices(column_keys: list[str], key: str) -> list[int]:
    """Return 1-based column indices for every occurrence of *key* in *column_keys*."""
    return [i + 1 for i, k in enumerate(column_keys) if k == key]


def apply_hyperlink_columns(ws, col_indices: list[int], start_row: int):
    """Replace file URI values with clickable 'Open' hyperlinks (View Email / source column)."""
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


def _sheet_name_for_excel_ref(name: str) -> str:
    """Build sheet name token for #'Name'!$A$1 style hyperlinks."""
    if "'" in name:
        return "'" + name.replace("'", "''") + "'"
    if " " in name or not name.replace("_", "").isalnum():
        return "'" + name + "'"
    return name


def apply_copy_path_hyperlink_columns(
    ws, col_indices: list[int], start_row: int, records: list[dict]
):
    """Copy Path: in-workbook # links only; real file URI in hidden column COPY_PATH_URI_COL."""
    col_uri = COPY_PATH_URI_COL
    sn = _sheet_name_for_excel_ref(ws.title)

    for i, record in enumerate(records):
        row_idx = start_row + i
        uri = clean_value(record.get("source_file_link"))
        if uri and isinstance(uri, str) and uri.startswith("file:///"):
            ws.cell(row=row_idx, column=col_uri, value=uri)

    ws.column_dimensions[get_column_letter(col_uri)].hidden = True

    for col_idx in col_indices:
        for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row,
                                min_col=col_idx, max_col=col_idx):
            cell = row[0]
            row_num = cell.row
            stored = ws.cell(row=row_num, column=col_uri).value
            if not stored or not isinstance(stored, str) or not stored.startswith("file:///"):
                continue
            col_letter = get_column_letter(col_idx)
            cell.value = "Copy Path"
            cell.hyperlink = f"#{sn}!${col_letter}${row_num}"
            cell.font = HYPERLINK_FONT
            cell.alignment = CENTER_ALIGN


def _tracking_urls_for_record(record: dict) -> list[str]:
    """Prefer ``tracking_links``; fall back to legacy ``tracking_link`` string."""
    raw = record.get("tracking_links")
    if isinstance(raw, list) and raw:
        return [str(u).strip() for u in raw if isinstance(u, str) and str(u).strip()]
    tl = clean_value(record.get("tracking_link"))
    if not isinstance(tl, str) or not tl.strip():
        return []
    s = tl.strip()
    if s == MULTIPLE_TRACKING_LINKS:
        return []
    low = s.lower()
    if low.startswith("http://") or low.startswith("https://"):
        return [s]
    return []


def apply_hidden_tracking_url_columns(
    ws,
    records: list[dict],
    start_row: int,
    *,
    vba_friendly: bool = False,
):
    """Write JSON ``tracking_links`` into hidden columns (29…43) for VBA / View Link List viewer."""
    if not records:
        return
    last_uri_col = TRACKING_URI_COL_START + TRACKING_URI_SLOT_COUNT - 1

    for i, record in enumerate(records):
        row_idx = start_row + i
        urls = _tracking_urls_for_record(record)
        for uc in range(TRACKING_URI_COL_START, last_uri_col + 1):
            ws.cell(row=row_idx, column=uc, value=None)
        if not vba_friendly:
            continue
        for j, u in enumerate(urls[:TRACKING_URI_SLOT_COUNT]):
            ws.cell(row=row_idx, column=TRACKING_URI_COL_START + j, value=u)

    if vba_friendly:
        for uc in range(TRACKING_URI_COL_START, last_uri_col + 1):
            ws.column_dimensions[get_column_letter(uc)].hidden = True


def apply_open_tracking_list_column(
    ws,
    column_keys: list[str],
    records: list[dict],
    start_row: int,
    *,
    vba_friendly: bool = False,
):
    """``View Link List`` cell: VBA launches :mod:`trackingLinkViewer` with all row tracking URLs."""
    cols = _col_indices(column_keys, "open_tracking_list")
    if not cols or not records:
        return
    col_idx = cols[0]
    col_letter = get_column_letter(col_idx)
    sn = _sheet_name_for_excel_ref(ws.title)

    for i, record in enumerate(records):
        row_idx = start_row + i
        cell = ws.cell(row=row_idx, column=col_idx)
        urls = _tracking_urls_for_record(record)
        cell.hyperlink = None
        if not urls:
            cell.value = None
            continue
        if vba_friendly:
            cell.value = "View Link List"
            cell.hyperlink = f"#{sn}!${col_letter}${row_idx}"
            cell.font = HYPERLINK_FONT
            cell.alignment = CENTER_ALIGN
        else:
            cell.value = "—"
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
    """Draw a hairline right border on each listed column (all rows)."""
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


def populate_orders_sheet(wb: Workbook, records: list[dict]) -> None:
    """Rebuild the Orders sheet from scratch (headers + rows + styling + hyperlinks)."""
    vba_friendly = getattr(wb, "vba_archive", None) is not None
    column_keys, header_labels = _build_column_order()

    if "Orders" in wb.sheetnames:
        ws = wb["Orders"]
    else:
        ws = wb.active
        ws.title = "Orders"

    if ws.max_row >= 1:
        ws.delete_rows(1, ws.max_row)

    ws.append(header_labels)
    style_header_row(ws, len(column_keys))

    for record in records:
        ws.append(_record_to_row(record, column_keys))

    copy_cols = _col_indices(column_keys, "copy_file_path")
    source_cols = _col_indices(column_keys, "source_file_link")
    section_boundary_cols = [
        idx for key in SECTION_BOUNDARY_KEYS for idx in _col_indices(column_keys, key)
    ]

    apply_cell_styles(ws, start_row=2)
    set_column_widths(ws, column_keys)

    category_col_idx = column_keys.index("email_category") + 1
    apply_category_colors(ws, start_row=2, category_col=category_col_idx)

    apply_row_borders(ws, start_row=2)
    apply_header_border(ws)
    apply_section_dividers(ws, section_boundary_cols)
    apply_table_outline(ws)
    apply_order_group_borders(ws, records, start_row=2)

    apply_copy_path_hyperlink_columns(ws, copy_cols, start_row=2, records=records)
    apply_hyperlink_columns(ws, source_cols, start_row=2)
    apply_hidden_tracking_url_columns(ws, records, start_row=2, vba_friendly=vba_friendly)
    apply_open_tracking_list_column(
        ws, column_keys, records, start_row=2, vba_friendly=vba_friendly
    )

    ws.freeze_panes = "B2"


def set_clipboard_ini_cell(wb: Workbook, ini_path: Path) -> None:
    """Write absolute ini path for Workbook_SheetFollowHyperlink; hide column AA."""
    ws = wb["Orders"]
    ws[CLIPBOARD_INI_CELL] = str(ini_path.resolve())
    ws.column_dimensions["AA"].hidden = True


def build_workbook(records: list[dict]) -> Workbook:
    wb = Workbook()
    populate_orders_sheet(wb, records)
    return wb


def _load_workbook_editable(path: str) -> Workbook:
    """Preserve VBA when editing macro-enabled workbooks."""
    if Path(path).suffix.lower() == ".xlsm":
        return load_workbook(path, keep_vba=True)
    return load_workbook(path)


def append_to_workbook(path: str, records: list[dict]):
    wb = _load_workbook_editable(path)
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

    copy_cols = _col_indices(desired_keys, "copy_file_path")
    source_cols = _col_indices(desired_keys, "source_file_link")
    section_boundary_cols = [
        idx for key in SECTION_BOUNDARY_KEYS for idx in _col_indices(desired_keys, key)
    ]

    category_col_idx = desired_keys.index("email_category") + 1
    apply_category_colors(ws, start_row=next_row, category_col=category_col_idx)

    apply_row_borders(ws, start_row=next_row)
    apply_header_border(ws)
    apply_section_dividers(ws, section_boundary_cols)
    apply_table_outline(ws)
    apply_order_group_borders(ws, records, start_row=next_row)

    vba_friendly = Path(path).suffix.lower() == ".xlsm" and getattr(
        wb, "vba_archive", None
    ) is not None
    apply_copy_path_hyperlink_columns(ws, copy_cols, start_row=next_row, records=records)
    apply_hyperlink_columns(ws, source_cols, start_row=next_row)
    apply_hidden_tracking_url_columns(
        ws, records, start_row=next_row, vba_friendly=vba_friendly
    )
    apply_open_tracking_list_column(
        ws, desired_keys, records, start_row=next_row, vba_friendly=vba_friendly
    )

    ws.freeze_panes = "B2"

    if Path(path).suffix.lower() == ".xlsm" and "Orders" in wb.sheetnames:
        macro_mod = _macro_template_module() if sys.platform == "win32" else None
        if macro_mod is not None:
            script_path = Path(__file__).resolve().parent / "copy_email_path_to_clipboard.py"
            ini_written = macro_mod.write_clipboard_launch_ini(
                CLIPBOARD_LAUNCH_INI_PATH,
                sys.executable,
                script_path,
                viewer_script=TRACKING_VIEWER_SCRIPT,
            )
            set_clipboard_ini_cell(wb, ini_written)

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

    using_template = ORDERS_TEMPLATE_PATH.is_file()
    macro_mod = None
    if sys.platform == "win32":
        macro_mod = _macro_template_module()
    if not using_template and macro_mod is not None:
        if macro_mod.ensure_macro_template(ORDERS_TEMPLATE_PATH):
            using_template = ORDERS_TEMPLATE_PATH.is_file()

    if using_template:
        wb = load_workbook(str(ORDERS_TEMPLATE_PATH), keep_vba=True)
        populate_orders_sheet(wb, records)
    else:
        wb = build_workbook(records)

    out_path = _resolve_excel_output_path(using_template)
    if using_template:
        m = macro_mod if macro_mod is not None else _macro_template_module()
        script_path = Path(__file__).resolve().parent / "copy_email_path_to_clipboard.py"
        ini_written = m.write_clipboard_launch_ini(
            CLIPBOARD_LAUNCH_INI_PATH,
            sys.executable,
            script_path,
            viewer_script=TRACKING_VIEWER_SCRIPT,
        )
        set_clipboard_ini_cell(wb, ini_written)
        print(f"Wrote clipboard launcher config: {ini_written}")

    wb.save(out_path)
    print(f"Wrote '{out_path}' with {len(records)} row(s).")
    if not using_template:
        print(
            "Note: No macro template - output is .xlsx; Copy Path cells open the file. "
            "On Windows, install pywin32 + Excel so the program can auto-create "
            f"'{ORDERS_TEMPLATE_PATH}', or add that file manually (CLIPBOARD_SETUP.txt)."
        )

    reset_duplicate_flags(JSON_PATH)


if __name__ == "__main__":
    print(f"\n{'='*60}")
    print(f"[createExcelDocument] Run started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}")

    try:
        main()
        print("Excel creation finished successfully.")
    except Exception as e:
        print(f"\nERROR: {e}")
        sys.exit(1)
