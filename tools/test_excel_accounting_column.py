from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

_TMP = tempfile.TemporaryDirectory()
os.environ.setdefault("BASE_DIR", _TMP.name)

from openpyxl import Workbook  # noqa: E402

from createExcelDocument import createExcelDocument as excel_doc  # noqa: E402


class ExcelAccountingColumnTests(unittest.TestCase):
    @classmethod
    def tearDownClass(cls) -> None:
        _TMP.cleanup()

    def test_populate_orders_sheet_adds_accounting_column_and_dropdown(self) -> None:
        wb = Workbook()
        excel_doc.populate_orders_sheet(
            wb,
            [
                {
                    "email_category": "Invoice",
                    "order_number": "1234",
                    "purchase_datetime": "2026-05-22",
                    "company": "Walgreens",
                    "total_amount_paid": 10.0,
                    "tax_paid": 1.0,
                }
            ],
        )
        ws = wb["Orders"]
        keys, _labels = excel_doc._build_column_order()
        accounting_col = keys.index("accounting") + 1
        tax_col = keys.index("tax_paid") + 1
        self.assertEqual(accounting_col, tax_col + 1)
        self.assertEqual(ws.cell(row=excel_doc.HEADER_ROW, column=accounting_col).value, "Accounting")
        self.assertIsNone(ws.cell(row=excel_doc.DATA_START_ROW, column=accounting_col).value)
        validations = list(ws.data_validations.dataValidation)
        self.assertTrue(
            any(f"{excel_doc.get_column_letter(accounting_col)}3" in str(dv.sqref) for dv in validations)
        )
        self.assertTrue(any(dv.formula1 == '"Complete,Review"' for dv in validations))

    def test_unknown_purchase_date_cell_is_blank(self) -> None:
        wb = Workbook()
        excel_doc.populate_orders_sheet(
            wb,
            [
                {
                    "email_category": "Invoice",
                    "order_number": "1234",
                    "purchase_datetime": None,
                    "company": "Walgreens",
                }
            ],
        )
        ws = wb["Orders"]
        keys, _labels = excel_doc._build_column_order()
        purchase_col = keys.index("purchase_datetime") + 1
        self.assertIsNone(ws.cell(row=excel_doc.DATA_START_ROW, column=purchase_col).value)

    def test_migrate_accounting_column_preserves_shifted_cells(self) -> None:
        wb = Workbook()
        ws = wb.active
        _keys, labels = excel_doc._build_column_order()
        old_labels = [label for label in labels if label != "Accounting"]
        ws.append([""] * len(old_labels))
        ws.append(old_labels)
        row = [None] * len(old_labels)
        row[old_labels.index("Category")] = "Invoice"
        row[old_labels.index("Order Number")] = "1234"
        row[old_labels.index("Purchase Date")] = "2026-05-22"
        row[old_labels.index("Company")] = "Walgreens"
        row[old_labels.index("Total Paid")] = 10
        row[old_labels.index("Tax Paid")] = 1
        row[old_labels.index("Invoice Link")] = "Invoice Link"
        ws.append(row)
        ws.cell(row=excel_doc.DATA_START_ROW, column=excel_doc.COPY_PATH_URI_COL, value="file:///C:/tmp/invoice.pdf")

        excel_doc._migrate_accounting_column_if_missing(ws, _keys)

        accounting_col = _keys.index("accounting") + 1
        self.assertEqual(ws.cell(row=excel_doc.HEADER_ROW, column=accounting_col).value, "Accounting")
        self.assertIsNone(ws.cell(row=excel_doc.DATA_START_ROW, column=accounting_col).value)
        self.assertEqual(ws.cell(row=excel_doc.DATA_START_ROW, column=accounting_col + 1).value, "Invoice Link")
        self.assertEqual(
            ws.cell(row=excel_doc.DATA_START_ROW, column=excel_doc.COPY_PATH_URI_COL).value,
            "file:///C:/tmp/invoice.pdf",
        )


if __name__ == "__main__":
    unittest.main()
