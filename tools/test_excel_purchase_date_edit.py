from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from createExcelDocument import macro_template  # noqa: E402
from shared.excel_user_edits import (  # noqa: E402
    coerce_user_edit_value,
    record_excel_user_edit,
)


class ExcelPurchaseDateEditTests(unittest.TestCase):
    def test_purchase_date_edit_accepts_common_date_and_marks_source(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            results_path = json_dir / "results.json"
            source_uri = "file:///C:/tmp/DOC Example Store 1234 INVOICE.pdf"
            results_path.write_text(
                json.dumps(
                    [
                        {
                            "source_file_link": source_uri,
                            "order_number": "1234",
                            "purchase_datetime": "",
                            "email_category": "Invoice",
                            "company": "Example Store",
                        }
                    ],
                    indent=2,
                ),
                encoding="utf-8",
            )

            summary = record_excel_user_edit(
                project_root,
                field="purchase_datetime",
                raw_value="5/22/2026",
                order_number="1234",
                source_uri=source_uri,
            )

            rows = json.loads(results_path.read_text(encoding="utf-8"))
            self.assertEqual(rows[0]["purchase_datetime"], "2026-05-22")
            self.assertEqual(rows[0]["purchase_datetime_source"], "user_excel_edit")
            self.assertEqual(rows[0]["purchase_datetime_confidence"], "high")
            self.assertTrue(rows[0]["modified_purchase_datetime"])
            self.assertEqual(summary["display_value"], "2026-05-22")

    def test_purchase_date_edit_rejects_not_a_date(self) -> None:
        with self.assertRaises(ValueError):
            coerce_user_edit_value("purchase_datetime", "not a date")

    def test_macro_allows_purchase_date_column(self) -> None:
        self.assertIn("Purchase Date", macro_template.EMAIL_SORTER_HOTKEYS_VBA)
        self.assertIn('Case "purchase date"', macro_template.EMAIL_SORTER_HOTKEYS_VBA)
        self.assertIn('EmailSorter_FieldKeyForColumn = "purchase_datetime"', macro_template.EMAIL_SORTER_HOTKEYS_VBA)


if __name__ == "__main__":
    unittest.main()
