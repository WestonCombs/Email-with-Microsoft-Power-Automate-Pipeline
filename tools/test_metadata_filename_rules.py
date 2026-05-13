"""Focused metadata, filename, and image-diagnostic checks."""

from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ.setdefault("OPENAI_API_KEY", "test-key")

from grabbingImportantEmailContent import grabbingImportantEmailContent as extractor  # noqa: E402


class MetadataFilenameRuleTests(unittest.TestCase):
    def test_no7_sender_beats_boots_legal_footer(self) -> None:
        self.assertEqual(
            extractor.choose_company_display(
                "The Boots Company PLC",
                "Your No7 order",
                "No7",
                "service@t.us.no7beauty.com",
            ),
            "No7",
        )

    def test_iris_aliases_normalize(self) -> None:
        self.assertEqual(extractor.normalize_company_display_name("Iris&Romeo"), "Iris & Romeo")
        self.assertEqual(
            extractor.normalize_company_display_name("irisandromeo.com"),
            "Iris & Romeo",
        )

    def test_fragrancenet_aliases_normalize(self) -> None:
        self.assertEqual(
            extractor.normalize_company_display_name("fragrancenet"),
            "FragranceNet.com",
        )
        self.assertEqual(
            extractor.normalize_company_display_name("fragrancenet.com"),
            "FragranceNet.com",
        )

    def test_zara_domain_fallback(self) -> None:
        self.assertEqual(
            extractor.choose_company_display(None, "", "", "orders@zara.com"),
            "Zara",
        )

    def test_filename_no_confident_date_uses_last4(self) -> None:
        self.assertEqual(
            extractor.build_convention_filename(
                {
                    "company": "FragranceNet.com",
                    "email_category": "Delivered",
                    "order_number": "393334447",
                    "purchase_datetime": None,
                }
            ),
            "DOC FragranceNet.com 4447 DELIVERED.pdf",
        )

    def test_filename_disallowed_date_source_uses_last4(self) -> None:
        record = {
            "company": "FragranceNet.com",
            "email_category": "Delivered",
            "order_number": "393334447",
            "purchase_datetime": "2026-04-28",
            "purchase_datetime_source": "forwarded_received_date",
            "purchase_datetime_confidence": "low",
        }
        self.assertFalse(extractor.is_filename_date_allowed(record))
        self.assertEqual(
            extractor.build_convention_filename(record),
            "DOC FragranceNet.com 4447 DELIVERED.pdf",
        )

    def test_filename_allowed_explicit_order_date(self) -> None:
        record = {
            "company": "Iris & Romeo",
            "email_category": "Invoice",
            "order_number": "IR181658",
            "purchase_datetime": "2026-04-14",
            "purchase_datetime_source": "explicit_order_date",
            "purchase_datetime_confidence": "high",
        }
        self.assertTrue(extractor.is_filename_date_allowed(record))
        self.assertEqual(
            extractor.build_convention_filename(record),
            "DOC Iris & Romeo 2026-04-14 INVOICE_1658.pdf",
        )

    def test_image_diagnostics_detect_delivery_remote_image(self) -> None:
        diag = extractor.diagnose_html_images_for_pdf(
            '<p>Image of Delivery</p><img src="https://example.com/a.png">'
        )
        self.assertTrue(diag["delivery_image_text_present"])
        self.assertEqual(diag["remote_src"], 1)


if __name__ == "__main__":
    unittest.main()
