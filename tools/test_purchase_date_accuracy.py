"""Focused checks for purchase-date accuracy fallbacks.

These tests avoid live email/OpenAI calls and exercise the consolidation and
filename behavior that can otherwise turn mailbox timing into an inaccurate
purchase date.
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ.setdefault("OPENAI_API_KEY", "test-key")

from grabbingImportantEmailContent import grabbingImportantEmailContent as extractor  # noqa: E402

extractor.RL.log = lambda *args, **kwargs: None

build_convention_filename = extractor.build_convention_filename
unify_purchase_dates_by_order = extractor.unify_purchase_dates_by_order


def _main_extraction_result(purchase_datetime: str | None = None) -> dict:
    return {
        "company": "Example Store",
        "order_number": "A1",
        "purchase_datetime": purchase_datetime,
        "total_amount_paid": 25.0,
        "subtotal_amount": 25.0,
        "tax_paid": 0.0,
        "gift_card_amount": None,
        "invoice_total_needs_review": False,
        "invoice_total_review_reason": None,
        "tracking_numbers": [],
        "email_category": "Invoice",
        "email_category_confidence": 99,
    }


class PurchaseDateAccuracyTests(unittest.TestCase):
    def test_grounded_medium_confidence_rescue_is_accepted(self) -> None:
        source = "Your payment was completed June 7, 2026 for order A1."
        result = extractor.validate_purchase_date_rescue_result(
            {
                "purchase_datetime": "2026-06-07",
                "source_type": "purchase_order_date",
                "confidence": "medium",
                "evidence": "payment was completed June 7, 2026",
                "reason": None,
            },
            source_text=source,
        )

        self.assertEqual(
            result,
            (
                "2026-06-07",
                "llm_date_rescue",
                "medium",
                "payment was completed June 7, 2026",
                None,
            ),
        )

    def test_rescue_rejects_delivery_date_even_when_model_calls_it_purchase_date(self) -> None:
        result = extractor.validate_purchase_date_rescue_result(
            {
                "purchase_datetime": "2026-06-09",
                "source_type": "purchase_order_date",
                "confidence": "high",
                "evidence": "June 9, 2026",
                "reason": None,
            },
            source_text="Your package was delivered June 9, 2026",
        )

        self.assertIsNone(result[0])
        self.assertIn("source context describes a non-purchase event date", result[4] or "")

    def test_rescue_rejects_evidence_not_copied_from_email(self) -> None:
        result = extractor.validate_purchase_date_rescue_result(
            {
                "purchase_datetime": "2026-06-07",
                "source_type": "purchase_order_date",
                "confidence": "high",
                "evidence": "Payment Date: June 7, 2026",
                "reason": None,
            },
            source_text="Thanks for your order. Your payment was successful.",
        )

        self.assertIsNone(result[0])
        self.assertIn("not found in the provided email", result[4] or "")

    def test_rescue_rejects_invalid_calendar_date(self) -> None:
        result = extractor.validate_purchase_date_rescue_result(
            {
                "purchase_datetime": "2026-02-30",
                "source_type": "purchase_order_date",
                "confidence": "high",
                "evidence": "Payment Date: 2026-02-30",
                "reason": None,
            },
            source_text="Payment Date: 2026-02-30",
        )

        self.assertIsNone(result[0])
        self.assertIn("invalid calendar date", result[4] or "")

    def test_process_file_runs_rescue_and_saves_audit_evidence(self) -> None:
        rescue_result = {
            "purchase_datetime": "2026-06-07",
            "source_type": "purchase_order_date",
            "confidence": "high",
            "evidence": "Completed June 7, 2026.",
            "reason": None,
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            email_path = Path(temp_dir) / "invoice.html"
            email_path.write_text(
                "<html><body><p>Thank you for shopping with Example Store.</p>"
                "<p>Completed June 7, 2026.</p></body></html>",
                encoding="utf-8",
            )
            with (
                patch.object(extractor, "API_KEY", "test-key"),
                patch.object(extractor, "extract_with_openai", return_value=_main_extraction_result()),
                patch.object(extractor, "should_run_is_gift_card", return_value=False),
                patch.object(
                    extractor,
                    "extract_purchase_date_rescue_with_openai",
                    return_value=rescue_result,
                ) as rescue_call,
                patch.object(extractor, "_write_openai_log"),
                patch.object(extractor, "_write_tracking_log"),
                redirect_stdout(StringIO()),
            ):
                record = extractor.process_file(
                    email_path,
                    subject="Your order A1 is confirmed",
                    sender_name="Example Store",
                    email="orders@example.com",
                )

        rescue_call.assert_called_once()
        self.assertEqual(record["purchase_datetime"], "2026-06-07")
        self.assertEqual(record["purchase_datetime_source"], "llm_date_rescue")
        self.assertEqual(record["purchase_datetime_confidence"], "high")
        self.assertTrue(record["purchase_datetime_rescue_attempted"])
        self.assertEqual(record["purchase_datetime_rescue_source_type"], "purchase_order_date")
        self.assertEqual(record["purchase_datetime_rescue_evidence"], "Completed June 7, 2026.")
        self.assertIsNone(record["purchase_datetime_rescue_reason"])
        self.assertTrue(record["_timings"]["step5c_date_rescue_ran"])
        self.assertTrue(extractor.is_filename_date_allowed(record))
        self.assertIn("2026-06-07", build_convention_filename(record))

    def test_process_file_skips_rescue_when_rule_finds_labeled_order_date(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            email_path = Path(temp_dir) / "invoice.html"
            email_path.write_text(
                "<html><body><p>Order Date: June 7, 2026</p></body></html>",
                encoding="utf-8",
            )
            with (
                patch.object(extractor, "API_KEY", "test-key"),
                patch.object(extractor, "extract_with_openai", return_value=_main_extraction_result()),
                patch.object(extractor, "should_run_is_gift_card", return_value=False),
                patch.object(extractor, "extract_purchase_date_rescue_with_openai") as rescue_call,
                patch.object(extractor, "_write_openai_log"),
                patch.object(extractor, "_write_tracking_log"),
                redirect_stdout(StringIO()),
            ):
                record = extractor.process_file(
                    email_path,
                    subject="Your order A1 is confirmed",
                    sender_name="Example Store",
                    email="orders@example.com",
                )

        rescue_call.assert_not_called()
        self.assertEqual(record["purchase_datetime"], "2026-06-07")
        self.assertEqual(record["purchase_datetime_source"], "explicit_order_date")
        self.assertFalse(record["purchase_datetime_rescue_attempted"])

    def test_missing_purchase_date_does_not_use_email_datetime(self) -> None:
        rows = [
            {
                "order_number": "A1",
                "purchase_datetime": "",
                "email_datetime": "2026-05-13T12:00:00Z",
                "company": "Example Store",
                "email_category": "Invoice",
            }
        ]

        with redirect_stdout(StringIO()):
            unify_purchase_dates_by_order(rows)

        self.assertEqual(rows[0]["purchase_datetime"], "")
        filename = build_convention_filename(rows[0])
        self.assertNotIn("2026-05-13", filename)
        self.assertEqual(filename, "DOC Example Store 0001 INVOICE.pdf")

    def test_missing_purchase_date_is_logged_for_review(self) -> None:
        rows = [
            {
                "order_number": "A1",
                "purchase_datetime": "",
                "company": "Example Store",
                "email_category": "Invoice",
            }
        ]
        log_lines: list[str] = []
        original_log = extractor.RL.log
        try:
            extractor.RL.log = lambda _segment, message: log_lines.append(str(message))
            with redirect_stdout(StringIO()):
                unify_purchase_dates_by_order(rows)
        finally:
            extractor.RL.log = original_log

        self.assertEqual(rows[0]["purchase_datetime"], "")
        self.assertTrue(any("could not resolve consolidated purchase date" in line for line in log_lines))

    def test_verified_order_date_from_same_order_is_shared(self) -> None:
        rows = [
            {
                "order_number": "B2",
                "purchase_datetime": "",
                "email_datetime": "2026-05-13T12:00:00Z",
                "company": "Example Store",
                "email_category": "Shipped",
            },
            {
                "order_number": "B2",
                "purchase_datetime": "2026-04-28",
                "email_datetime": "2026-05-14T12:00:00Z",
                "company": "Example Store",
                "email_category": "Invoice",
            },
        ]

        with redirect_stdout(StringIO()):
            unify_purchase_dates_by_order(rows)

        self.assertEqual(rows[0]["purchase_datetime"], "2026-04-28")
        self.assertEqual(rows[1]["purchase_datetime"], "2026-04-28")
        self.assertEqual(
            rows[0]["purchase_datetime_source"],
            "order_consensus:source=explicit_order_date;best_order_date_evidence:index=1",
        )
        self.assertEqual(rows[0]["purchase_datetime_confidence"], "medium")
        self.assertIn("2026-04-28", build_convention_filename(rows[0]))

    def test_invoice_order_date_beats_delivery_event_date_for_order(self) -> None:
        rows = [
            {
                "source_file": "delivered.html",
                "order_number": "228928",
                "purchase_datetime": "2026-02-26",
                "purchase_datetime_source": "event_date",
                "purchase_datetime_confidence": "high",
                "company": "Natasha Denona",
                "email_category": "Delivered",
            },
            {
                "source_file": "invoice.html",
                "order_number": "228928",
                "purchase_datetime": "2026-02-26",
                "purchase_datetime_source": "order_consensus:source=event_date;earliest_row_extracted:index=0",
                "purchase_datetime_confidence": "high",
                "company": "Natasha Denona",
                "email_category": "Invoice",
            },
        ]

        original_html_reader = extractor._candidate_html_text_for_record
        try:
            extractor._candidate_html_text_for_record = lambda record: (
                "Order No. 228928 February 20, 2026"
                if record.get("source_file") == "invoice.html"
                else ""
            )
            with redirect_stdout(StringIO()):
                unify_purchase_dates_by_order(rows)
        finally:
            extractor._candidate_html_text_for_record = original_html_reader

        self.assertEqual(rows[0]["purchase_datetime"], "2026-02-20")
        self.assertEqual(rows[1]["purchase_datetime"], "2026-02-20")
        self.assertEqual(
            rows[1]["purchase_datetime_source"],
            "order_consensus:source=html_text;best_order_date_evidence:index=1",
        )

    def test_labeled_receipt_date_formats_are_normalized(self) -> None:
        examples = {
            "Date Ordered: June 4th 2026": "2026-06-04",
            "Transaction Date 06.05.2026": "2026-06-05",
            "Payment Date: 6/6/26": "2026-06-06",
        }

        for text, expected in examples.items():
            with self.subTest(text=text):
                self.assertEqual(extractor.infer_order_date_from_labeled_text(text), expected)

    def test_invoice_date_near_order_number_can_precede_order(self) -> None:
        rows = [
            {
                "source_file": "invoice.html",
                "order_number": "VG-123456",
                "purchase_datetime": "",
                "company": "Example Store",
                "email_category": "Invoice",
            }
        ]

        original_html_reader = extractor._candidate_html_text_for_record
        try:
            extractor._candidate_html_text_for_record = lambda _record: (
                "Transaction Date June 7 2026 Billing details Order # VG-123456"
            )
            with redirect_stdout(StringIO()):
                unify_purchase_dates_by_order(rows)
        finally:
            extractor._candidate_html_text_for_record = original_html_reader

        self.assertEqual(rows[0]["purchase_datetime"], "2026-06-07")
        self.assertEqual(
            rows[0]["purchase_datetime_source"],
            "order_consensus:source=html_text;best_order_date_evidence:index=0",
        )


if __name__ == "__main__":
    unittest.main()
