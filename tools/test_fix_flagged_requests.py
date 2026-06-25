from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from shared.fix_flagged import (  # noqa: E402
    REQUEST_TYPE_EXCEL_FLAGGED,
    active_requests,
    load_requests,
    load_results,
    mark_record_excel_flagged,
    resolve_pod_review_with_manual_scan,
    save_record_updates,
    skip_record_eternally,
    sync_fix_flagged_requests,
)
from shared.excel_user_edits import record_identity  # noqa: E402
from shared.order_store import stable_record_id  # noqa: E402
from createExcelDocument.excel_user_edit_sync import _record_flagged_edit  # noqa: E402
from proofOfDelivery.pod_data import pod_status_viewer_rows  # noqa: E402


class FixFlaggedRequestTests(unittest.TestCase):
    def test_invoice_total_mismatch_creates_request_and_resolve_sticks(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            record = {
                "source_file_link": "file:///C:/tmp/order.pdf",
                "email_category": "Invoice",
                "company": "Example",
                "order_number": "1234",
                "subtotal_amount": 20.0,
                "tax_paid": 2.0,
                "gift_card_amount": 0.0,
                "total_amount_paid": 30.0,
            }
            (json_dir / "results.json").write_text(
                json.dumps([record], indent=2),
                encoding="utf-8",
            )

            requests = sync_fix_flagged_requests(project_root)
            active = active_requests(requests)
            self.assertEqual(len(active), 1)
            self.assertEqual(active[0]["type"], "invoice_totals_ambiguous")

            save_record_updates(
                project_root,
                record_id=record_identity(record),
                updates={"total_amount_paid": "22.00"},
                request_id=active[0]["id"],
                status="resolved",
            )
            requests_after = load_requests(project_root)
            self.assertEqual(active_requests(requests_after), [])

            rows = json.loads((json_dir / "results.json").read_text(encoding="utf-8"))
            self.assertEqual(rows[0]["total_amount_paid"], 22.0)

    def test_corrected_invoice_total_auto_resolves_request(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            record = {
                "source_file_link": "file:///C:/tmp/order.pdf",
                "email_category": "Invoice",
                "company": "Example",
                "order_number": "1234",
                "subtotal_amount": 20.0,
                "tax_paid": 2.0,
                "gift_card_amount": 0.0,
                "total_amount_paid": 30.0,
            }
            (json_dir / "results.json").write_text(
                json.dumps([record], indent=2),
                encoding="utf-8",
            )

            requests = sync_fix_flagged_requests(project_root)
            active = active_requests(requests)
            self.assertEqual(len(active), 1)

            save_record_updates(
                project_root,
                record_id=record_identity(record),
                updates={"total_amount_paid": "22.00"},
            )

            requests_after = load_requests(project_root)
            self.assertEqual(active_requests(requests_after), [])

    def test_clearing_fix_flagged_field_resets_manual_override(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            record = {
                "source_file_link": "file:///C:/tmp/order.pdf",
                "email_category": "Invoice",
                "company": "Original Co",
                "order_number": "1234",
                "total_amount_paid": 10.0,
            }
            (json_dir / "results.json").write_text(
                json.dumps([record], indent=2),
                encoding="utf-8",
            )

            save_record_updates(
                project_root,
                record_id=record_identity(record),
                updates={"company": "Edited Co", "total_amount_paid": "22.00"},
            )
            edited = load_results(project_root)[0]
            self.assertEqual(edited["company"], "Edited Co")
            self.assertEqual(edited["total_amount_paid"], 22.0)
            self.assertTrue(edited["modified_company"])
            self.assertTrue(edited["modified_total_amount_paid"])

            save_record_updates(
                project_root,
                record_id=record_identity(edited),
                updates={"company": "", "total_amount_paid": ""},
            )
            reset = load_results(project_root)[0]
            self.assertEqual(reset["company"], "Original Co")
            self.assertEqual(reset["total_amount_paid"], 10.0)
            self.assertNotIn("modified_company", reset)
            self.assertNotIn("modified_total_amount_paid", reset)

    def test_clearing_company_restores_llm_original_not_latest_modified_value(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            record = {
                "source_file_link": "file:///C:/tmp/order.pdf",
                "email_category": "Invoice",
                "company": "/",
                "original_llm_obtained_company": "Original Co",
                "llm_obtained_company": "Original Co",
                "modified_company": True,
                "order_number": "1234",
            }
            (json_dir / "results.json").write_text(
                json.dumps([record], indent=2),
                encoding="utf-8",
            )

            save_record_updates(
                project_root,
                record_id=record_identity(record),
                updates={"company": ""},
            )

            reset = load_results(project_root)[0]
            self.assertEqual(reset["company"], "Original Co")
            self.assertNotIn("modified_company", reset)

    def test_invalid_category_update_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            record = {
                "source_file_link": "file:///C:/tmp/order.pdf",
                "email_category": "Invoice",
                "company": "Example",
                "order_number": "1234",
            }
            (json_dir / "results.json").write_text(
                json.dumps([record], indent=2),
                encoding="utf-8",
            )

            with self.assertRaises(ValueError):
                save_record_updates(
                    project_root,
                    record_id=record_identity(record),
                    updates={"email_category": "Not A Real Category"},
                )

    def test_mark_record_excel_flagged_sets_plain_marker_and_request(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            record = {
                "source_file_link": "file:///C:/tmp/order.pdf",
                "email_category": "Invoice",
                "company": "Example",
                "order_number": "1234",
            }
            (json_dir / "results.json").write_text(
                json.dumps([record], indent=2),
                encoding="utf-8",
            )

            mark_record_excel_flagged(
                project_root,
                record_id=record_identity(record),
                flagged=True,
            )
            flagged = load_results(project_root)[0]
            self.assertIs(flagged["excel_flagged"], True)
            self.assertNotIn("excel_active", flagged)
            self.assertNotIn("modified_excel_flagged", flagged)
            active = active_requests(load_requests(project_root))
            self.assertEqual(len(active), 1)
            self.assertEqual(active[0]["type"], REQUEST_TYPE_EXCEL_FLAGGED)

    def test_excel_flagged_edit_bridge_toggles_marker_and_request(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            record = {
                "source_file_link": "file:///C:/tmp/order.pdf",
                "email_category": "Invoice",
                "company": "Example",
                "order_number": "1234",
            }
            (json_dir / "results.json").write_text(
                json.dumps([record], indent=2),
                encoding="utf-8",
            )
            stored = load_results(project_root)[0]
            record_id = str(stored["_record_id"])

            checked = _record_flagged_edit(
                project_root,
                {"record_id": record_id, "value": "True"},
            )
            self.assertEqual(checked["display_value"], "True")
            self.assertIs(load_results(project_root)[0]["excel_flagged"], True)
            self.assertEqual(active_requests(load_requests(project_root))[0]["type"], REQUEST_TYPE_EXCEL_FLAGGED)

            unchecked = _record_flagged_edit(
                project_root,
                {"record_id": record_id, "value": "False"},
            )
            self.assertEqual(unchecked["display_value"], "")
            self.assertNotIn("excel_flagged", load_results(project_root)[0])
            self.assertNotIn("excel_active", load_results(project_root)[0])
            self.assertEqual(active_requests(load_requests(project_root)), [])

    def test_excel_flagged_edit_bridge_updates_pod_row_by_workbook_id(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            (json_dir / "results.json").write_text(
                json.dumps(
                    [
                        {
                            "source_file_link": "file:///C:/tmp/base.pdf",
                            "email_category": "Delivered",
                            "company": "Example",
                            "order_number": "1234",
                        }
                    ],
                    indent=2,
                ),
                encoding="utf-8",
            )
            pod_uri = "file:///C:/tmp/DOC%20Example%202026-01-02%20TRACKING_INV_1234.pdf"
            pod_record = {
                "source_file_link": pod_uri,
                "email_category": "POD",
                "company": "Example",
                "order_number": "1234",
            }
            (json_dir / "proof_of_delivery.json").write_text(
                json.dumps([pod_record], indent=2),
                encoding="utf-8",
            )

            summary = _record_flagged_edit(
                project_root,
                {
                    "record_id": stable_record_id(pod_record),
                    "source_uri": pod_uri,
                    "order_number": "1234",
                    "value": "True",
                },
            )

            self.assertEqual(summary["display_value"], "True")
            pod_rows = json.loads((json_dir / "proof_of_delivery.json").read_text(encoding="utf-8"))
            self.assertIs(pod_rows[0]["excel_flagged"], True)
            self.assertEqual(pod_rows[0]["_record_id"], stable_record_id(pod_record))
            result_rows = load_results(project_root)
            self.assertNotIn("excel_flagged", result_rows[0])
            self.assertIn("proof_of_delivery.json", summary["changed_files"][0])

    def test_resolving_excel_flagged_request_clears_marker(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            record = {
                "source_file_link": "file:///C:/tmp/order.pdf",
                "email_category": "Invoice",
                "company": "Example",
                "order_number": "1234",
                "excel_flagged": True,
            }
            (json_dir / "results.json").write_text(
                json.dumps([record], indent=2),
                encoding="utf-8",
            )

            request = active_requests(sync_fix_flagged_requests(project_root))[0]
            self.assertEqual(request["type"], REQUEST_TYPE_EXCEL_FLAGGED)
            save_record_updates(
                project_root,
                record_id=record_identity(load_results(project_root)[0]),
                updates={},
                request_id=request["id"],
                status="resolved",
            )

            self.assertNotIn("excel_flagged", load_results(project_root)[0])
            self.assertEqual(active_requests(load_requests(project_root)), [])

    def test_skip_eternally_resolves_pod_review_row_and_request(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            (json_dir / "results.json").write_text(
                json.dumps(
                    [
                        {
                            "source_file_link": "file:///C:/tmp/order.pdf",
                            "email_category": "Shipped",
                            "company": "Example",
                            "order_number": "1234",
                            "tracking_number": "1Z999",
                            "tracking_numbers": ["1Z999"],
                        }
                    ],
                    indent=2,
                ),
                encoding="utf-8",
            )
            pod_record = {
                "source_file_link": "file:///C:/tmp/review.pdf",
                "email_category": "POD",
                "company": "Example",
                "order_number": "1234",
                "pod_review_required": True,
                "pod_review_status": "active",
                "pod_tracking_number": "1Z999",
            }
            (json_dir / "proof_of_delivery.json").write_text(
                json.dumps([pod_record], indent=2),
                encoding="utf-8",
            )

            request = active_requests(sync_fix_flagged_requests(project_root))[0]
            result = skip_record_eternally(
                project_root,
                record_id=record_identity(pod_record),
                request_id=request["id"],
            )

            self.assertEqual(result["removed_records"], 0)
            self.assertEqual(result["resolved_records"], 1)
            pod_rows = json.loads((json_dir / "proof_of_delivery.json").read_text(encoding="utf-8"))
            self.assertEqual(len(pod_rows), 1)
            self.assertFalse(pod_rows[0]["pod_review_required"])
            self.assertEqual(pod_rows[0]["pod_review_status"], "resolved")
            self.assertTrue(pod_rows[0]["pod_review_skipped_eternally"])
            self.assertEqual(active_requests(load_requests(project_root)), [])
            viewer_rows = pod_status_viewer_rows(project_root)
            self.assertEqual(viewer_rows[0]["pod_review_status"], "resolved")
            self.assertTrue(viewer_rows[0]["pod_review_resolved"])

    def test_manual_rescan_replaces_existing_pod_review_without_duplicate(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            pdf_dir = project_root / "email_contents" / "pdf"
            json_dir.mkdir(parents=True)
            pdf_dir.mkdir(parents=True)
            (json_dir / "results.json").write_text(
                json.dumps(
                    [
                        {
                            "source_file_link": "file:///C:/tmp/order.pdf",
                            "email_category": "Shipped",
                            "company": "Example",
                            "order_number": "1234",
                            "purchase_datetime": "2026-01-02",
                            "tracking_numbers": ["1Z9999"],
                        }
                    ],
                    indent=2,
                ),
                encoding="utf-8",
            )
            old_review_pdf = pdf_dir / "old_review.pdf"
            old_review_pdf.write_bytes(b"old")
            scan_pdf = pdf_dir / "DOC Example 2026-01-02 TRACKING_INV_9999_manual_rescan.pdf"
            scan_pdf.write_bytes(b"new")
            pod_record = {
                "source_file": str(old_review_pdf),
                "source_file_link": old_review_pdf.as_uri(),
                "email_category": "POD",
                "company": "Example",
                "order_number": "1234",
                "pod_source_purchase_datetime": "2026-01-02",
                "pod_review_required": True,
                "pod_review_status": "active",
                "pod_tracking_number": "1Z9999",
                "pod_carrier": "UPS",
                "excel_flagged": True,
            }
            (json_dir / "proof_of_delivery.json").write_text(
                json.dumps([pod_record], indent=2),
                encoding="utf-8",
            )

            request = active_requests(sync_fix_flagged_requests(project_root))[0]
            result = resolve_pod_review_with_manual_scan(
                project_root,
                record_id=record_identity(pod_record),
                scanned_pdf_path=scan_pdf,
                request_id=request["id"],
            )

            pod_rows = json.loads((json_dir / "proof_of_delivery.json").read_text(encoding="utf-8"))
            self.assertEqual(len(pod_rows), 1)
            self.assertFalse(pod_rows[0]["pod_review_required"])
            self.assertEqual(pod_rows[0]["pod_review_status"], "resolved")
            self.assertNotIn("excel_flagged", pod_rows[0])
            final_pdf = Path(str(result["final_pdf"]))
            self.assertTrue(final_pdf.is_file())
            self.assertEqual(final_pdf.name, "DOC Example 2026-01-02 TRACKING_INV_9999.pdf")
            self.assertFalse(scan_pdf.exists())
            self.assertEqual(active_requests(load_requests(project_root)), [])


if __name__ == "__main__":
    unittest.main()
