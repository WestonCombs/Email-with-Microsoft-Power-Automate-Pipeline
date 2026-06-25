from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from proofOfDelivery.pod_data import sync_proof_of_delivery_records  # noqa: E402


def _write_json(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


class PodSyncAuditRecoveryTests(unittest.TestCase):
    def test_audited_pdf_recreates_pod_row_when_current_name_changed(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            pdf_dir = project_root / "email_contents" / "pdf"
            pdf_dir.mkdir(parents=True)
            pdf_path = pdf_dir / "DOC Correct Store 2026-01-02 TRACKING_INV_9999.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n")

            _write_json(
                json_dir / "results.json",
                [
                    {
                        "email_category": "Delivered",
                        "order_number": "123456",
                        "company": "Corret Store",
                        "purchase_datetime": "2026-01-02",
                        "email": "buyer@example.com",
                        "tracking_numbers": ["1Z9999"],
                    }
                ],
            )
            _write_json(
                json_dir / "tracking_pdf_audit.json",
                [
                    {
                        "company": "Correct Store",
                        "order_number": "123456",
                        "category": "Delivered",
                        "purchase_datetime": "2026-01-02",
                        "tracking_number": "1Z9999",
                        "timestamp_captured": "2026-06-16T18:00:00+00:00",
                        "path": str(pdf_path),
                        "filename": pdf_path.name,
                    }
                ],
            )

            records, changed = sync_proof_of_delivery_records(project_root)

            self.assertTrue(changed)
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]["email_category"], "POD")
            self.assertEqual(records[0]["source_file"], str(pdf_path.resolve()))
            self.assertEqual(records[0]["pod_tracking_number"], "1Z9999")

    def test_existing_pod_row_is_kept_while_pdf_still_exists(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            pdf_dir = project_root / "email_contents" / "pdf"
            pdf_dir.mkdir(parents=True)
            pdf_path = pdf_dir / "DOC Original Store 2026-01-02 TRACKING_INV_8888.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n")

            _write_json(
                json_dir / "results.json",
                [
                    {
                        "email_category": "Delivered",
                        "order_number": "123456",
                        "company": "Renamed Store",
                        "purchase_datetime": "2026-01-02",
                        "tracking_numbers": ["1Z8888"],
                    }
                ],
            )
            _write_json(
                json_dir / "proof_of_delivery.json",
                [
                    {
                        "email_category": "POD",
                        "order_number": "123456",
                        "company": "Original Store",
                        "source_file": str(pdf_path),
                        "source_file_link": pdf_path.resolve().as_uri(),
                        "tracking_numbers": ["1Z8888"],
                        "pod_tracking_number": "1Z8888",
                        "pod_generated_file_name": pdf_path.name,
                    }
                ],
            )

            records, changed = sync_proof_of_delivery_records(project_root)

            self.assertFalse(changed)
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]["source_file"], str(pdf_path.resolve()))


if __name__ == "__main__":
    unittest.main()
