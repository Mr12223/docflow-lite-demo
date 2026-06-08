import json
import tempfile
import unittest
from pathlib import Path

from docflow.webapp.services import batch_jobs


class BatchJobsTests(unittest.TestCase):
    def test_empty_report_dir_does_not_publish_missing_report_links(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            report_dir = Path(tmp_dir) / "batch_test_empty"
            report_dir.mkdir()

            summary, records, failed_cases, unexpected_cases, report_urls = (
                batch_jobs._build_report_payload(report_dir)
            )

        self.assertEqual(summary, {})
        self.assertEqual(records, [])
        self.assertEqual(failed_cases, [])
        self.assertEqual(unexpected_cases, [])
        self.assertEqual(report_urls, {})

    def test_report_payload_only_links_existing_report_files(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            report_dir = Path(tmp_dir) / "batch_test_partial"
            report_dir.mkdir()
            (report_dir / "results.json").write_text(
                json.dumps({"summary": {"total": 1}, "records": []}),
                encoding="utf-8",
            )

            summary, _records, _failed_cases, _unexpected_cases, report_urls = (
                batch_jobs._build_report_payload(report_dir)
            )

        self.assertEqual(summary, {"total": 1})
        self.assertEqual(report_urls, {"json": "/reports/batch_test_partial/results.json"})

    def test_nonzero_return_code_is_failure_even_when_summary_exists(self):
        self.assertTrue(batch_jobs._is_batch_job_success(0, False, {"total": 1}))
        self.assertFalse(batch_jobs._is_batch_job_success(2, False, {"total": 1}))

        message = batch_jobs._build_batch_failure_message(
            cancelled=False,
            return_code=2,
            summary={"total": 1},
            failed_cases=[],
            unexpected_cases=[{"filename": "unexpected.pdf"}],
            report_dir=Path("reports/batch_test_strict"),
        )

        self.assertIn("unexpected.pdf", message)


if __name__ == "__main__":
    unittest.main()
