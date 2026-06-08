import json
import tempfile
import unittest
from pathlib import Path

from app import app
from docflow.webapp.services import batch_jobs


class BatchJobsTests(unittest.TestCase):
    def tearDown(self):
        with batch_jobs.BATCH_TEST_LOCK:
            batch_jobs.BATCH_TEST_JOBS.clear()
            batch_jobs.BATCH_TEST_PROCESSES.clear()

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

    def test_batch_subprocess_env_disables_risky_ocr_parallelism(self):
        env = batch_jobs._build_batch_subprocess_env(
            {
                "DOCFLOW_RAPIDOCR_PREWARM": "1",
                "DOCFLOW_IMAGE_OCR_PARALLEL_ENGINES": "1",
                "OMP_NUM_THREADS": "8",
            }
        )

        self.assertEqual(env["DOCFLOW_BATCH_SAFE_IMAGE_OCR"], "1")
        self.assertEqual(env["DOCFLOW_BATCH_SAFE_PDF_OCR"], "1")
        self.assertEqual(env["DOCFLOW_RAPIDOCR_PREWARM"], "0")
        self.assertEqual(env["DOCFLOW_IMAGE_OCR_ORDER"], "rapidocr")
        self.assertEqual(env["DOCFLOW_IMAGE_OCR_PARALLEL_ENGINES"], "0")
        self.assertEqual(env["DOCFLOW_IMAGE_OCR_VARIANT_WORKERS"], "1")
        self.assertEqual(env["DOCFLOW_IMAGE_OCR_TESSERACT_TAX_ID_REFINEMENT"], "0")
        self.assertEqual(env["OMP_NUM_THREADS"], "1")

    def test_batch_report_dir_is_unique_for_job(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            report_dir = batch_jobs._build_batch_report_dir("job/../123", Path(tmp_dir))

        self.assertEqual(report_dir.parent, Path(tmp_dir))
        self.assertRegex(report_dir.name, r"^batch_test_\d{8}_\d{6}_job123$")

    def test_batch_command_uses_explicit_report_dir(self):
        report_dir = Path("reports") / "batch_test_job123"
        cmd = batch_jobs._build_batch_command(
            ["test_documents"],
            keywords=False,
            strict=True,
            pdf_mode="fast",
            report_dir=report_dir,
        )

        self.assertIn("--report-dir", cmd)
        self.assertEqual(cmd[cmd.index("--report-dir") + 1], str(report_dir))
        self.assertIn("--strict", cmd)
        self.assertEqual(cmd[cmd.index("--pdf-mode") + 1], "fast")

    def test_run_batch_tests_reuses_active_job(self):
        with batch_jobs.BATCH_TEST_LOCK:
            batch_jobs.BATCH_TEST_JOBS["active123"] = {
                "job_id": "active123",
                "state": "running",
                "total": 51,
                "suites": ["test_documents"],
                "pdf_mode": "fast",
            }

        client = app.test_client()
        response = client.post(
            "/run-batch-tests",
            json={"suites": ["test_documents_edge_cases"], "pdf_mode": "fast"},
        )
        payload = response.get_json()

        self.assertEqual(response.status_code, 200)
        self.assertTrue(payload["success"])
        self.assertTrue(payload["already_running"])
        self.assertEqual(payload["job_id"], "active123")
        with batch_jobs.BATCH_TEST_LOCK:
            self.assertEqual(list(batch_jobs.BATCH_TEST_JOBS), ["active123"])


if __name__ == "__main__":
    unittest.main()
