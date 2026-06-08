import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from scripts import run_batch_tests


class BatchScriptTests(unittest.TestCase):
    def test_safe_image_ocr_mode_returns_stable_result_without_processor(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            image_path = Path(tmp_dir) / "sample.png"
            image_path.write_bytes(b"not a real image")

            with patch.dict(os.environ, {"DOCFLOW_BATCH_SAFE_IMAGE_OCR": "1"}):
                result = run_batch_tests.run_single_case(
                    processor=None,
                    suite_name="test_documents",
                    file_path=image_path,
                    extract_keywords=False,
                    pdf_mode="fast",
                )

        self.assertTrue(result["success"])
        self.assertEqual(result["ocr_engine"], "BatchSafeOCR")
        self.assertEqual(result["error"], "")


if __name__ == "__main__":
    unittest.main()
