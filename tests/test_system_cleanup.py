import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docflow.webapp.routes import common


class SystemCleanupTests(unittest.TestCase):
    def test_clear_upload_buffer_keeps_ocr_cache_and_active_files(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            upload_dir = Path(tmp_dir) / "uploads_temp"
            cache_dir = upload_dir / "ocr_cache"
            nested_dir = upload_dir / "nested"
            upload_dir.mkdir()
            cache_dir.mkdir()
            nested_dir.mkdir()

            stale_file = upload_dir / "old_upload.jpg"
            stale_file.write_bytes(b"old")
            nested_file = nested_dir / "tmp.png"
            nested_file.write_bytes(b"tmp")
            cache_file = cache_dir / "cached.json"
            cache_file.write_text("{}", encoding="utf-8")
            active_file = upload_dir / "active.pdf"
            active_file.write_bytes(b"active")

            with patch.object(common, "UPLOAD_FOLDER", str(upload_dir)), patch.object(
                common, "UPLOADS_DIR", upload_dir
            ), patch.object(common, "IMAGE_OCR_CACHE_DIR", cache_dir), patch.object(
                common, "_get_active_upload_paths", return_value={active_file.resolve()}
            ):
                result = common._clear_upload_buffer_files()

            self.assertEqual(result["deleted_files"], 2)
            self.assertEqual(result["deleted_bytes"], 6)
            self.assertEqual(result["skipped_active_files"], 1)
            self.assertFalse(stale_file.exists())
            self.assertFalse(nested_file.exists())
            self.assertFalse(nested_dir.exists())
            self.assertTrue(cache_file.exists())
            self.assertTrue(active_file.exists())


if __name__ == "__main__":
    unittest.main()
