import os
import tempfile
import unittest
from unittest.mock import patch

from PIL import Image

from docflow.webapp.services import ocr_engines


class ImageOcrCacheProfileTests(unittest.TestCase):
    def setUp(self):
        handle = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        self.image_path = handle.name
        handle.close()
        Image.new("RGB", (48, 32), "white").save(self.image_path, format="PNG")

    def tearDown(self):
        try:
            os.remove(self.image_path)
        except FileNotFoundError:
            pass

    def test_cache_key_changes_when_resize_policy_changes(self):
        base_env = {
            "DOCFLOW_IMAGE_OCR_TARGET_LONG_EDGE": "1200",
            "DOCFLOW_RAPIDOCR_TARGET_LONG_EDGE": "1200",
            "DOCFLOW_TESSERACT_MAX_LONG_EDGE": "1600",
        }
        tuned_env = {
            "DOCFLOW_IMAGE_OCR_TARGET_LONG_EDGE": "1320",
            "DOCFLOW_RAPIDOCR_TARGET_LONG_EDGE": "1440",
            "DOCFLOW_TESSERACT_MAX_LONG_EDGE": "1180",
        }

        with patch.dict(os.environ, base_env, clear=False):
            key_before, _, profile_before = ocr_engines._build_image_ocr_cache_key(self.image_path)

        with patch.dict(os.environ, tuned_env, clear=False):
            key_after, _, profile_after = ocr_engines._build_image_ocr_cache_key(self.image_path)

        self.assertNotEqual(key_before, key_after)
        self.assertEqual(profile_before["target_long_edge"], 1200)
        self.assertEqual(profile_after["target_long_edge"], 1320)
        self.assertEqual(profile_before["provider_resize_policies"]["rapidocr"]["target_long_edge"], 1200)
        self.assertEqual(profile_after["provider_resize_policies"]["rapidocr"]["target_long_edge"], 1440)
        self.assertEqual(profile_before["provider_resize_policies"]["tesseract"]["max_long_edge"], 1600)
        self.assertEqual(profile_after["provider_resize_policies"]["tesseract"]["max_long_edge"], 1180)

    def test_clear_image_ocr_cache_clears_disk_and_memory(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            cache_dir = os.path.join(tmp_dir, "ocr_cache")
            os.makedirs(cache_dir, exist_ok=True)
            with open(os.path.join(cache_dir, "a.json"), "w", encoding="utf-8") as f:
                f.write("{}")
            os.makedirs(os.path.join(cache_dir, "nested"), exist_ok=True)
            with open(os.path.join(cache_dir, "nested", "b.json"), "w", encoding="utf-8") as f:
                f.write("{}")

            with patch.object(ocr_engines, "IMAGE_OCR_CACHE_DIR", ocr_engines.Path(cache_dir)):
                ocr_engines.IMAGE_OCR_MEMORY_CACHE["x"] = {"result": {}}
                result = ocr_engines.clear_image_ocr_cache()

            self.assertEqual(result["deleted_files"], 2)
            self.assertEqual(result["cleared_memory_items"], 1)
            self.assertTrue(os.path.isdir(cache_dir))
            self.assertEqual(os.listdir(cache_dir), [])
            self.assertEqual(ocr_engines.IMAGE_OCR_MEMORY_CACHE, {})


if __name__ == "__main__":
    unittest.main()
