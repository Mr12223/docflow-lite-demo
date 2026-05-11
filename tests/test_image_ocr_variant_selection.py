import os
import tempfile
import unittest
from unittest.mock import patch

from PIL import Image

from docflow.webapp.services import ocr, ocr_engines


def _build_eval(score, *, is_invoice, confidence="high", field_count=0, major_field_count=0):
    field_names = ["invoice_number"] if field_count else []
    return {
        "score": score,
        "is_invoice": is_invoice,
        "confidence": confidence if is_invoice else "none",
        "field_count": field_count,
        "major_field_count": major_field_count,
        "header_pair": False,
        "money_pair": False,
        "party_pair": False,
        "fields": {},
        "field_names": field_names,
        "invoice_fields": {
            "is_invoice": is_invoice,
            "confidence": confidence if is_invoice else "none",
            "field_count": field_count,
            "fields": {},
        },
    }


class ImageOcrVariantSelectionTests(unittest.TestCase):
    def setUp(self):
        handle = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        self.image_path = handle.name
        handle.close()
        Image.new("RGB", (32, 32), "white").save(self.image_path, format="PNG")

    def tearDown(self):
        try:
            os.remove(self.image_path)
        except FileNotFoundError:
            pass

    def test_process_image_ocr_prefers_best_preprocess_variant(self):
        rapid_results = [
            {
                "text": "rapid rgb weak",
                "meta": {"preprocess_variant": "rgb", "variant_name": "rgb"},
            },
            {
                "text": "rapid detail best",
                "meta": {"preprocess_variant": "detail_boost", "variant_name": "detail_boost"},
            },
        ]
        tesseract_results = [
            {
                "text": "tesseract gray medium",
                "meta": {
                    "preprocess_variant": "gray_autocontrast",
                    "variant_name": "gray_autocontrast",
                },
            }
        ]
        eval_map = {
            "rapid rgb weak": _build_eval(20, is_invoice=False),
            "rapid detail best": _build_eval(180, is_invoice=True, field_count=4, major_field_count=4),
            "tesseract gray medium": _build_eval(110, is_invoice=True, field_count=2, major_field_count=2),
        }
        merged_invoice_fields = {
            "is_invoice": True,
            "confidence": "high",
            "field_count": 1,
            "fields": {
                "invoice_number": {
                    "value": "12345678",
                    "confidence": "high",
                }
            },
        }

        with patch.object(ocr, "_get_image_ocr_order", return_value=["rapidocr", "tesseract"]), patch.object(
            ocr, "_is_image_ocr_cache_enabled", return_value=False
        ), patch.object(
            ocr, "_run_rapidocr_variants", return_value=rapid_results
        ), patch.object(
            ocr, "_run_tesseract_variants", return_value=tesseract_results
        ), patch.object(
            ocr, "_collect_rapidocr_tax_id_field_candidates", return_value=[]
        ), patch.object(
            ocr, "_collect_tesseract_tax_id_crop_candidates", return_value=[]
        ), patch.object(
            ocr, "_collect_tesseract_tax_id_box_references", return_value={}
        ), patch.object(
            ocr,
            "_evaluate_ocr_invoice_text",
            side_effect=lambda text: eval_map.get(text, _build_eval(0, is_invoice=False)),
        ), patch.object(
            ocr,
            "_merge_ocr_invoice_fields",
            return_value=(merged_invoice_fields, {"applied": False, "fields": {}}),
        ), patch.object(
            ocr, "_format_ocr_field_merge_summary", return_value=""
        ):
            result = ocr.process_image_ocr(self.image_path, "sample.png")

        self.assertTrue(result["success"])
        self.assertEqual(result["metadata"]["engine"], "RapidOCR")
        self.assertEqual(result["metadata"]["ocr_preprocess"]["preprocess_variant"], "detail_boost")
        self.assertIn("RapidOCR[detail_boost]", result["metadata"]["ocr_selection_summary"])
        self.assertEqual(result["statistics"]["invoice_fields"]["fields"]["invoice_number"]["value"], "12345678")

        candidate_summaries = result["metadata"]["ocr_candidates"]
        self.assertEqual(len(candidate_summaries), 3)
        self.assertEqual(candidate_summaries[0]["preprocess_variant"], "detail_boost")
        self.assertTrue(candidate_summaries[0]["selected"])

    def test_build_image_variant_entries_includes_detail_boost(self):
        with patch.dict(
            os.environ,
            {"DOCFLOW_RAPIDOCR_IMAGE_OCR_VARIANTS": "rgb,gray,detail,contrast,binary_170"},
            clear=False,
        ):
            base_image = Image.new("RGB", (48, 32), "white")
            entries = ocr_engines._build_image_variant_entries(base_image, "rapidocr")
            base_image.close()

        try:
            variant_names = [name for name, _, _ in entries]
            self.assertIn("detail_boost", variant_names)
            self.assertIn("contrast_sharp", variant_names)
            self.assertIn("binary_170", variant_names)
        finally:
            for _, variant_image, _ in entries:
                variant_image.close()


if __name__ == "__main__":
    unittest.main()
