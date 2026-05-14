import os
import threading
import time
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


def _build_invoice_eval_with_fields(fields):
    return {
        "score": 260,
        "is_invoice": True,
        "confidence": "high",
        "field_count": len(fields),
        "major_field_count": len([name for name in fields if name in {"invoice_code", "invoice_number", "invoice_date", "amount", "total", "buyer_name", "seller_name", "invoice_type"}]),
        "header_pair": "invoice_code" in fields and "invoice_number" in fields,
        "money_pair": "amount" in fields and "total" in fields,
        "party_pair": "buyer_name" in fields and "seller_name" in fields,
        "fields": {name: {"value": value, "confidence": "high"} for name, value in fields.items()},
        "field_names": sorted(fields.keys()),
        "invoice_fields": {
            "is_invoice": True,
            "confidence": "high",
            "field_count": len(fields),
            "fields": {name: {"value": value, "confidence": "high"} for name, value in fields.items()},
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

    def test_process_image_ocr_force_reprocess_bypasses_read_cache_and_saves_new_result(self):
        rapid_results = [
            {
                "text": "fresh rapid text",
                "meta": {"preprocess_variant": "rgb", "variant_name": "rgb"},
            }
        ]
        eval_map = {
            "fresh rapid text": _build_eval(90, is_invoice=False),
        }

        with patch.object(ocr, "_get_image_ocr_order", return_value=["rapidocr"]), patch.object(
            ocr, "_is_image_ocr_cache_enabled", return_value=True
        ), patch.object(
            ocr, "_build_image_ocr_cache_key", return_value=("cache-key", "file-sha", {"profile": "test"})
        ), patch.object(
            ocr, "_load_image_ocr_cache", return_value={"result": {"text": "stale cached text"}}
        ) as load_cache, patch.object(
            ocr, "_run_rapidocr_variants", return_value=rapid_results
        ) as run_rapidocr, patch.object(
            ocr,
            "_evaluate_ocr_invoice_text",
            side_effect=lambda text: eval_map.get(text, _build_eval(0, is_invoice=False)),
        ), patch.object(
            ocr, "_save_image_ocr_cache"
        ) as save_cache:
            result = ocr.process_image_ocr(self.image_path, "sample.png", force_reprocess=True)

        self.assertTrue(result["success"])
        self.assertEqual(result["text"], "fresh rapid text")
        self.assertFalse(result["metadata"]["image_ocr_cache"]["hit"])
        self.assertTrue(result["metadata"]["image_ocr_cache"]["bypassed"])
        load_cache.assert_not_called()
        run_rapidocr.assert_called_once_with(self.image_path)
        save_cache.assert_called_once()

    def test_process_image_ocr_runs_configured_engines_in_parallel(self):
        starts = []
        lock = threading.Lock()
        release = threading.Event()

        def make_runner(text):
            def runner(_image_path):
                with lock:
                    starts.append((text, time.time()))
                    if len(starts) >= 2:
                        release.set()
                release.wait(timeout=2)
                return [{"text": text, "meta": {"preprocess_variant": "gray_autocontrast"}}]

            return runner

        eval_map = {
            "rapid text": _build_eval(80, is_invoice=False),
            "tesseract text": _build_eval(70, is_invoice=False),
        }

        with patch.dict(
            os.environ,
            {
                "DOCFLOW_IMAGE_OCR_PARALLEL_ENGINES": "1",
                "DOCFLOW_IMAGE_OCR_ENGINE_WORKERS": "2",
            },
            clear=False,
        ), patch.object(
            ocr, "_get_image_ocr_order", return_value=["rapidocr", "tesseract"]
        ), patch.object(
            ocr, "_is_image_ocr_cache_enabled", return_value=False
        ), patch.object(
            ocr, "_run_rapidocr_variants", side_effect=make_runner("rapid text")
        ), patch.object(
            ocr, "_run_tesseract_variants", side_effect=make_runner("tesseract text")
        ), patch.object(
            ocr,
            "_evaluate_ocr_invoice_text",
            side_effect=lambda text: eval_map.get(text, _build_eval(0, is_invoice=False)),
        ):
            result = ocr.process_image_ocr(self.image_path, "sample.png")

        self.assertTrue(result["success"])
        self.assertEqual(len(starts), 2)
        self.assertLess(abs(starts[1][1] - starts[0][1]), 1.0)
        self.assertEqual([item["engine_key"] for item in result["metadata"]["ocr_candidates"]], ["rapidocr", "tesseract"])

    def test_process_image_ocr_skips_invoice_refinement_when_tax_ids_are_complete(self):
        fields = {
            "invoice_code": "3200164130",
            "invoice_number": "13638082",
            "invoice_date": "2017-04-20",
            "amount": "20984.62",
            "total": "24552.00",
            "buyer_name": "中国科学院自动化研究所",
            "seller_name": "福尔哈贝传动技术(太仓)有限公司",
            "buyer_tax_id": "110108400010945",
            "seller_tax_id": "91320585554626259G",
            "invoice_type": "增值税专用发票",
        }
        invoice_eval = _build_invoice_eval_with_fields(fields)

        with patch.object(ocr, "_get_image_ocr_order", return_value=["rapidocr"]), patch.object(
            ocr, "_is_image_ocr_cache_enabled", return_value=False
        ), patch.object(
            ocr, "_run_rapidocr_variants", return_value=[{"text": "complete invoice text", "meta": {"preprocess_variant": "gray_autocontrast"}}]
        ), patch.object(
            ocr, "_evaluate_ocr_invoice_text", return_value=invoice_eval
        ), patch.object(
            ocr, "_collect_rapidocr_tax_id_field_candidates", return_value=[]
        ) as rapid_refine, patch.object(
            ocr, "_collect_tesseract_tax_id_crop_candidates", return_value=[]
        ) as tesseract_refine, patch.object(
            ocr, "_collect_tesseract_tax_id_box_references", return_value={}
        ) as box_ref:
            result = ocr.process_image_ocr(self.image_path, "sample.png")

        self.assertTrue(result["success"])
        self.assertFalse(result["metadata"]["ocr_invoice_refinement_applied"])
        self.assertEqual(result["metadata"]["ocr_invoice_refinement_reason"], "tax_id_complete")
        rapid_refine.assert_not_called()
        tesseract_refine.assert_not_called()
        box_ref.assert_not_called()

    def test_process_image_ocr_runs_invoice_refinement_when_tax_id_missing(self):
        fields = {
            "invoice_code": "3200164130",
            "invoice_number": "13638082",
            "buyer_tax_id": "110108400010945",
            "amount": "20984.62",
            "total": "24552.00",
        }
        invoice_eval = _build_invoice_eval_with_fields(fields)

        with patch.dict(os.environ, {}, clear=True), patch.object(
            ocr, "_get_image_ocr_order", return_value=["rapidocr"]
        ), patch.object(
            ocr, "_is_image_ocr_cache_enabled", return_value=False
        ), patch.object(
            ocr, "_run_rapidocr_variants", return_value=[{"text": "missing seller tax id", "meta": {"preprocess_variant": "gray_autocontrast"}}]
        ), patch.object(
            ocr, "_evaluate_ocr_invoice_text", return_value=invoice_eval
        ), patch.object(
            ocr, "_collect_rapidocr_tax_id_field_candidates", return_value=[]
        ) as rapid_refine, patch.object(
            ocr, "_collect_tesseract_tax_id_crop_candidates", return_value=[]
        ) as tesseract_refine, patch.object(
            ocr, "_collect_tesseract_tax_id_box_references", return_value={}
        ) as box_ref:
            result = ocr.process_image_ocr(self.image_path, "sample.png")

        self.assertTrue(result["success"])
        self.assertTrue(result["metadata"]["ocr_invoice_refinement_applied"])
        self.assertEqual(result["metadata"]["ocr_invoice_refinement_reason"], "missing_seller_tax_id")
        rapid_refine.assert_called_once_with(self.image_path, target_fields=("seller_tax_id",))
        tesseract_refine.assert_not_called()
        box_ref.assert_not_called()

    def test_process_image_ocr_skips_duplicate_tax_id_refinement_by_default(self):
        fields = {
            "invoice_code": "3200164130",
            "invoice_number": "13638082",
            "buyer_tax_id": "91340000711771143J",
            "seller_tax_id": "91340000711771143J",
            "amount": "1708.55",
            "total": "1999.00",
        }
        invoice_eval = _build_invoice_eval_with_fields(fields)

        with patch.dict(os.environ, {}, clear=True), patch.object(
            ocr, "_get_image_ocr_order", return_value=["rapidocr"]
        ), patch.object(
            ocr, "_is_image_ocr_cache_enabled", return_value=False
        ), patch.object(
            ocr, "_run_rapidocr_variants", return_value=[{"text": "duplicate tax id", "meta": {"preprocess_variant": "gray_autocontrast"}}]
        ), patch.object(
            ocr, "_evaluate_ocr_invoice_text", return_value=invoice_eval
        ), patch.object(
            ocr, "_collect_rapidocr_tax_id_field_candidates", return_value=[]
        ) as rapid_refine:
            result = ocr.process_image_ocr(self.image_path, "sample.png")

        self.assertTrue(result["success"])
        self.assertFalse(result["metadata"]["ocr_invoice_refinement_applied"])
        self.assertEqual(result["metadata"]["ocr_invoice_refinement_reason"], "duplicate_tax_id_skipped")
        rapid_refine.assert_not_called()

    def test_process_image_ocr_can_enable_duplicate_tax_id_refinement(self):
        fields = {
            "invoice_code": "3200164130",
            "invoice_number": "13638082",
            "buyer_tax_id": "91340000711771143J",
            "seller_tax_id": "91340000711771143J",
        }
        invoice_eval = _build_invoice_eval_with_fields(fields)

        with patch.dict(
            os.environ,
            {"DOCFLOW_IMAGE_OCR_DUPLICATE_TAX_ID_REFINEMENT": "1"},
            clear=True,
        ), patch.object(
            ocr, "_get_image_ocr_order", return_value=["rapidocr"]
        ), patch.object(
            ocr, "_is_image_ocr_cache_enabled", return_value=False
        ), patch.object(
            ocr, "_run_rapidocr_variants", return_value=[{"text": "duplicate tax id", "meta": {"preprocess_variant": "gray_autocontrast"}}]
        ), patch.object(
            ocr, "_evaluate_ocr_invoice_text", return_value=invoice_eval
        ), patch.object(
            ocr, "_collect_rapidocr_tax_id_field_candidates", return_value=[]
        ) as rapid_refine:
            result = ocr.process_image_ocr(self.image_path, "sample.png")

        self.assertTrue(result["metadata"]["ocr_invoice_refinement_applied"])
        self.assertEqual(result["metadata"]["ocr_invoice_refinement_reason"], "duplicate_tax_id")
        rapid_refine.assert_called_once_with(self.image_path, target_fields=("buyer_tax_id", "seller_tax_id"))

    def test_process_image_ocr_can_enable_tesseract_tax_id_refinement(self):
        fields = {
            "invoice_code": "3200164130",
            "invoice_number": "13638082",
            "buyer_tax_id": "110108400010945",
        }
        invoice_eval = _build_invoice_eval_with_fields(fields)

        with patch.dict(
            os.environ,
            {"DOCFLOW_IMAGE_OCR_TESSERACT_TAX_ID_REFINEMENT": "1"},
            clear=True,
        ), patch.object(
            ocr, "_get_image_ocr_order", return_value=["rapidocr"]
        ), patch.object(
            ocr, "_is_image_ocr_cache_enabled", return_value=False
        ), patch.object(
            ocr, "_run_rapidocr_variants", return_value=[{"text": "missing seller tax id", "meta": {"preprocess_variant": "gray_autocontrast"}}]
        ), patch.object(
            ocr, "_evaluate_ocr_invoice_text", return_value=invoice_eval
        ), patch.object(
            ocr, "_collect_rapidocr_tax_id_field_candidates", return_value=[]
        ), patch.object(
            ocr, "_collect_tesseract_tax_id_crop_candidates", return_value=[]
        ) as tesseract_refine, patch.object(
            ocr, "_collect_tesseract_tax_id_box_references", return_value={}
        ) as box_ref:
            result = ocr.process_image_ocr(self.image_path, "sample.png")

        self.assertTrue(result["metadata"]["ocr_invoice_refinement_applied"])
        self.assertEqual(result["metadata"]["ocr_invoice_refinement_reason"], "missing_seller_tax_id")
        tesseract_refine.assert_called_once_with(self.image_path)
        box_ref.assert_called_once_with(self.image_path)

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

    def test_default_fast_profile_keeps_rapidocr_variant_count_small(self):
        with patch.dict(os.environ, {}, clear=True):
            base_image = Image.new("RGB", (48, 32), "white")
            entries = ocr_engines._build_image_variant_entries(base_image, "rapidocr")
            base_image.close()

        try:
            self.assertEqual([name for name, _, _ in entries], ["gray_autocontrast", "detail_boost"])
        finally:
            for _, variant_image, _ in entries:
                variant_image.close()

    def test_accurate_profile_keeps_full_rapidocr_variants_available(self):
        with patch.dict(os.environ, {"DOCFLOW_IMAGE_OCR_SPEED_PROFILE": "accurate"}, clear=True):
            base_image = Image.new("RGB", (48, 32), "white")
            entries = ocr_engines._build_image_variant_entries(base_image, "rapidocr")
            base_image.close()

        try:
            variant_names = [name for name, _, _ in entries]
            self.assertIn("rgb", variant_names)
            self.assertIn("contrast_sharp", variant_names)
            self.assertIn("binary_170", variant_names)
            self.assertIn("binary_190", variant_names)
        finally:
            for _, variant_image, _ in entries:
                variant_image.close()

    def test_tesseract_variants_can_run_in_parallel_and_keep_result_order(self):
        starts = []
        lock = threading.Lock()
        release = threading.Event()

        def fake_runner(image):
            with lock:
                starts.append(time.time())
                if len(starts) >= 2:
                    release.set()
            release.wait(timeout=2)
            return f"text {len(starts)}", {}

        with patch.dict(
            os.environ,
            {
                "DOCFLOW_TESSERACT_IMAGE_OCR_VARIANTS": "gray,contrast",
                "DOCFLOW_TESSERACT_IMAGE_OCR_VARIANT_WORKERS": "2",
            },
            clear=False,
        ):
            results = ocr_engines._run_pil_ocr_variants(self.image_path, "tesseract", fake_runner)

        self.assertEqual(len(results), 2)
        self.assertLess(abs(starts[1] - starts[0]), 1.0)
        self.assertEqual([item["meta"]["variant_index"] for item in results], [0, 1])


if __name__ == "__main__":
    unittest.main()
