import copy
import os
import unittest
from unittest.mock import patch

os.environ.setdefault("DOCFLOW_RAPIDOCR_PREWARM", "0")

from docflow.webapp.services import invoice_merge


BUYER_TAX_ID = "91310106MA1G12345A"
SELLER_TAX_ID = "91320214MA2B67890C"


class _StubExtractor:
    def __init__(self, payload):
        self.payload = copy.deepcopy(payload)

    def extract(self, text):
        return copy.deepcopy(self.payload)


def _build_invoice_payload(fields, confidence="high", is_invoice=True):
    normalized_fields = {}
    for field_name, value in (fields or {}).items():
        if isinstance(value, dict):
            field_payload = dict(value)
        else:
            field_payload = {
                "value": value,
                "confidence": confidence,
                "label": invoice_merge.FIELD_LABELS.get(field_name, field_name),
            }
        normalized_fields[field_name] = field_payload

    return {
        "is_invoice": is_invoice,
        "confidence": confidence if is_invoice else "none",
        "field_count": len(normalized_fields),
        "fields": normalized_fields,
    }


def _build_candidate(
    *,
    engine,
    invoice_fields=None,
    ocr_tax_id_fields=None,
    score=100,
    confidence="high",
    char_count=80,
    engine_index=0,
):
    invoice_payload = _build_invoice_payload(invoice_fields or {}, confidence=confidence)
    return {
        "engine": engine,
        "engine_index": engine_index,
        "char_count": char_count,
        "invoice_eval": {
            "score": score,
            "is_invoice": invoice_payload["is_invoice"],
            "confidence": invoice_payload["confidence"],
            "field_count": invoice_payload["field_count"],
            "major_field_count": invoice_payload["field_count"],
            "invoice_fields": invoice_payload,
        },
        "ocr_tax_id_fields": copy.deepcopy(ocr_tax_id_fields or {}),
    }


class InvoiceMergeFieldTests(unittest.TestCase):
    def test_extract_tax_id_values_from_line_supports_repeated_labels(self):
        line = (
            f"纳税人识别号:{BUYER_TAX_ID} "
            f"纳税人识别号:{SELLER_TAX_ID}"
        )

        values = invoice_merge._extract_ocr_tax_id_values_from_line(line)

        self.assertEqual(values, [BUYER_TAX_ID, SELLER_TAX_ID])

    def test_extract_ocr_tax_id_fields_detects_same_line_pair(self):
        text = "\n".join(
            [
                "购买方信息",
                "名称 北京甲公司",
                f"纳税人识别号:{BUYER_TAX_ID} 纳税人识别号:{SELLER_TAX_ID}",
            ]
        )

        result = invoice_merge._extract_ocr_tax_id_fields(text)

        self.assertEqual(result["buyer_tax_id"]["value"], BUYER_TAX_ID)
        self.assertEqual(result["seller_tax_id"]["value"], SELLER_TAX_ID)
        self.assertEqual(result["buyer_tax_id"]["method"], "same_line_pair")
        self.assertEqual(result["seller_tax_id"]["method"], "same_line_pair")

    def test_extract_ocr_tax_id_fields_detects_ordered_pair_across_sections(self):
        text = "\n".join(
            [
                "购买方信息",
                "名称 北京甲公司",
                f"纳税人识别号:{BUYER_TAX_ID}",
                "销售方信息",
                "名称 上海乙公司",
                f"纳税人识别号:{SELLER_TAX_ID}",
            ]
        )

        result = invoice_merge._extract_ocr_tax_id_fields(text)

        self.assertEqual(result["buyer_tax_id"]["value"], BUYER_TAX_ID)
        self.assertEqual(result["seller_tax_id"]["value"], SELLER_TAX_ID)
        self.assertEqual(result["buyer_tax_id"]["method"], "ordered_pair")
        self.assertEqual(result["seller_tax_id"]["method"], "ordered_pair")

    def test_extract_single_ocr_tax_id_field_skips_check_code_context(self):
        text = "\n".join(
            [
                "校验码",
                "67398706512584952290",
                "名",
                "称：中国科学院自动化研究所",
            ]
        )

        result = invoice_merge._extract_single_ocr_tax_id_field(text, "buyer_tax_id", "crop")

        self.assertEqual(result, {})

    def test_repair_ocr_tax_id_value_prefers_expected_length(self):
        truncated = "91320214MA2B6789C"

        repaired = invoice_merge._repair_ocr_tax_id_value(truncated, SELLER_TAX_ID)

        self.assertEqual(repaired, SELLER_TAX_ID)

    def test_merge_ocr_invoice_fields_fills_missing_tax_id_from_other_engine(self):
        selected = _build_candidate(
            engine="RapidOCR",
            invoice_fields={"invoice_number": "12345678"},
            ocr_tax_id_fields={
                "buyer_tax_id": {
                    "value": BUYER_TAX_ID,
                    "confidence": "high",
                    "method": "selected_structured",
                }
            },
        )
        other = _build_candidate(
            engine="Tesseract",
            invoice_fields={"invoice_number": "12345678"},
            ocr_tax_id_fields={
                "seller_tax_id": {
                    "value": SELLER_TAX_ID,
                    "confidence": "medium",
                    "method": "other_engine",
                }
            },
            engine_index=1,
        )

        merged, meta = invoice_merge._merge_ocr_invoice_fields(selected, [selected, other])

        self.assertEqual(merged["fields"]["buyer_tax_id"]["value"], BUYER_TAX_ID)
        self.assertEqual(merged["fields"]["seller_tax_id"]["value"], SELLER_TAX_ID)
        self.assertEqual(meta["fields"]["seller_tax_id"]["action"], "filled_from_other_engine")
        self.assertEqual(meta["fields"]["seller_tax_id"]["engine"], "Tesseract")

    def test_merge_ocr_invoice_fields_repairs_short_tax_id_from_other_engine(self):
        selected = _build_candidate(
            engine="RapidOCR",
            invoice_fields={
                "invoice_number": "12345678",
                "buyer_tax_id": {
                    "value": "91320214MA2B6789C",
                    "confidence": "medium",
                    "label": invoice_merge.FIELD_LABELS["buyer_tax_id"],
                },
            },
            ocr_tax_id_fields={
                "buyer_tax_id": {
                    "value": "91320214MA2B6789C",
                    "confidence": "medium",
                    "method": "selected_structured",
                }
            },
        )
        other = _build_candidate(
            engine="Tesseract",
            ocr_tax_id_fields={
                "buyer_tax_id": {
                    "value": SELLER_TAX_ID,
                    "confidence": "high",
                    "method": "other_engine",
                }
            },
            engine_index=1,
        )

        merged, meta = invoice_merge._merge_ocr_invoice_fields(selected, [selected, other])

        self.assertEqual(merged["fields"]["buyer_tax_id"]["value"], SELLER_TAX_ID)
        self.assertEqual(meta["fields"]["buyer_tax_id"]["action"], "repaired_from_other_engine")
        self.assertEqual(meta["fields"]["buyer_tax_id"]["engine"], "Tesseract")

    def test_merge_ocr_invoice_fields_clears_duplicate_unstructured_tax_id(self):
        selected = _build_candidate(
            engine="RapidOCR",
            invoice_fields={
                "buyer_tax_id": BUYER_TAX_ID,
                "seller_tax_id": BUYER_TAX_ID,
            },
            ocr_tax_id_fields={
                "buyer_tax_id": {
                    "value": BUYER_TAX_ID,
                    "confidence": "high",
                    "method": "selected_structured",
                }
            },
        )

        merged, meta = invoice_merge._merge_ocr_invoice_fields(selected, [selected])

        self.assertIn("buyer_tax_id", merged["fields"])
        self.assertNotIn("seller_tax_id", merged["fields"])
        self.assertEqual(meta["fields"]["seller_tax_id"]["action"], "cleared_duplicate")

    def test_merge_ocr_invoice_fields_marks_missing_without_filling_value(self):
        selected = _build_candidate(
            engine="RapidOCR",
            invoice_fields={
                "invoice_code": "1100153350",
                "invoice_number": "03159334",
                "total": "500.00",
            },
        )

        merged, meta = invoice_merge._merge_ocr_invoice_fields(selected, [selected])

        self.assertNotIn("buyer_tax_id", merged["fields"])
        self.assertIn("buyer_tax_id", merged["missing_fields"])
        self.assertEqual(merged["missing_fields"]["buyer_tax_id"]["message"], "原图疑似为空/未识别")
        self.assertFalse(meta["applied"])

    def test_electronic_invoice_marks_tax_ids_not_applicable(self):
        result = invoice_merge.extract_invoice_fields(
            "\n".join(
                [
                    "全国统一电子发票",
                    "发票号码：E3402243667",
                    "开票日期：2024年10月27日",
                    "购买方：上海贸易股份有限公司",
                    "销售方：成都软件开发有限公司",
                    "价税合计：￥43246.00",
                ]
            )
        )

        self.assertTrue(result["is_invoice"])
        self.assertEqual(result["invoice_category"], "electronic")
        self.assertNotIn("buyer_tax_id", result["missing_fields"])
        self.assertIn("buyer_tax_id", result["not_applicable_fields"])
        self.assertIn("invoice_number", result["expected_fields"])

    def test_train_ticket_uses_train_schema(self):
        result = invoice_merge.extract_invoice_fields(
            "\n".join(
                [
                    "中国铁路电子客票",
                    "乘车人：张三",
                    "车次：G1234",
                    "北京南站-上海虹桥站",
                    "乘车日期：2026年05月13日",
                    "二等座",
                    "票价：￥553.00",
                ]
            )
        )

        self.assertTrue(result["is_invoice"])
        self.assertEqual(result["invoice_category"], "train")
        self.assertEqual(result["fields"]["train_no"]["value"], "G1234")
        self.assertEqual(result["fields"]["fare"]["value"], "553.00")
        self.assertIn("departure_station", result["expected_fields"])
        self.assertIn("buyer_tax_id", result["not_applicable_fields"])

    def test_evaluate_ocr_invoice_text_scores_structured_payload(self):
        payload = _build_invoice_payload(
            {
                "invoice_code": "123456789012",
                "invoice_number": "87654321",
                "invoice_date": "2026-03-14",
                "amount": "100.00",
                "total": "106.00",
                "buyer_name": "北京甲公司",
                "seller_name": "上海乙公司",
                "invoice_type": "增值税普通发票",
            }
        )

        with patch.object(
            invoice_merge,
            "_get_invoice_extractor",
            return_value=_StubExtractor(payload),
        ):
            result = invoice_merge._evaluate_ocr_invoice_text(
                "发票代码123456789012 发票号码87654321 开票日期2026-03-14 金额100.00 合计106.00"
            )

        self.assertTrue(result["is_invoice"])
        self.assertEqual(result["confidence"], "high")
        self.assertEqual(result["field_count"], 8)
        self.assertTrue(result["header_pair"])
        self.assertTrue(result["money_pair"])
        self.assertTrue(result["party_pair"])
        self.assertGreater(result["score"], 150)


if __name__ == "__main__":
    unittest.main()
