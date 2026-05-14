import unittest
from tempfile import TemporaryDirectory
from pathlib import Path
from unittest.mock import patch

from invoice_schema import (
    FIELD_LABELS,
    INVOICE_FIELD_ORDER,
    get_invoice_schema,
    get_invoice_schemas,
    list_invoice_schema_payloads,
    save_custom_invoice_schemas,
    infer_category_from_type_name,
    validate_invoice_schema_payload,
)


class InvoiceSchemaTests(unittest.TestCase):
    def test_all_schema_fields_have_labels_and_display_order(self):
        for schema_key in ("vat", "electronic", "normal", "taxi", "train", "quota", "receipt", "toll", "vehicle"):
            schema = get_invoice_schema(schema_key)
            with self.subTest(schema=schema_key):
                self.assertEqual(schema.key, schema_key)
                self.assertIn("invoice_type", schema.fields)
                for field_name in schema.fields:
                    self.assertIn(field_name, FIELD_LABELS)
                    self.assertIn(field_name, INVOICE_FIELD_ORDER)

    def test_aliases_infer_category(self):
        cases = {
            "增值税专用发票": "vat",
            "全国统一电子发票": "electronic",
            "出租车专用发票": "taxi",
            "中国铁路电子客票": "train",
            "定额发票": "quota",
            "收款收据": "receipt",
        }
        for type_name, expected in cases.items():
            with self.subTest(type_name=type_name):
                self.assertEqual(infer_category_from_type_name(type_name), expected)

    def test_custom_schema_can_be_saved_and_loaded(self):
        with TemporaryDirectory() as tmp_dir:
            custom_path = Path(tmp_dir) / "invoice_templates.json"
            with patch("invoice_schema.INVOICE_TEMPLATES_PATH", custom_path):
                save_custom_invoice_schemas(
                    [
                        {
                            "key": "hotel_receipt",
                            "name": "酒店收据",
                            "aliases": ["酒店收据", "住宿费"],
                            "fields": ["invoice_type", "invoice_number", "invoice_date", "total", "hotel_name"],
                            "labels": {"hotel_name": "酒店名称"},
                        }
                    ]
                )

                schemas = get_invoice_schemas()
                self.assertIn("hotel_receipt", schemas)
                self.assertEqual(schemas["hotel_receipt"].name, "酒店收据")
                self.assertEqual(infer_category_from_type_name("酒店收据"), "hotel_receipt")
                payloads = {item["key"]: item for item in list_invoice_schema_payloads()}
                self.assertEqual(payloads["hotel_receipt"]["labels"]["hotel_name"], "酒店名称")
                self.assertEqual(payloads["vat"]["labels"]["invoice_code"], "发票代码")

    def test_validate_invoice_schema_payload_rejects_invalid_input(self):
        normalized, error_message = validate_invoice_schema_payload(
            {
                "key": "bad template",
                "name": "无效模板",
                "fields": ["invoice_type", "hotel-name"],
            }
        )
        self.assertIsNone(normalized)
        self.assertEqual(error_message, "字段标识不合法：hotel-name")


if __name__ == "__main__":
    unittest.main()
