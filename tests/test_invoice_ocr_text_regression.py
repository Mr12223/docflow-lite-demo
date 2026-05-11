import json
import re
import unittest
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

from docflow.webapp.services.invoice_merge import extract_invoice_fields


FIXTURE_PATH = Path(__file__).parent / "fixtures" / "invoice_ocr_text_samples.json"
MONEY_FIELDS = {"amount", "tax", "total"}
IDENTITY_FIELDS = {"invoice_code", "invoice_number", "buyer_tax_id", "seller_tax_id"}


def _collapse_spaces(value):
    return " ".join(str(value or "").strip().split())


def _normalize_text(value):
    return _collapse_spaces(value)


def _normalize_identity(value):
    return re.sub(r"[^A-Z0-9]", "", _normalize_text(value).upper())


def _normalize_amount(value):
    if value is None:
        return ""

    if isinstance(value, Decimal):
        amount = value
    elif isinstance(value, (int, float)):
        amount = Decimal(str(value))
    else:
        text = _normalize_text(value)
        if not text:
            return ""
        text = text.replace(",", "").replace("¥", "").replace("￥", "").replace("元", "").replace("圆", "")
        match = re.search(r"-?\d+(?:\.\d+)?", text)
        if not match:
            return ""
        try:
            amount = Decimal(match.group(0))
        except InvalidOperation:
            return ""

    return format(amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP), "f")


def _normalize_date(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()

    text = _normalize_text(value)
    if not text:
        return ""

    direct_match = re.search(r"((?:19|20)\d{2})[-/.年\s]*(\d{1,2})[-/.月\s]*(\d{1,2})", text)
    if direct_match:
        year, month, day = direct_match.groups()
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"

    digits_match = re.search(r"\b((?:19|20)\d{2})(\d{2})(\d{2})\b", text)
    if digits_match:
        year, month, day = digits_match.groups()
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"

    return text


def _normalize_field(field_name, value):
    if field_name in MONEY_FIELDS:
        return _normalize_amount(value)
    if field_name == "invoice_date":
        return _normalize_date(value)
    if field_name in IDENTITY_FIELDS:
        return _normalize_identity(value)
    return _normalize_text(value)


class InvoiceOcrTextRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.samples = json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))

    def test_fixture_contains_real_ocr_samples(self):
        self.assertGreaterEqual(len(self.samples), 6)
        for sample in self.samples:
            with self.subTest(sample_id=sample.get("sample_id")):
                self.assertTrue(sample.get("ocr_text"))
                self.assertTrue(sample.get("expected_fields"))
                self.assertEqual(sample.get("ocr_engine"), "RapidOCR")

    def test_extract_invoice_fields_matches_real_ocr_snapshots(self):
        for sample in self.samples:
            with self.subTest(sample_id=sample["sample_id"]):
                result = extract_invoice_fields(sample["ocr_text"])

                self.assertTrue(
                    result.get("is_invoice"),
                    f"{sample['sample_id']} should still be recognized as invoice",
                )

                actual_fields = result.get("fields") or {}
                for field_name, expected_value in sample["expected_fields"].items():
                    actual_value = (actual_fields.get(field_name) or {}).get("value", "")
                    self.assertEqual(
                        _normalize_field(field_name, actual_value),
                        _normalize_field(field_name, expected_value),
                        (
                            f"{sample['sample_id']} field {field_name} mismatch: "
                            f"expected={expected_value!r} actual={actual_value!r}"
                        ),
                    )


if __name__ == "__main__":
    unittest.main()
