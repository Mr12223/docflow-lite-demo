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
VAT_0014_SPLIT_LAYOUT_OCR_TEXT = """江苏增镇税专用发票
3200164130
No 13638082
3200164130
13638082
开票日期：2017年04月20日
名
称：中国科学院自动化研究所
买方
纳税人识别号：
110108400010945
密
94430129+5+1+754*>/+1-7*1>0373
区
2016]311
电机
货物或应税劳务、服务名称
MCBL3006SRS
MCBL3002SRS
规格型号
单位
个
数量
7
2
2082.0512821
3205.1282051
单价
金额
14574.36
6410.26
税率
17%
税额
2477.64
1089.74
第三联
发票联
号南京造币有限公司
购买方记账凭证
合
计
¥20984.62
¥3567.38
价税合计（大写）
贰万肆仟伍佰伍拾贰圆整
（小写）¥24552.00
销
名
称：福尔哈贝传动技术（太仓）有限公司
备
售方
纳税人识别号：
91320585554626259G
地址、电话：太仓市经济开发区北京西路6号0512-53372626
开户行及账号：中国建设银行股份有限公司太仓支行32201997336051511825
注
尔
91320585554626259G
收款人：王琴
复核：赵予
开票人：孙涛
销售方专（章）"""

VAT_0014_EXPECTED_FIELDS = {
    "invoice_code": "3200164130",
    "invoice_number": "13638082",
    "invoice_date": "2017-04-20",
    "buyer_name": "中国科学院自动化研究所",
    "seller_name": "福尔哈贝传动技术(太仓)有限公司",
    "buyer_tax_id": "110108400010945",
    "seller_tax_id": "91320585554626259G",
    "amount": "20984.62",
    "tax": "3567.38",
    "total": "24552.00",
    "invoice_type": "增值税专用发票",
}


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

    def test_extract_invoice_fields_handles_split_vat_invoice_layout(self):
        result = extract_invoice_fields(VAT_0014_SPLIT_LAYOUT_OCR_TEXT)

        self.assertTrue(result.get("is_invoice"))
        actual_fields = result.get("fields") or {}
        for field_name, expected_value in VAT_0014_EXPECTED_FIELDS.items():
            actual_value = (actual_fields.get(field_name) or {}).get("value", "")
            self.assertEqual(
                _normalize_field(field_name, actual_value),
                _normalize_field(field_name, expected_value),
                f"vat_0014 field {field_name} mismatch",
            )


if __name__ == "__main__":
    unittest.main()
