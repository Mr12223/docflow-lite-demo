"""
发票结构化信息提取模块。
"""

from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Optional

from docflow.invoice.schema import (
    get_field_labels,
    get_invoice_schema,
    infer_category_from_type_name,
)


_SEP = r"[：:\s]*"
_VALID_DATE_TEXT = (
    r"((?:19|20)\d{2}"
    r"[\s年/\-.]"
    r"(?:1[0-2]|0?[1-9])"
    r"[\s月/\-.]"
    r"(?:3[01]|[12]\d|0?[1-9])"
    r"\s*日?)"
)
_VALID_DATE_DIGITS = r"((?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01]))"

FIELD_PATTERNS: dict[str, list[tuple[str, str]]] = {
    "invoice_code": [
        (r"(?:发票代码|代码)" + _SEP + r"(\d{10,12})", "high"),
    ],
    "invoice_number": [
        (r"(?:发票号码|号码|No\.?)" + _SEP + r"([A-Za-z]?\d{8,20})", "high"),
    ],
    "invoice_date": [
        (r"(?:开票日期|日期|Date)" + _SEP + _VALID_DATE_TEXT, "high"),
        (r"(?:开票日期|日期|Date)" + _SEP + _VALID_DATE_DIGITS, "medium"),
        (_VALID_DATE_TEXT, "low"),
        (_VALID_DATE_DIGITS, "low"),
    ],
    "amount": [
        (r"金额" + _SEP + r"[¥￥]?\s*(\d[\d,]*(?:\.\d{1,2})?)", "medium"),
    ],
    "tax": [
        (r"税额" + _SEP + r"[¥￥]?\s*(\d[\d,]*(?:\.\d{1,2})?)", "medium"),
    ],
    "total": [
        (r"(?:价税合计|小写)" + _SEP + r"[¥￥]?\s*(\d[\d,]*(?:\.\d{1,2})?)", "high"),
    ],
    "buyer_name": [
        (
            r"(?:购买方|购方)(?:名称)?"
            + _SEP
            + r"(.{2,40}?)(?:纳税人识别号|地址电话|开户地址|开户行及账号|密码区|货物或应税劳务|$)",
            "high",
        ),
    ],
    "seller_name": [
        (
            r"(?:销售方|销方)(?:名称)?"
            + _SEP
            + r"(.{2,40}?)(?:纳税人识别号|地址电话|开户地址|开户行及账号|发票专用章|收款人|复核|开票人|$)",
            "high",
        ),
    ],
    "buyer_tax_id": [
        (
            r"(?:购买方|购方).{0,80}?(?:纳税人)?识别号"
            + _SEP
            + r"([A-Za-z0-9]{15,20})",
            "high",
        ),
    ],
    "seller_tax_id": [
        (
            r"(?:销售方|销方).{0,80}?(?:纳税人)?识别号"
            + _SEP
            + r"([A-Za-z0-9]{15,20})",
            "high",
        ),
    ],
    "machine_number": [
        (r"机器编号" + _SEP + r"(\d{6,14})", "high"),
    ],
    "check_code": [
        (r"校\s*验\s*码" + _SEP + r"([\d\s]{12,25})", "high"),
    ],
    "fuel_oil_surcharge": [
        (r"(?:燃油附加费|Fuel\s*[Oo]il\s*[Ss]urcharge)" + _SEP + r"[¥￥]?\s*(\d[\d,]*(?:\.\d{1,2})?)", "high"),
    ],
}

INVOICE_TYPE_KEYWORDS: list[tuple[str, str]] = [
    ("收款收据", "收据"),
    ("收据", "收据"),
    ("RECEIPT", "收据"),
    ("Receipt", "收据"),
    ("增值税电子专用发票", "增值税电子专用发票"),
    ("增值税电子普通发票", "增值税电子普通发票"),
    ("增值税专用发票", "增值税专用发票"),
    ("增镇税专用发票", "增值税专用发票"),
    ("增值税普通发票", "增值税普通发票"),
    ("增值税发票", "增值税发票"),
    ("通发票", "增值税发票"),
    ("电子发票", "电子发票"),
    ("出租车专用发票", "出租车发票"),
    ("出租车发票", "出租车发票"),
    ("TAXI", "出租车发票"),
    ("Taxi", "出租车发票"),
    ("机动车销售统一发票", "机动车销售统一发票"),
    ("通行费电子发票", "通行费电子发票"),
]

_NAME_STOPWORDS = {
    "纳税人",
    "记账凭证",
    "收款人",
    "复核",
    "开票人",
    "管理员",
    "销售方",
    "购买方",
    "发票联",
    "密码区",
    "价税合计",
    "校验码",
    "机器编号",
    "地址电话",
    "开户行及账号",
    "货物或应税劳务",
    "货物或应税劳务、服务名称",
    "商品名称",
    "发票专用章",
    "注",
    "备",
}
_NAME_BAD_PARTS = ("（章", "(章", "发票专用章", "收款人", "复核", "开票人", "管理员")
_COMPANY_HINTS = ("公司", "研究所", "大学", "学院", "银行", "中心", "商店", "饭店", "超市", "分公司")


class InvoiceExtractor:
    def extract(self, text: str) -> dict:
        if not text or not text.strip():
            return self._empty_result()

        cleaned = self._preprocess(text)
        compact = self._compact_text(cleaned)
        lines = self._split_lines(cleaned)
        invoice_category = self._detect_invoice_category(cleaned, compact, lines)

        fields: dict[str, dict] = {}
        for field_name, patterns in FIELD_PATTERNS.items():
            result = self._try_field(field_name, cleaned, compact, patterns)
            if not result:
                continue
            value, confidence = result
            value = self._clean_value(field_name, value)
            if value:
                fields[field_name] = self._pack_field(field_name, value, confidence)

        self._apply_fallbacks(cleaned, compact, lines, fields, invoice_category)

        invoice_type = self._detect_invoice_type(cleaned, compact)
        if invoice_type:
            fields["invoice_type"] = self._pack_field("invoice_type", invoice_type, "high")
        elif invoice_category != "unknown":
            schema_name = get_invoice_schema(invoice_category).name
            fields["invoice_type"] = self._pack_field("invoice_type", schema_name, "medium")

        score = self._score_invoice(fields, compact)
        has_money = "amount" in fields or "total" in fields
        has_header = "invoice_code" in fields or "invoice_number" in fields
        has_parties = "buyer_name" in fields or "seller_name" in fields

        has_typed_schema = invoice_category != "unknown" and len(fields) >= 2
        is_invoice = (
            {"invoice_code", "invoice_number"}.issubset(fields)
            or has_typed_schema
            or (
                "invoice_date" in fields
                and has_money
                and ("invoice_type" in fields or "check_code" in fields or has_header or has_parties)
            )
            or (score >= 5)
        )

        if not is_invoice:
            overall_confidence = "none"
        elif score >= 7 or (has_typed_schema and len(fields) >= 4):
            overall_confidence = "high"
        elif score >= 5 or has_typed_schema:
            overall_confidence = "medium"
        else:
            overall_confidence = "low"

        return {
            "is_invoice": is_invoice,
            "confidence": overall_confidence,
            "field_count": len(fields),
            "fields": fields,
            "invoice_category": invoice_category if is_invoice else "unknown",
            "schema_name": get_invoice_schema(invoice_category).name,
        }

    def _apply_fallbacks(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
        fields: dict[str, dict],
        invoice_category: str = "unknown",
    ) -> None:
        expected_fields = set(self._expected_fields_for_category(invoice_category))
        ordered_extractors = [
            ("invoice_code", self._find_invoice_code),
            ("invoice_number", self._find_invoice_number),
            ("invoice_date", self._find_invoice_date),
            ("machine_number", self._find_machine_number),
            ("check_code", self._find_check_code),
            ("buyer_name", self._find_buyer_name),
            ("seller_name", self._find_seller_name),
            ("buyer_tax_id", self._find_buyer_tax_id),
            ("seller_tax_id", self._find_seller_tax_id),
            ("passenger_name", self._find_train_passenger_name),
            ("train_no", self._find_train_no),
            ("departure_station", self._find_departure_station),
            ("arrival_station", self._find_arrival_station),
            ("travel_date", self._find_travel_date),
            ("seat_class", self._find_seat_class),
            ("taxi_car_no", self._find_taxi_car_no),
            ("taxi_start_time", self._find_taxi_start_time),
            ("taxi_end_time", self._find_taxi_end_time),
            ("fare", self._find_fare),
            ("quota_amount", self._find_quota_amount),
            ("fuel_oil_surcharge", self._find_fuel_oil_surcharge),
        ]

        for field_name, extractor in ordered_extractors:
            if invoice_category != "unknown" and field_name not in expected_fields:
                continue
            if field_name in fields:
                continue
            result = extractor(cleaned, compact, lines)
            if not result:
                continue
            value, confidence = result
            value = self._clean_value(field_name, value)
            if value:
                fields[field_name] = self._pack_field(field_name, value, confidence)

        money_fields = self._find_money_fields(cleaned, compact, lines, fields)
        if invoice_category == "train":
            money_fields = self._rename_money_fields(money_fields, amount_field="fare")
        elif invoice_category == "taxi":
            money_fields = self._rename_money_fields(money_fields, amount_field="fare")
        elif invoice_category == "quota":
            money_fields = self._rename_money_fields(money_fields, amount_field="quota_amount")

        for field_name, (value, confidence) in money_fields.items():
            if invoice_category != "unknown" and field_name not in expected_fields:
                continue
            if field_name in fields:
                continue
            value = self._clean_value(field_name, value)
            if value:
                fields[field_name] = self._pack_field(field_name, value, confidence)

    def _find_invoice_code(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        if self._looks_like_receipt(cleaned, compact):
            return None

        top_lines = lines[:12]
        counts: dict[str, int] = {}
        first_pos: dict[str, int] = {}
        for idx, line in enumerate(top_lines):
            digits = self._digits_only(line)
            if 10 <= len(digits) <= 12:
                counts[digits] = counts.get(digits, 0) + 1
                first_pos.setdefault(digits, idx)
        if counts:
            best = sorted(counts, key=lambda value: (-counts[value], first_pos[value]))[0]
            confidence = "high" if counts[best] >= 2 else "medium"
            return best, confidence

        match = re.search(r"(?<!\d)(\d{10,12})(?!\d)", compact)
        if match:
            return match.group(1), "low"
        return None

    def _find_invoice_number(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        receipt_number = self._find_receipt_number(lines)
        if receipt_number:
            return receipt_number

        labeled_number = self._find_labeled_invoice_number(lines)
        if labeled_number:
            return labeled_number

        match = re.search(r"\bNo\.?\s*([A-Za-z]?\d{8,20})\b", cleaned, re.IGNORECASE)
        if match:
            return match.group(1), "high"

        top_lines = lines[:12]
        counts: dict[str, int] = {}
        first_pos: dict[str, int] = {}
        for idx, line in enumerate(top_lines):
            for candidate in self._extract_invoice_number_candidates(line):
                normalized = re.sub(r"[^A-Za-z0-9]", "", candidate)
                if not self._is_plausible_invoice_number(normalized, line=line):
                    continue
                counts[normalized] = counts.get(normalized, 0) + 1
                first_pos.setdefault(normalized, idx)
        if counts:
            best = sorted(counts, key=lambda value: (-counts[value], first_pos[value]))[0]
            confidence = "high" if counts[best] >= 2 else "medium"
            return best, confidence

        for match in re.finditer(r"(?<![A-Za-z0-9])([A-Za-z]?\d{8,20})(?![A-Za-z0-9])", compact):
            candidate = match.group(1)
            if self._is_plausible_invoice_number(candidate):
                return candidate, "low"
        return None

    def _find_invoice_date(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        labeled_date = self._find_labeled_invoice_date(lines)
        if labeled_date:
            return labeled_date

        for line in lines[:20]:
            date_value = self._extract_valid_date(line)
            if date_value:
                return date_value, "medium"

        date_value = self._extract_valid_date(compact)
        if date_value:
            return date_value, "low"
        return None

    def _find_machine_number(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        for line in lines[:20]:
            match = re.search(r"机器编号" + _SEP + r"(\d{6,14})", line)
            if match:
                return match.group(1), "high"
        return None

    def _find_check_code(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        for idx, line in enumerate(lines):
            if "校验码" not in line:
                continue
            merged = "".join(lines[idx : idx + 2])
            match = re.search(r"校\s*验\s*码" + _SEP + r"([\d\s]{12,25})", merged)
            if match:
                return re.sub(r"\s+", "", match.group(1)), "high"
        return None

    def _find_buyer_name(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        paired_name = self._find_dual_section_value(lines, ("名称",), index=0)
        if paired_name and self._looks_like_name(paired_name):
            return paired_name, "high"

        ordered_name = self._find_ordered_party_name(lines, index=0)
        if ordered_name:
            return ordered_name, "medium"

        receipt_name = self._find_label_value(lines, ("付款单位", "付款方"), self._looks_like_name)
        if receipt_name:
            return receipt_name, "high"

        return self._find_party_name(compact, lines, ("购买方", "购方", "买方"), prefer="first")

    def _find_seller_name(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        paired_name = self._find_dual_section_value(lines, ("名称",), index=1)
        if paired_name and self._looks_like_name(paired_name):
            return paired_name, "high"

        ordered_name = self._find_ordered_party_name(lines, index=1)
        if ordered_name:
            return ordered_name, "medium"

        receipt_name = self._find_label_value(lines, ("收款单位", "收款方"), self._looks_like_name)
        if receipt_name:
            return receipt_name, "high"

        return self._find_party_name(compact, lines, ("销售方", "销方", "售方"), prefer="last")

    def _find_buyer_tax_id(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        paired_tax_id = self._find_dual_section_value(lines, ("纳税人识别号", "税号"), index=0)
        if paired_tax_id:
            return paired_tax_id, "high"

        ordered_tax_id = self._find_ordered_tax_id(lines, index=0)
        if ordered_tax_id:
            return ordered_tax_id, "high"

        return self._find_party_tax_id(compact, lines, ("购买方", "购方", "买方"))

    def _find_seller_tax_id(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        paired_tax_id = self._find_dual_section_value(lines, ("纳税人识别号", "税号"), index=1)
        if paired_tax_id:
            return paired_tax_id, "high"

        ordered_tax_id = self._find_ordered_tax_id(lines, index=1)
        if ordered_tax_id:
            return ordered_tax_id, "high"

        return self._find_party_tax_id(compact, lines, ("销售方", "销方", "售方"))

    def _find_train_passenger_name(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        value = self._find_label_value(lines, ("乘车人", "旅客", "姓名"), self._looks_like_person_name)
        return (value, "high") if value else None

    def _find_train_no(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        for line in lines[:20]:
            match = re.search(r"(?:车次|列车)[：:\s]*([GCDZTK]?\d{1,5})", line, re.IGNORECASE)
            if match:
                return match.group(1).upper(), "high"
            match = re.search(r"(?<![A-Za-z0-9])([GCDZTK]\d{1,5})(?![A-Za-z0-9])", line, re.IGNORECASE)
            if match:
                return match.group(1).upper(), "medium"
        return None

    def _find_departure_station(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        station_pair = self._find_train_station_pair(lines)
        if station_pair:
            return station_pair[0], "medium"
        return None

    def _find_arrival_station(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        station_pair = self._find_train_station_pair(lines)
        if station_pair:
            return station_pair[1], "medium"
        return None

    def _find_travel_date(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        for line in lines[:20]:
            if not any(term in line for term in ("乘车日期", "发车日期", "日期", "开车时间")):
                continue
            date_value = self._extract_valid_date(line)
            if date_value:
                return date_value, "high"
        date_value = self._extract_valid_date(cleaned)
        return (date_value, "medium") if date_value else None

    def _find_seat_class(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        for line in lines[:30]:
            match = re.search(r"(商务座|特等座|一等座|二等座|硬座|软座|硬卧|软卧|无座)", line)
            if match:
                return match.group(1), "medium"
        return None

    def _find_taxi_car_no(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        _PLATE_CHARS = r"[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使A-Z]"
        _PLATE_BODY = r"[A-Z0-9\-\.挂学警港澳]{4,9}"
        _LABEL = r"(?:车牌号|车号|车辆号|Car\s*No\.?|Taxi\s*No\.?)"
        search_lines = lines[:30]
        for i, line in enumerate(search_lines):
            # 正常情况：标签和车牌在同一行
            match = re.search(
                rf"{_LABEL}[：:\s]*({_PLATE_CHARS}{_PLATE_BODY})",
                line, re.IGNORECASE,
            )
            if match:
                plate = re.sub(r"[-.]", "", match.group(1)).upper()
                return plate, "high"
            # OCR 换行情况：省份字符在标签行末，车牌其余部分在下一行
            # 例如 "车号京\nB-N4585" 或 "车号京\nB.06908"
            split_match = re.search(
                rf"{_LABEL}[：:\s]*({_PLATE_CHARS})\s*$",
                line, re.IGNORECASE,
            )
            if split_match and i + 1 < len(search_lines):
                next_line = search_lines[i + 1].strip()
                body_match = re.match(r"([A-Z0-9\-\.挂学警港澳]{4,9})", next_line, re.IGNORECASE)
                if body_match:
                    plate = re.sub(r"[-.]", "", split_match.group(1) + body_match.group(1)).upper()
                    return plate, "high"
        return None

    def _find_taxi_time_range(self, lines: list[str]) -> Optional[tuple[str, str]]:
        """从 OCR 行中提取出租车时间范围字符串，返回 (start, end) 或 None。
        处理 "HH:MM-HH:MM" 格式，支持标签前后两种布局。
        """
        _TIME = r"\d{1,2}:\d{2}"
        _RANGE = rf"({_TIME})\s*[-–]\s*({_TIME})"
        _LABELS = ("时间", "Time", "上车时间", "下车时间")
        label_expr = "|".join(map(re.escape, _LABELS))
        for idx, line in enumerate(lines[:40]):
            # 同行有范围时间
            m = re.search(_RANGE, line)
            if m:
                return m.group(1), m.group(2)
            # 标签行：往下一行找范围时间
            if re.search(label_expr, line, re.IGNORECASE) and idx + 1 < len(lines):
                m = re.search(_RANGE, lines[idx + 1])
                if m:
                    return m.group(1), m.group(2)
            # 标签行：往上一行找范围时间（值在标签前的布局）
            if re.search(label_expr, line, re.IGNORECASE) and idx > 0:
                m = re.search(_RANGE, lines[idx - 1])
                if m:
                    return m.group(1), m.group(2)
        return None

    def _find_taxi_start_time(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        result = self._find_taxi_time_range(lines)
        if result:
            return result[0], "high"
        return self._find_labeled_time(lines, ("上车时间", "上车", "Start", "Start Time"))

    def _find_taxi_end_time(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        result = self._find_taxi_time_range(lines)
        if result:
            return result[1], "high"
        return self._find_labeled_time(lines, ("下车时间", "下车", "End", "End Time"))

    def _find_fare(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        return self._find_labeled_money_multiline(lines, ("票价", "金额", "Fare", "合计"))

    def _find_quota_amount(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        return self._find_labeled_money(lines, ("定额", "金额", "票面金额", "合计"))

    def _find_fuel_oil_surcharge(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        return self._find_labeled_money_multiline(lines, ("燃油附加费", "Fuel Oil Surcharge", "Fuel oll surcharge"))

    def _find_party_name(
        self,
        compact: str,
        lines: list[str],
        role_terms: tuple[str, ...],
        prefer: str,
    ) -> Optional[tuple[str, str]]:
        marker_indexes = [
            idx
            for idx, line in enumerate(lines)
            if any(term in line for term in role_terms)
            and not any(stop in line for stop in ("记账凭证", "发票联"))
        ]
        for idx in marker_indexes:
            direct = self._extract_role_name(lines[idx], role_terms)
            if direct and self._looks_like_name(direct):
                return direct, "high"

            window = lines[max(0, idx - 3) : min(len(lines), idx + 6)]
            candidates: list[tuple[int, str]] = []
            for offset, line in enumerate(window):
                direct = self._extract_role_name(line, role_terms)
                if direct and self._looks_like_name(direct):
                    candidates.append((abs((max(0, idx - 3) + offset) - idx), direct))
                    continue
                candidate = self._extract_name_candidate(line)
                if candidate and self._looks_like_name(candidate):
                    candidates.append((abs((max(0, idx - 3) + offset) - idx), candidate))
            if candidates:
                candidates.sort(key=lambda item: item[0])
                return candidates[0][1], "medium"

        named_candidates: list[str] = []
        for line in lines:
            if "称" not in line:
                continue
            candidate = self._extract_name_candidate(line)
            if candidate and self._looks_like_name(candidate):
                named_candidates.append(candidate)
        if named_candidates:
            selected = named_candidates[0] if prefer == "first" else named_candidates[-1]
            return selected, "low"
        return None

    def _find_ordered_party_name(self, lines: list[str], index: int) -> str:
        candidates: list[str] = []
        seen: set[str] = set()

        def add_candidate(raw_value: str) -> bool:
            candidate = self._clean_name(raw_value)
            if not candidate or candidate in seen or not self._looks_like_name(candidate):
                return False
            seen.add(candidate)
            candidates.append(candidate)
            return True

        for idx, line in enumerate(lines):
            compact_line = self._compact_text(line)
            if self._is_table_header_line(line) or any(term in compact_line for term in ("商品名称", "项目名称", "服务名称")):
                continue

            if compact_line in {"名", "名称"}:
                for next_line in lines[idx + 1 : min(len(lines), idx + 4)]:
                    if add_candidate(self._extract_name_candidate(next_line)):
                        break
                continue

            if not compact_line.startswith("称"):
                continue
            add_candidate(self._extract_name_candidate(line))

        return candidates[index] if len(candidates) > index else ""

    def _find_party_tax_id(
        self,
        compact: str,
        lines: list[str],
        role_terms: tuple[str, ...],
    ) -> Optional[tuple[str, str]]:
        role_expr = "|".join(map(re.escape, role_terms))
        pattern = (
            rf"(?:{role_expr}).{{0,80}}?(?:纳税人)?识别号"
            + _SEP
            + r"([A-Za-z0-9]{15,20})"
        )
        match = re.search(pattern, compact)
        if match:
            return match.group(1), "high"

        marker_indexes = [idx for idx, line in enumerate(lines) if any(term in line for term in role_terms)]
        for idx in marker_indexes:
            window = lines[idx : min(len(lines), idx + 8)]
            joined = "".join(window)
            match = re.search(r"([A-Za-z0-9]{15,20})", joined)
            if match:
                return match.group(1), "medium"
        return None

    def _find_ordered_tax_id(self, lines: list[str], index: int) -> str:
        candidates: list[str] = []
        seen: set[str] = set()
        label_pattern = re.compile(r"(?:纳税人识别号|税号)")

        for idx, line in enumerate(lines):
            if not label_pattern.search(line):
                continue

            tail = label_pattern.split(line, maxsplit=1)[-1]
            snippet_lines = [tail, *lines[idx + 1 : min(len(lines), idx + 4)]]
            for snippet in snippet_lines:
                for match in re.finditer(r"[A-Za-z0-9]{15,20}", snippet):
                    candidate = self._normalize_tax_id(match.group(0))
                    if not self._is_plausible_tax_id(candidate) or candidate in seen:
                        continue
                    seen.add(candidate)
                    candidates.append(candidate)
                    break
                if len(candidates) > index:
                    return candidates[index]

        return candidates[index] if len(candidates) > index else ""

    def _normalize_tax_id(self, value: str) -> str:
        return re.sub(r"[^A-Za-z0-9]", "", str(value or "").upper())

    def _is_plausible_tax_id(self, value: str) -> bool:
        value = self._normalize_tax_id(value)
        if not 15 <= len(value) <= 20:
            return False
        return len(re.findall(r"\d", value)) >= 8

    def _find_money_fields(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
        fields: dict[str, dict],
    ) -> dict[str, tuple[str, str]]:
        result: dict[str, tuple[str, str]] = {}

        summary_fields = self._find_summary_money_fields(lines)
        result.update(summary_fields)

        total = result.get("total")
        if not total:
            total = self._find_total(lines, compact)
            if total:
                result["total"] = total

        table_amount_tax = self._find_amount_tax_from_table(
            lines,
            known_total=(result.get("total") or ("", ""))[0],
        )
        for field_name, value in table_amount_tax.items():
            result.setdefault(field_name, value)

        summary_amount_tax = self._find_amount_tax_from_summary(lines, known_total=(total or ("", ""))[0])
        for field_name, value in summary_amount_tax.items():
            result.setdefault(field_name, value)

        if "amount" not in result or "tax" not in result or "total" not in result:
            derived_fields = self._derive_money_fields(
                lines,
                cleaned,
                compact,
                known_total=(result.get("total") or ("", ""))[0],
                known_amount=(result.get("amount") or ("", ""))[0],
                known_tax=(result.get("tax") or ("", ""))[0],
            )
            for field_name, value in derived_fields.items():
                result.setdefault(field_name, value)

        if "amount" not in result:
            explicit_amount = self._find_explicit_amount(lines)
            if explicit_amount:
                result.setdefault("amount", explicit_amount)

        if "tax" not in result:
            explicit_tax = self._find_explicit_tax(
                lines,
                known_total=(total or ("", ""))[0],
                known_amount=(result.get("amount") or ("", ""))[0],
            )
            if explicit_tax:
                result.setdefault("tax", explicit_tax)

        if self._looks_like_taxi_invoice(cleaned, compact):
            if "amount" in result and "total" not in result:
                result["total"] = result["amount"]
            elif "total" in result and "amount" not in result:
                result["amount"] = result["total"]

        if self._looks_like_receipt(cleaned, compact):
            if "amount" in result and "total" not in result:
                result["total"] = result["amount"]
            elif "total" in result and "amount" not in result:
                result["amount"] = result["total"]
            result.pop("tax", None)

        self._reconcile_money_fields(result)

        return result

    def _find_total(self, lines: list[str], compact: str) -> Optional[tuple[str, str]]:
        summary_fields = self._find_summary_money_fields(lines)
        if "total" in summary_fields:
            return summary_fields["total"]

        for idx, line in enumerate(lines):
            if self._is_table_header_line(line):
                continue
            compact_line = self._compact_text(line)
            if "小写" in line:
                values = self._collect_money_values([line])
                if values:
                    return values[-1], "high"
            if "价税合计" in line or compact_line == "合计" or compact_line.startswith("合计"):
                values = self._collect_money_values([line])
                if values:
                    return values[-1], "high"

                if idx + 1 < len(lines):
                    values = self._collect_money_values(lines[idx : idx + 2])
                    if values:
                        return values[-1], "medium"
        return None

    def _find_amount_tax_from_summary(
        self,
        lines: list[str],
        known_total: str = "",
    ) -> dict[str, tuple[str, str]]:
        result: dict[str, tuple[str, str]] = {}
        summary_fields = self._find_summary_money_fields(lines, known_total=known_total)
        if "amount" in summary_fields:
            result["amount"] = summary_fields["amount"]
        if "tax" in summary_fields:
            result["tax"] = summary_fields["tax"]
        return result

    def _find_amount_tax_from_table(
        self,
        lines: list[str],
        known_total: str = "",
    ) -> dict[str, tuple[str, str]]:
        result: dict[str, tuple[str, str]] = {}
        amount_candidates, tax_candidates = self._collect_rate_amount_tax_candidates(
            lines,
            known_total=known_total,
        )
        pairs = self._collect_rate_amount_tax_pairs(lines, known_total=known_total)

        if pairs:
            amount_sum = self._sum_money_values([amount for amount, _tax in pairs])
            tax_sum = self._sum_money_values([tax for _amount, tax in pairs])
            total_sum = self._sum_money_values(
                [amount for amount, _tax in pairs] + [tax for _amount, tax in pairs]
            )
            if not known_total or (total_sum and self._same_money(total_sum, known_total)):
                confidence = "high" if len(pairs) >= 2 else "medium"
                if amount_sum:
                    result["amount"] = (amount_sum, confidence)
                if tax_sum:
                    result["tax"] = (tax_sum, confidence)
                if total_sum:
                    result.setdefault("total", (total_sum, confidence))
                return result

        if known_total:
            repaired = self._repair_amount_tax_from_candidates(
                amount_candidates,
                tax_candidates,
                known_total,
            )
            if repaired:
                return repaired
        return result

    def _find_amount_tax_from_rate(
        self,
        lines: list[str],
        known_total: str = "",
    ) -> dict[str, tuple[str, str]]:
        return self._find_amount_tax_from_table(lines, known_total=known_total)

    def _find_explicit_amount(self, lines: list[str]) -> Optional[tuple[str, str]]:
        for idx, line in enumerate(lines):
            if "金额" not in line and line.lower() != "fare":
                continue
            if self._is_table_header_line(line):
                continue
            prev_values = self._collect_money_values([lines[idx - 1]]) if idx > 0 else []
            values = self._collect_money_values(lines[max(0, idx - 1) : idx + 4])
            if values:
                if prev_values and prev_values[-1] in values:
                    return prev_values[-1], "medium"
                return self._max_money_value(values), "medium"
        return None

    def _find_explicit_tax(
        self,
        lines: list[str],
        known_total: str = "",
        known_amount: str = "",
    ) -> Optional[tuple[str, str]]:
        for idx, line in enumerate(lines):
            if "税额" not in line:
                continue
            if self._is_table_header_line(line):
                continue
            values = self._collect_money_values(lines[idx : idx + 4])
            for value in values:
                if known_amount and self._same_money(value, known_amount):
                    continue
                if known_total and self._same_money(value, known_total):
                    continue
                return value, "medium"
        return None

    def _try_field(
        self,
        field_name: str,
        cleaned: str,
        compact: str,
        patterns: list[tuple[str, str]],
    ) -> Optional[tuple[str, str]]:
        fallback_only_fields = {
            "amount",
            "tax",
            "total",
            "buyer_name",
            "seller_name",
            "buyer_tax_id",
            "seller_tax_id",
        }
        if field_name in fallback_only_fields:
            return None

        cleaned_only_fields = {
            "total",
        }
        result = self._try_patterns(cleaned, patterns)
        if result or field_name in cleaned_only_fields:
            return result
        return self._try_patterns(compact, patterns)

    def _try_patterns(
        self,
        text: str,
        patterns: list[tuple[str, str]],
    ) -> Optional[tuple[str, str]]:
        for pattern, confidence in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip(), confidence
        return None

    def _detect_invoice_type(self, cleaned: str, compact: str) -> Optional[str]:
        for keyword, type_name in INVOICE_TYPE_KEYWORDS:
            if keyword in cleaned or keyword in compact:
                return type_name
        lower_text = cleaned.lower()
        if self._looks_like_train_ticket(cleaned, compact):
            return "火车票"
        if "定额发票" in cleaned or "定额发票" in compact:
            return "定额发票"
        if ("invoice" in lower_text or "nvoice" in lower_text) and ("taxi" in lower_text or "出租" in cleaned):
            return "出租车发票"
        if "价税合计" in compact or "校验码" in compact or "发票专用章" in compact:
            return "增值税发票"
        return None

    def _detect_invoice_category(self, cleaned: str, compact: str, lines: list[str]) -> str:
        type_name = self._detect_invoice_type(cleaned, compact) or ""
        schema_category = infer_category_from_type_name(type_name)
        if schema_category != "unknown":
            return schema_category
        if self._looks_like_train_ticket(cleaned, compact):
            return "train"
        if self._looks_like_taxi_invoice(cleaned, compact):
            return "taxi"
        if "定额发票" in compact:
            return "quota"
        if self._looks_like_receipt(cleaned, compact):
            return "receipt"
        if "全国统一电子发票" in cleaned:
            return "electronic"
        if any(term in compact for term in ("价税合计", "发票专用章", "校验码")):
            return "vat"
        if type_name:
            return "unknown"
        return "unknown"

    def _expected_fields_for_category(self, invoice_category: str) -> tuple[str, ...]:
        return tuple(get_invoice_schema(invoice_category).fields)

    def _rename_money_fields(self, money_fields: dict[str, tuple[str, str]], amount_field: str) -> dict[str, tuple[str, str]]:
        result = dict(money_fields)
        source = result.get("total") or result.get("amount")
        if source:
            result.setdefault(amount_field, source)
        return result

    def _score_invoice(self, fields: dict[str, dict], compact: str) -> int:
        score = 0
        if "invoice_code" in fields:
            score += 2
        if "invoice_number" in fields:
            score += 2
        if "invoice_date" in fields:
            score += 1
        if "amount" in fields:
            score += 1
        if "total" in fields:
            score += 1
        if "check_code" in fields:
            score += 1
        if "machine_number" in fields:
            score += 1
        if "invoice_type" in fields:
            score += 1
        if "buyer_name" in fields and "seller_name" in fields:
            score += 1
        elif "buyer_name" in fields or "seller_name" in fields:
            score += 0
        if "发票" in compact or "invoice" in compact.lower() or "nvoice" in compact.lower():
            score += 1
        return int(score)

    def _pack_field(self, field_name: str, value: str, confidence: str) -> dict:
        return {
            "value": value,
            "confidence": confidence,
            "label": get_field_labels().get(field_name, field_name),
        }

    def _preprocess(self, text: str) -> str:
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = text.replace("\uff1a", "：")
        text = text.replace("（", "(").replace("）", ")")
        text = text.replace("，", ",")
        text = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", text)
        text = re.sub(r"[^\S\n]+", " ", text)
        return text.strip()

    def _compact_text(self, text: str) -> str:
        return re.sub(r"\s+", "", text)

    def _split_lines(self, text: str) -> list[str]:
        return [line.strip() for line in text.splitlines() if line.strip()]

    def _clean_value(self, field_name: str, value: str) -> str:
        value = str(value or "").strip()
        if not value:
            return ""

        if field_name in {"amount", "tax", "total", "fare", "quota_amount"}:
            value = value.replace(",", "").replace("，", "")
            value = value.replace("¥", "").replace("￥", "").strip()
            if not self._looks_like_money(value):
                return ""
            return value

        if field_name in {"invoice_date", "travel_date"}:
            return self._normalize_date(value)

        if field_name in {"buyer_tax_id", "seller_tax_id"}:
            value = re.sub(r"\s+", "", value)
            return value if 15 <= len(value) <= 20 else ""

        if field_name == "check_code":
            value = re.sub(r"\s+", "", value)
            return value if len(value) >= 12 else ""

        if field_name in {"buyer_name", "seller_name"}:
            value = self._clean_name(value)
            return value if self._looks_like_name(value) else ""

        return value

    def _clean_name(self, value: str) -> str:
        value = str(value or "").strip()
        value = re.sub(r"^(?:购买方名称|销售方名称|购方名称|销方名称|名称)[:：\s]*", "", value)
        value = re.sub(
            r"(购方税号|销方税号|购买方税号|销售方税号|纳税人识别号|地址电话|开户地址|开户行及账号|密码区|货物或应税劳务).*$",
            "",
            value,
        )
        value = re.sub(r"[：:]+$", "", value)
        value = re.sub(r"\s{2,}", " ", value)
        value = value.strip(" ：:")
        while value and value[-1] in {"密", "备", "注", "区"}:
            value = value[:-1].rstrip(" ：:")
        if len(value) > 50:
            value = value[:50].rstrip(" ：:")
        return value.strip()

    def _extract_name_candidate(self, line: str) -> str:
        line = line.strip()
        if not line:
            return ""
        match = re.search(
            r"(?:购买方名称|销售方名称|购方名称|销方名称|名称)[：:\s]+(.+?)(?=(?:购方税号|销方税号|购买方税号|销售方税号|纳税人识别号|地址电话|开户地址|开户行及账号|密码区|货物或应税劳务|$))",
            line,
        )
        if match:
            return match.group(1).strip()
        if "称" in line and ("：" in line or ":" in line):
            return line.split("：", 1)[-1].split(":", 1)[-1].strip()
        return line

    def _looks_like_name(self, value: str) -> bool:
        value = self._clean_name(value)
        if not value or len(value) < 2 or len(value) > 50:
            return False
        if len(value) <= 2 and value != "个人":
            return False
        if self._extract_valid_date(value):
            return False
        if value in _NAME_STOPWORDS:
            return False
        if any(part in value for part in _NAME_BAD_PARTS):
            return False
        if any(keyword in value for keyword in ("日期", "地址", "电话", "开户", "规格型号", "圆整", "价税合计", "密码", "校验码")):
            return False
        if re.search(r"\d{6,}$", value):
            return False
        if re.fullmatch(r"[A-Za-z0-9\s\-+/*<>.]+", value):
            return False
        if any(keyword in value for keyword in _NAME_STOPWORDS if len(keyword) >= 3):
            return False
        if value == "个人":
            return True
        chinese_count = len(re.findall(r"[\u4e00-\u9fff]", value))
        if chinese_count >= 2:
            return True
        return any(hint in value for hint in _COMPANY_HINTS)

    def _extract_valid_date(self, text: str) -> str:
        for pattern in (
            _VALID_DATE_TEXT,
            _VALID_DATE_DIGITS,
            r"((?:19|20)\d{2}[-/.](?:0?[1-9]|1[0-2])[-/.](?:0?[1-9]|[12]\d|3[01]))",
        ):
            match = re.search(pattern, text)
            if not match:
                continue
            normalized = self._normalize_date(match.group(1))
            if normalized:
                return normalized
        return ""

    def _looks_like_receipt(self, cleaned: str, compact: str) -> bool:
        return (
            "收据" in cleaned
            or "收据" in compact
            or ("付款单位" in cleaned and "收款单位" in cleaned)
        )

    def _find_receipt_number(self, lines: list[str]) -> Optional[tuple[str, str]]:
        for idx, line in enumerate(lines[:12]):
            if not re.search(r"(收据编号|收据号|编号)", line):
                continue
            snippet_lines = lines[idx : min(len(lines), idx + 3)]
            snippet = " ".join(snippet_lines)
            for match in re.finditer(r"([A-Za-z]{1,4}\d{6,20}|\d{8,20})", snippet):
                candidate = re.sub(r"[^A-Za-z0-9]", "", match.group(1)).upper()
                if not candidate or self._is_plausible_date_digits(self._digits_only(candidate)):
                    continue
                return candidate, "high"
        return None

    def _find_label_value(
        self,
        lines: list[str],
        labels: tuple[str, ...],
        validator,
    ) -> str:
        label_expr = "|".join(map(re.escape, labels))
        for idx, line in enumerate(lines[:20]):
            match = re.search(rf"(?:{label_expr})[：:\s]+(.+)$", line)
            if match:
                value = match.group(1).strip()
                if validator(value):
                    return value

            if self._compact_text(line) not in {self._compact_text(label) for label in labels}:
                continue

            for next_line in lines[idx + 1 : min(len(lines), idx + 3)]:
                value = next_line.strip()
                if validator(value):
                    return value
        return ""

    def _find_dual_section_value(
        self,
        lines: list[str],
        labels: tuple[str, ...],
        index: int,
    ) -> str:
        buyer_header = next(
            (idx for idx, line in enumerate(lines[:12]) if any(term in line for term in ("购买方", "购方"))),
            -1,
        )
        seller_header = next(
            (
                idx
                for idx, line in enumerate(lines[:12])
                if idx > buyer_header and any(term in line for term in ("销售方", "销方"))
            ),
            -1,
        )
        if buyer_header < 0 or seller_header < 0 or seller_header - buyer_header > 3:
            return ""

        label_expr = "|".join(map(re.escape, labels))
        values: list[str] = []
        for line in lines[seller_header + 1 : min(len(lines), seller_header + 10)]:
            if self._is_table_header_line(line) or "货物名称" in line:
                break
            match = re.search(rf"(?:{label_expr})[：:\s]+(.+)$", line)
            if match:
                values.append(match.group(1).strip())
                if len(values) > index:
                    return values[index]
        return ""

    def _normalize_date(self, value: str) -> str:
        value = str(value or "").strip()
        if not value:
            return ""

        match = re.match(
            r"((?:19|20)\d{2})\s*年\s*(1[0-2]|0?[1-9])\s*月\s*(3[01]|[12]\d|0?[1-9])\s*日?",
            value,
        )
        if match:
            return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}-{int(match.group(3)):02d}"

        match = re.match(
            r"((?:19|20)\d{2})[-/.](1[0-2]|0?[1-9])[-/.](3[01]|[12]\d|0?[1-9])",
            value,
        )
        if match:
            return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}-{int(match.group(3)):02d}"

        match = re.match(
            r"((?:19|20)\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])",
            value,
        )
        if match:
            return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"

        return ""

    def _collect_money_values(self, lines: list[str]) -> list[str]:
        values: list[str] = []
        seen: set[str] = set()
        for line in lines:
            for pattern in (
                r"[¥￥]\s*(\d[\d,]*(?:\.\d{1,2})?)",
                r"(?<![\dA-Za-z,])(\d{1,12}(?:,\d{3})*(?:\.\d{1,2})?)(?![\dA-Za-z])",
            ):
                for match in re.finditer(pattern, line):
                    value = self._clean_value("amount", match.group(1))
                    if not value or value in seen:
                        continue
                    if not self._is_money_candidate(value, line, match.start(1), match.end(1)):
                        continue
                    seen.add(value)
                    values.append(value)
        return values

    def _nearest_money(
        self,
        lines: list[str],
        anchor_index: int,
        direction: int,
        limit: int,
        skip_values: Optional[set[str]] = None,
    ) -> str:
        skip_values = {item for item in (skip_values or set()) if item}
        step_range = range(1, limit + 1)
        for step in step_range:
            idx = anchor_index + step * direction
            if idx < 0 or idx >= len(lines):
                break
            values = self._collect_money_values([lines[idx]])
            candidates = reversed(values) if direction < 0 else values
            for value in candidates:
                if value in skip_values:
                    continue
                return value
        return ""

    def _looks_like_money(self, value: str) -> bool:
        return bool(re.fullmatch(r"\d[\d,]*(?:\.\d{1,2})?", value))

    def _same_money(self, left: str, right: str) -> bool:
        if not left or not right:
            return False
        try:
            return abs(float(left) - float(right)) < 0.01
        except ValueError:
            return False

    def _reconcile_money_fields(self, result: dict[str, tuple[str, str]]) -> None:
        amount = (result.get("amount") or ("", ""))[0]
        tax = (result.get("tax") or ("", ""))[0]
        total = (result.get("total") or ("", ""))[0]

        if amount and tax and total and self._same_money(self._add_money(amount, tax), total):
            return

        if amount and total:
            derived_tax = self._subtract_money(total, amount)
            total_amount = self._to_decimal(total)
            derived_tax_amount = self._to_decimal(derived_tax)
            tax_ratio_ok = bool(
                total_amount
                and derived_tax_amount is not None
                and derived_tax_amount > Decimal("0.00")
                and total_amount > Decimal("0.00")
                and Decimal("0.00") <= (derived_tax_amount / total_amount) <= Decimal("0.20")
            )
            if derived_tax and tax_ratio_ok and (not tax or not self._same_money(derived_tax, tax)):
                tax_confidence = "high" if tax else "medium"
                result["tax"] = (derived_tax, tax_confidence)
                tax = derived_tax

        if tax and total:
            derived_amount = self._subtract_money(total, tax)
            if (
                derived_amount
                and self._is_plausible_tax_value(tax, total=total)
                and (not amount or self._to_decimal(amount) is None or self._to_decimal(amount) > self._to_decimal(total))
            ):
                amount_confidence = "high" if amount else "medium"
                result["amount"] = (derived_amount, amount_confidence)
                amount = derived_amount

        if amount and tax and not total:
            derived_total = self._add_money(amount, tax)
            if derived_total:
                result["total"] = (derived_total, "medium")

    def _digits_only(self, value: str) -> str:
        return re.sub(r"\D", "", value or "")

    def _is_plausible_date_digits(self, value: str) -> bool:
        if not re.fullmatch(r"\d{8}", value):
            return False
        return bool(
            re.fullmatch(
                r"(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])",
                value,
            )
        )

    def _looks_like_taxi_invoice(self, cleaned: str, compact: str) -> bool:
        lower_text = cleaned.lower()
        return (
            "taxi" in lower_text
            or "出租" in cleaned
            or "fare" in lower_text
            or "fuel oil surcharge" in lower_text
            or "carno" in compact.lower()
        )

    def _looks_like_train_ticket(self, cleaned: str, compact: str) -> bool:
        return (
            "火车票" in cleaned
            or "铁路" in cleaned
            or "中国铁路" in cleaned
            or ("车次" in cleaned and ("出发" in cleaned or "到达" in cleaned or "乘车" in cleaned))
            or bool(re.search(r"(?<![A-Za-z0-9])[GCDZTK]\d{1,5}(?![A-Za-z0-9])", cleaned, re.IGNORECASE))
        )

    def _looks_like_person_name(self, value: str) -> bool:
        value = str(value or "").strip()
        if not 2 <= len(value) <= 20:
            return False
        if self._extract_valid_date(value):
            return False
        if re.search(r"\d{4,}", value):
            return False
        return bool(re.search(r"[\u4e00-\u9fa5A-Za-z]", value))

    def _find_train_station_pair(self, lines: list[str]) -> Optional[tuple[str, str]]:
        for line in lines[:30]:
            compact_line = self._compact_text(line)
            match = re.search(r"([\u4e00-\u9fa5]{2,12}站?)\s*(?:-|—|→|->|至|到)\s*([\u4e00-\u9fa5]{2,12}站?)", compact_line)
            if match:
                return match.group(1), match.group(2)
            match = re.search(r"出发站[：:\s]*([\u4e00-\u9fa5]{2,12}站?).{0,12}到达站[：:\s]*([\u4e00-\u9fa5]{2,12}站?)", compact_line)
            if match:
                return match.group(1), match.group(2)
        return None

    def _find_labeled_time(self, lines: list[str], labels: tuple[str, ...]) -> Optional[tuple[str, str]]:
        label_expr = "|".join(map(re.escape, labels))
        for line in lines[:40]:
            match = re.search(rf"(?:{label_expr})[：:\s]*(\d{{1,2}}[:：]\d{{2}}(?::\d{{2}})?)", line, re.IGNORECASE)
            if match:
                return match.group(1).replace("：", ":"), "high"
        return None

    def _find_labeled_money(self, lines: list[str], labels: tuple[str, ...]) -> Optional[tuple[str, str]]:
        label_expr = "|".join(map(re.escape, labels))
        for line in lines[:40]:
            if not re.search(label_expr, line, re.IGNORECASE):
                continue
            values = self._collect_money_values([line])
            if values:
                return values[-1], "high"
        return None

    def _find_labeled_money_multiline(self, lines: list[str], labels: tuple[str, ...]) -> Optional[tuple[str, str]]:
        """查找标签行及其前后相邻行中的金额（处理标签和金额分行的情况）。"""
        label_expr = "|".join(map(re.escape, labels))
        for idx, line in enumerate(lines[:40]):
            if not re.search(label_expr, line, re.IGNORECASE):
                continue
            # 先查同行
            values = self._collect_money_values([line])
            if values:
                return values[-1], "high"
            # 再查下一行
            if idx + 1 < len(lines):
                next_line = lines[idx + 1]
                match = re.search(r"[¥￥]?\s*(\d[\d,]*(?:\.\d{1,2})?)", next_line)
                if match:
                    raw = match.group(1).replace(",", "")
                    if self._looks_like_money(raw):
                        return raw, "medium"
            # 再查上一行（处理值在标签前面的布局）
            if idx > 0:
                prev_values = self._collect_money_values([lines[idx - 1]])
                if prev_values:
                    return prev_values[-1], "medium"
        return None

    def _extract_role_name(self, line: str, role_terms: tuple[str, ...]) -> str:
        role_expr = "|".join(map(re.escape, role_terms))
        match = re.search(
            rf"(?:{role_expr})(?:名称)?[：:\s]+(.+?)(?=(?:购买方|购方|销售方|销方)?(?:税号|纳税人识别号)|地址电话|地址|电话|开户地址|开户行及账号|密码区|货物或应税劳务|$)",
            line,
        )
        if not match:
            return ""
        return self._clean_name(match.group(1))

    def _find_summary_money_fields(
        self,
        lines: list[str],
        known_total: str = "",
    ) -> dict[str, tuple[str, str]]:
        for idx, line in enumerate(lines):
            if not self._looks_like_summary_line(line):
                continue

            snippets = [line]
            if idx + 1 < len(lines):
                snippets.append(" ".join(lines[idx : idx + 2]))
            if idx > 0:
                snippets.append(" ".join(lines[idx - 1 : idx + 1]))

            for snippet in snippets:
                values = self._collect_money_values([snippet])
                triplet = self._pick_summary_triplet(values, known_total=known_total)
                if not triplet:
                    if len(values) == 1 and self._looks_like_total_marker(snippet) and snippet.startswith(line):
                        return {"total": (values[-1], "high")}
                    continue
                amount, tax, total = triplet
                return {
                    "amount": (amount, "high"),
                    "tax": (tax, "high"),
                    "total": (total, "high"),
                }
        return {}

    def _looks_like_summary_line(self, line: str) -> bool:
        compact_line = self._compact_text(line)
        if not compact_line:
            return False
        if self._is_table_header_line(line):
            return False
        return (
            "合计大写" in compact_line
            or "价税合计" in compact_line
            or "价税合计(小写)" in compact_line
            or "价税合计小写" in compact_line
            or compact_line in {"合", "计", "合计", "小写"}
            or compact_line.startswith("合计")
        )

    def _pick_summary_triplet(
        self,
        values: list[str],
        known_total: str = "",
    ) -> Optional[tuple[str, str, str]]:
        if len(values) < 3:
            return None
        if known_total:
            for idx, value in enumerate(values):
                if self._same_money(value, known_total) and idx >= 2:
                    return values[idx - 2], values[idx - 1], value
        return values[-3], values[-2], values[-1]

    def _is_table_header_line(self, line: str) -> bool:
        compact_line = self._compact_text(line)
        # 这些词单独出现时可明确判定为表头（不含"金额"，因为它也常作字段标签）
        unambiguous_headers = {
            "项目",
            "货物名称",
            "商品名称",
            "规格",
            "型号",
            "单位",
            "数量",
            "单价",
            "税率",
            "税额",
            "价税合计",
        }
        if compact_line in unambiguous_headers:
            return True
        return sum(token in compact_line for token in ("金额", "税率", "税额", "单价", "数量")) >= 2

    def _is_money_candidate(self, value: str, line: str, start: int, end: int) -> bool:
        digits = self._digits_only(value)
        if not digits:
            return False
        if len(digits) == 8 and self._is_plausible_date_digits(digits):
            return False
        if "." not in value and len(digits) < 3:
            return False

        left = line[start - 1] if start > 0 else ""
        right = line[end] if end < len(line) else ""
        if (left and left in "-/.年月日") or (right and right in "-/.年月日"):
            return False

        try:
            numeric = float(value)
        except ValueError:
            return False
        if 0 < numeric < 1:
            return False
        return True

    def _find_labeled_invoice_number(self, lines: list[str]) -> Optional[tuple[str, str]]:
        for idx, line in enumerate(lines[:20]):
            if not re.search(r"(发票号码|票号|No\.?|NO\.?)", line, re.IGNORECASE):
                continue
            if "代码" in line and "号码" not in line:
                continue
            candidates = self._extract_invoice_number_candidates(" ".join(lines[idx : idx + 3]))
            for candidate in candidates:
                if self._is_plausible_invoice_number(candidate, line=line):
                    confidence = "high" if idx < 12 else "medium"
                    return candidate, confidence
        return None

    def _extract_invoice_number_candidates(self, text: str) -> list[str]:
        candidates: list[str] = []
        for pattern in (
            r"(?:发票号码|票号|No\.?|NO\.?)" + _SEP + r"([A-Za-z]?\d{8,20})",
            r"(?<![A-Za-z0-9])([A-Za-z]?\d{8,20})(?![A-Za-z0-9])",
        ):
            for match in re.finditer(pattern, text, re.IGNORECASE):
                candidate = re.sub(r"[^A-Za-z0-9]", "", match.group(1)).upper()
                if candidate and candidate not in candidates:
                    candidates.append(candidate)
        return candidates

    def _is_plausible_invoice_number(self, value: str, line: str = "") -> bool:
        value = re.sub(r"[^A-Za-z0-9]", "", value or "").upper()
        if not re.fullmatch(r"[A-Z]?\d{8,20}", value):
            return False
        digits = self._digits_only(value)
        if len(digits) == 8 and self._is_plausible_date_digits(digits):
            return False
        if any(keyword in line for keyword in ("机器编号", "校验码", "税号", "账号", "代码")) and "号码" not in line:
            return False
        return True

    def _find_labeled_invoice_date(self, lines: list[str]) -> Optional[tuple[str, str]]:
        for idx, line in enumerate(lines[:20]):
            normalized = line.strip().lower()
            if "开票日期" not in line and line.strip() != "日期" and normalized != "date":
                continue
            for candidate in lines[idx : idx + 4]:
                date_value = self._extract_valid_date(candidate)
                if date_value:
                    confidence = "high" if idx < 12 else "medium"
                    return date_value, confidence
            merged = " ".join(lines[idx : idx + 3])
            date_value = self._extract_valid_date(merged)
            if date_value:
                return date_value, "medium"
        return None

    def _derive_money_fields(
        self,
        lines: list[str],
        cleaned: str,
        compact: str,
        known_total: str = "",
        known_amount: str = "",
        known_tax: str = "",
    ) -> dict[str, tuple[str, str]]:
        result: dict[str, tuple[str, str]] = {}
        body_values = self._collect_body_money_values(lines, skip_values={known_total})
        has_tax_signal = self._has_tax_signal(lines)

        if known_total and known_tax and not known_amount:
            target_amount = self._subtract_money(known_total, known_tax)
            if target_amount and self._is_plausible_tax_value(known_tax, amount=target_amount, total=known_total):
                subset = self._pick_money_subset(body_values, target_amount)
                if subset or not has_tax_signal:
                    confidence = "high" if subset else "medium"
                    result["amount"] = (target_amount, confidence)

        if known_total and known_amount and not known_tax:
            derived_tax = self._subtract_money(known_total, known_amount)
            if (
                derived_tax
                and not self._same_money(derived_tax, "0.00")
                and self._is_plausible_tax_value(derived_tax, amount=known_amount, total=known_total)
            ):
                result["tax"] = (derived_tax, "medium")

        if known_total and not known_amount:
            if has_tax_signal:
                amount_tax_pair = self._find_amount_tax_pair_by_total(body_values, known_total)
                if amount_tax_pair:
                    amount_value, tax_value = amount_tax_pair
                    result.setdefault("amount", (amount_value, "medium"))
                    if tax_value and not known_tax:
                        result.setdefault("tax", (tax_value, "medium"))
            elif not self._looks_like_taxi_invoice(cleaned, compact):
                result.setdefault("amount", (known_total, "low"))

        if known_amount and known_tax and not known_total:
            derived_total = self._add_money(known_amount, known_tax)
            if derived_total:
                result["total"] = (derived_total, "medium")

        return result

    def _collect_rate_amount_tax_pairs(
        self,
        lines: list[str],
        known_total: str = "",
    ) -> list[tuple[str, str]]:
        pairs: list[tuple[str, str]] = []
        seen: set[tuple[str, str]] = set()
        for idx, line in enumerate(lines):
            if not self._looks_like_rate_anchor(line):
                continue
            amount = self._nearest_money(lines, idx, direction=-1, limit=3, skip_values={known_total})
            tax = self._nearest_money(
                lines,
                idx,
                direction=1,
                limit=3,
                skip_values={known_total, amount},
            )
            if not amount or not tax:
                continue
            amount_decimal = self._to_decimal(amount)
            tax_decimal = self._to_decimal(tax)
            if amount_decimal is None or tax_decimal is None:
                continue
            if amount_decimal <= 0 or tax_decimal < 0 or tax_decimal >= amount_decimal:
                continue
            if not self._is_plausible_tax_value(tax, amount=amount):
                continue
            key = (amount, tax)
            if key in seen:
                continue
            seen.add(key)
            pairs.append(key)
        return pairs

    def _collect_rate_amount_tax_candidates(
        self,
        lines: list[str],
        known_total: str = "",
    ) -> tuple[list[str], list[str]]:
        amount_candidates: list[str] = []
        tax_candidates: list[str] = []
        seen_amounts: set[str] = set()
        seen_taxes: set[str] = set()

        for idx, line in enumerate(lines):
            if not self._looks_like_rate_anchor(line):
                continue
            amount = self._nearest_money(lines, idx, direction=-1, limit=3, skip_values={known_total})
            tax = self._nearest_money(
                lines,
                idx,
                direction=1,
                limit=3,
                skip_values={known_total, amount},
            )
            if amount and amount not in seen_amounts:
                seen_amounts.add(amount)
                amount_candidates.append(amount)
            if tax and tax not in seen_taxes:
                seen_taxes.add(tax)
                tax_candidates.append(tax)
        return amount_candidates, tax_candidates

    def _repair_amount_tax_from_candidates(
        self,
        amount_candidates: list[str],
        tax_candidates: list[str],
        known_total: str,
    ) -> dict[str, tuple[str, str]]:
        result: dict[str, tuple[str, str]] = {}

        total_amount = self._to_decimal(known_total)
        if total_amount is None or total_amount <= Decimal("0.00"):
            return result

        valid_amounts = [
            value
            for value in amount_candidates
            if (self._to_decimal(value) or Decimal("0.00")) > Decimal("0.00")
            and (self._to_decimal(value) or Decimal("0.00")) < total_amount
        ]
        valid_taxes = [
            value
            for value in tax_candidates
            if self._is_plausible_tax_value(value, total=known_total)
        ]

        amount_sum = self._sum_money_values(valid_amounts)
        if amount_sum:
            derived_tax = self._subtract_money(known_total, amount_sum)
            if self._same_money(derived_tax, "0.00"):
                result["amount"] = (amount_sum, "medium")
                result["total"] = (known_total, "high")
                return result
            if derived_tax and self._is_plausible_tax_value(derived_tax, amount=amount_sum, total=known_total):
                result["amount"] = (amount_sum, "medium")
                result["tax"] = (derived_tax, "medium")
                result["total"] = (known_total, "high")
                return result

        tax_sum = self._sum_money_values(valid_taxes)
        if tax_sum:
            derived_amount = self._subtract_money(known_total, tax_sum)
            if (
                derived_amount
                and self._to_decimal(derived_amount) is not None
                and self._to_decimal(derived_amount) > Decimal("0.00")
                and self._is_plausible_tax_value(tax_sum, amount=derived_amount, total=known_total)
            ):
                result["amount"] = (derived_amount, "medium")
                result["tax"] = (tax_sum, "medium")
                result["total"] = (known_total, "high")
                return result

        return result

    def _looks_like_rate_anchor(self, line: str) -> bool:
        compact_line = self._compact_text(line)
        if not compact_line or self._is_table_header_line(line):
            return False
        if any(symbol in compact_line for symbol in ("￥", "¥", "$")):
            return False
        normalized_line = compact_line.replace("％", "%")
        if re.fullmatch(r"(?:\d{1,2}%|%\d{1,2})", normalized_line):
            return self._digits_only(normalized_line) in {"1", "3", "5", "6", "9", "10", "13"}
        if normalized_line in {
            "0.01",
            "0.03",
            "0.05",
            "0.06",
            "0.09",
            "0.10",
            "0.13",
            "1.00",
            "3.00",
            "5.00",
            "6.00",
            "9.00",
            "10.00",
            "13.00",
        }:
            return True
        digits = self._digits_only(normalized_line)
        if digits in {"100", "300", "500", "600", "900", "1000", "1300"}:
            return True
        return "税率" in compact_line

    def _collect_body_money_values(
        self,
        lines: list[str],
        skip_values: Optional[set[str]] = None,
    ) -> list[str]:
        skip_values = {item for item in (skip_values or set()) if item}
        summary_index = self._find_summary_index(lines)
        start_index = self._find_table_start_index(lines)
        scoped_lines = lines[start_index:summary_index] if summary_index > start_index else lines[start_index:]

        values: list[str] = []
        for line in scoped_lines:
            if self._is_table_header_line(line) or self._looks_like_total_marker(line):
                continue
            for value in self._collect_money_values([line]):
                if value in skip_values:
                    continue
                values.append(value)
        return values

    def _find_summary_index(self, lines: list[str]) -> int:
        for idx, line in enumerate(lines):
            if self._looks_like_total_marker(line):
                return idx
        return len(lines)

    def _find_table_start_index(self, lines: list[str]) -> int:
        for idx, line in enumerate(lines):
            if self._is_table_header_line(line):
                return idx
        return 0

    def _looks_like_total_marker(self, line: str) -> bool:
        compact_line = self._compact_text(line)
        return (
            "价税合计" in compact_line
            or "小写" in compact_line
            or compact_line == "合计"
            or compact_line.startswith("合计")
        )

    def _pick_money_subset(self, values: list[str], target: str) -> list[str]:
        target_cents = self._money_to_cents(target)
        if target_cents is None or target_cents <= 0:
            return []

        candidate_values = sorted(
            [
                amount_cents
                for amount_cents in (self._money_to_cents(value) for value in values)
                if amount_cents is not None and 0 < amount_cents <= target_cents
            ],
            reverse=True,
        )
        if not candidate_values:
            return []

        states: dict[int, list[int]] = {0: []}
        for amount_cents in candidate_values:
            next_states = dict(states)
            for subtotal, picked in states.items():
                new_total = subtotal + amount_cents
                if new_total > target_cents:
                    continue
                candidate_path = picked + [amount_cents]
                current_path = next_states.get(new_total)
                if current_path is None or len(candidate_path) < len(current_path):
                    next_states[new_total] = candidate_path
            states = next_states
            if target_cents in states:
                break

        subset = states.get(target_cents, [])
        return [self._format_cents(value) for value in subset]

    def _find_amount_tax_pair_by_total(
        self,
        values: list[str],
        total: str,
    ) -> Optional[tuple[str, str]]:
        target_total = self._to_decimal(total)
        if target_total is None or target_total <= Decimal("0.00"):
            return None

        normalized_values: list[str] = []
        seen: set[str] = set()
        for value in values:
            cleaned_value = self._clean_value("amount", value)
            if not cleaned_value or cleaned_value in seen:
                continue
            numeric = self._to_decimal(cleaned_value)
            if numeric is None or numeric <= Decimal("0.00") or numeric >= target_total:
                continue
            seen.add(cleaned_value)
            normalized_values.append(cleaned_value)

        best_pair: Optional[tuple[str, str]] = None
        best_amount = Decimal("-1")
        best_tax = Decimal("-1")
        for left_index, left in enumerate(normalized_values):
            for right in normalized_values[left_index + 1 :]:
                left_amount = self._to_decimal(left)
                right_amount = self._to_decimal(right)
                if left_amount is None or right_amount is None:
                    continue
                amount_value, tax_value = (left, right) if left_amount >= right_amount else (right, left)
                amount_numeric = max(left_amount, right_amount)
                tax_numeric = min(left_amount, right_amount)
                pair_total = self._add_money(amount_value, tax_value)
                if not pair_total or not self._same_money(pair_total, total):
                    continue
                if not self._is_plausible_tax_value(tax_value, amount=amount_value, total=total):
                    continue
                if amount_numeric > best_amount or (amount_numeric == best_amount and tax_numeric > best_tax):
                    best_pair = (amount_value, tax_value)
                    best_amount = amount_numeric
                    best_tax = tax_numeric
        return best_pair

    def _sum_money_values(self, values: list[str]) -> str:
        total = Decimal("0.00")
        seen_any = False
        for value in values:
            amount = self._to_decimal(value)
            if amount is None:
                continue
            total += amount
            seen_any = True
        return self._format_decimal(total) if seen_any else ""

    def _add_money(self, left: str, right: str) -> str:
        left_amount = self._to_decimal(left)
        right_amount = self._to_decimal(right)
        if left_amount is None or right_amount is None:
            return ""
        return self._format_decimal(left_amount + right_amount)

    def _subtract_money(self, left: str, right: str) -> str:
        left_amount = self._to_decimal(left)
        right_amount = self._to_decimal(right)
        if left_amount is None or right_amount is None:
            return ""
        diff = left_amount - right_amount
        if diff < Decimal("0.00"):
            return ""
        return self._format_decimal(diff)

    def _max_money_value(self, values: list[str]) -> str:
        best_value = ""
        best_amount = Decimal("-1")
        for value in values:
            amount = self._to_decimal(value)
            if amount is None or amount <= best_amount:
                continue
            best_amount = amount
            best_value = self._format_decimal(amount)
        return best_value

    def _has_tax_signal(self, lines: list[str]) -> bool:
        for line in lines:
            if "税额" in line or "税率" in line or self._looks_like_rate_anchor(line):
                return True
        return False

    def _is_plausible_tax_value(
        self,
        tax: str,
        amount: str = "",
        total: str = "",
    ) -> bool:
        tax_amount = self._to_decimal(tax)
        if tax_amount is None or tax_amount < Decimal("0.00"):
            return False
        if tax_amount == Decimal("0.00"):
            return True

        for base_value in (amount, total):
            base_amount = self._to_decimal(base_value)
            if base_amount is None or base_amount <= Decimal("0.00"):
                continue
            ratio = tax_amount / base_amount
            if Decimal("0.00") < ratio <= Decimal("0.20"):
                return True
        return False

    def _to_decimal(self, value: str) -> Optional[Decimal]:
        text = self._clean_value("amount", value)
        if not text:
            return None
        try:
            return Decimal(text)
        except InvalidOperation:
            return None

    def _format_decimal(self, value: Decimal) -> str:
        return format(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP), "f")

    def _money_to_cents(self, value: str) -> Optional[int]:
        amount = self._to_decimal(value)
        if amount is None:
            return None
        quantized = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return int((quantized * 100).to_integral_value(rounding=ROUND_HALF_UP))

    def _format_cents(self, cents: int) -> str:
        return self._format_decimal(Decimal(cents) / Decimal("100"))

    def _empty_result(self) -> dict:
        return {
            "is_invoice": False,
            "confidence": "none",
            "field_count": 0,
            "fields": {},
            "invoice_category": "unknown",
            "schema_name": get_invoice_schema("unknown").name,
        }
