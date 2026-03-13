"""
发票结构化信息提取模块。
"""

from __future__ import annotations

import re
from typing import Optional


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
}

INVOICE_TYPE_KEYWORDS: list[tuple[str, str]] = [
    ("增值税电子专用发票", "增值税电子专用发票"),
    ("增值税电子普通发票", "增值税电子普通发票"),
    ("增值税专用发票", "增值税专用发票"),
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

FIELD_LABELS: dict[str, str] = {
    "invoice_code": "发票代码",
    "invoice_number": "发票号码",
    "invoice_date": "开票日期",
    "amount": "金额（不含税）",
    "tax": "税额",
    "total": "价税合计",
    "buyer_name": "购买方名称",
    "seller_name": "销售方名称",
    "buyer_tax_id": "购买方税号",
    "seller_tax_id": "销售方税号",
    "invoice_type": "发票类型",
    "machine_number": "机器编号",
    "check_code": "校验码",
}

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

        fields: dict[str, dict] = {}
        for field_name, patterns in FIELD_PATTERNS.items():
            result = self._try_field(field_name, cleaned, compact, patterns)
            if not result:
                continue
            value, confidence = result
            value = self._clean_value(field_name, value)
            if value:
                fields[field_name] = self._pack_field(field_name, value, confidence)

        self._apply_fallbacks(cleaned, compact, lines, fields)

        invoice_type = self._detect_invoice_type(cleaned, compact)
        if invoice_type:
            fields["invoice_type"] = self._pack_field("invoice_type", invoice_type, "high")

        score = self._score_invoice(fields, compact)
        has_money = "amount" in fields or "total" in fields
        has_header = "invoice_code" in fields or "invoice_number" in fields
        has_parties = "buyer_name" in fields or "seller_name" in fields

        is_invoice = (
            {"invoice_code", "invoice_number"}.issubset(fields)
            or (
                "invoice_date" in fields
                and has_money
                and ("invoice_type" in fields or "check_code" in fields or has_header or has_parties)
            )
            or (score >= 5)
        )

        if not is_invoice:
            overall_confidence = "none"
        elif score >= 7:
            overall_confidence = "high"
        elif score >= 5:
            overall_confidence = "medium"
        else:
            overall_confidence = "low"

        return {
            "is_invoice": is_invoice,
            "confidence": overall_confidence,
            "field_count": len(fields),
            "fields": fields,
        }

    def _apply_fallbacks(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
        fields: dict[str, dict],
    ) -> None:
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
        ]

        for field_name, extractor in ordered_extractors:
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
        for field_name, (value, confidence) in money_fields.items():
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
        match = re.search(r"\bNo\.?\s*([A-Za-z]?\d{8,20})\b", cleaned, re.IGNORECASE)
        if match:
            return match.group(1), "high"

        top_lines = lines[:12]
        counts: dict[str, int] = {}
        first_pos: dict[str, int] = {}
        for idx, line in enumerate(top_lines):
            digits = self._digits_only(line)
            if len(digits) != 8 or self._is_plausible_date_digits(digits):
                continue
            counts[digits] = counts.get(digits, 0) + 1
            first_pos.setdefault(digits, idx)
        if counts:
            best = sorted(counts, key=lambda value: (-counts[value], first_pos[value]))[0]
            confidence = "high" if counts[best] >= 2 else "medium"
            return best, confidence

        match = re.search(r"(?<!\d)(\d{8})(?!\d)", compact)
        if match and not self._is_plausible_date_digits(match.group(1)):
            return match.group(1), "low"
        return None

    def _find_invoice_date(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        for idx, line in enumerate(lines):
            if "开票日期" in line or line == "日期" or line.lower() == "date":
                for candidate in lines[idx : idx + 4]:
                    date_value = self._extract_valid_date(candidate)
                    if date_value:
                        return date_value, "high"

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
        return self._find_party_name(compact, lines, ("购买方", "购方"), prefer="first")

    def _find_seller_name(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        return self._find_party_name(compact, lines, ("销售方", "销方"), prefer="last")

    def _find_buyer_tax_id(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        return self._find_party_tax_id(compact, lines, ("购买方", "购方"))

    def _find_seller_tax_id(
        self,
        cleaned: str,
        compact: str,
        lines: list[str],
    ) -> Optional[tuple[str, str]]:
        return self._find_party_tax_id(compact, lines, ("销售方", "销方"))

    def _find_party_name(
        self,
        compact: str,
        lines: list[str],
        role_terms: tuple[str, ...],
        prefer: str,
    ) -> Optional[tuple[str, str]]:
        marker_indexes = [idx for idx, line in enumerate(lines) if any(term in line for term in role_terms)]
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

        summary_amount_tax = self._find_amount_tax_from_summary(lines, known_total=(total or ("", ""))[0])
        for field_name, value in summary_amount_tax.items():
            result.setdefault(field_name, value)

        explicit_amount = self._find_explicit_amount(lines)
        if explicit_amount:
            result.setdefault("amount", explicit_amount)

        explicit_tax = self._find_explicit_tax(lines, known_total=(total or ("", ""))[0], known_amount=(result.get("amount") or ("", ""))[0])
        if explicit_tax:
            result.setdefault("tax", explicit_tax)

        if "amount" not in result or "tax" not in result:
            amount_tax = self._find_amount_tax_from_rate(
                lines,
                known_total=(result.get("total") or ("", ""))[0],
            )
            for field_name, value in amount_tax.items():
                result.setdefault(field_name, value)

        if self._looks_like_taxi_invoice(cleaned, compact):
            if "amount" in result and "total" not in result:
                result["total"] = result["amount"]
            elif "total" in result and "amount" not in result:
                result["amount"] = result["total"]

        return result

    def _find_total(self, lines: list[str], compact: str) -> Optional[tuple[str, str]]:
        summary_fields = self._find_summary_money_fields(lines)
        if "total" in summary_fields:
            return summary_fields["total"]

        for idx, line in enumerate(lines):
            if self._is_table_header_line(line):
                continue
            if "小写" in line:
                values = self._collect_money_values([line])
                if values:
                    return values[-1], "high"
            if "价税合计" in line:
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

    def _find_amount_tax_from_rate(
        self,
        lines: list[str],
        known_total: str = "",
    ) -> dict[str, tuple[str, str]]:
        result: dict[str, tuple[str, str]] = {}
        explicit_rate_indexes = [
            idx for idx, line in enumerate(lines) if re.search(r"(?:\d{1,2}%|%\d{1,2})", line)
        ]
        fallback_rate_indexes = [
            idx for idx, line in enumerate(lines) if "税率" in line and idx not in explicit_rate_indexes
        ]

        for idx in explicit_rate_indexes or fallback_rate_indexes:
            line = lines[idx]
            if not re.search(r"(?:\d{1,2}%|%\d{1,2}|税率)", line):
                continue

            amount = self._nearest_money(lines, idx, direction=-1, limit=3)
            if amount and "amount" not in result:
                result["amount"] = (amount, "high")

            tax = self._nearest_money(
                lines,
                idx,
                direction=1,
                limit=5,
                skip_values={amount or "", known_total},
            )
            if tax:
                result["tax"] = (tax, "high")
            break
        return result

    def _find_explicit_amount(self, lines: list[str]) -> Optional[tuple[str, str]]:
        for idx, line in enumerate(lines):
            if "金额" not in line and line.lower() != "fare":
                continue
            if self._is_table_header_line(line):
                continue
            prev_values = self._collect_money_values([lines[idx - 1]]) if idx > 0 else []
            values = self._collect_money_values(lines[max(0, idx - 1) : idx + 4])
            if values:
                if prev_values and prev_values[0] in values:
                    return prev_values[0], "medium"
                return values[0], "medium"
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
        if ("invoice" in lower_text or "nvoice" in lower_text) and ("taxi" in lower_text or "出租" in cleaned):
            return "出租车发票"
        if "价税合计" in compact or "校验码" in compact or "发票专用章" in compact:
            return "增值税发票"
        return None

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
            "label": FIELD_LABELS.get(field_name, field_name),
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

        if field_name in {"amount", "tax", "total"}:
            value = value.replace(",", "").replace("，", "")
            value = value.replace("¥", "").replace("￥", "").strip()
            if not self._looks_like_money(value):
                return ""
            return value

        if field_name == "invoice_date":
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
                r"(?<![\dA-Za-z])(\d{1,12}(?:\.\d{1,2})?)(?![\dA-Za-z])",
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
            for value in values:
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
            if idx > 0:
                snippets.append(" ".join(lines[idx - 1 : idx + 1]))
            if idx + 1 < len(lines):
                snippets.append(" ".join(lines[idx : idx + 2]))

            for snippet in snippets:
                values = self._collect_money_values([snippet])
                triplet = self._pick_summary_triplet(values, known_total=known_total)
                if not triplet:
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
        return all(keyword in line for keyword in ("金额", "税率", "税额", "价税合计"))

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

    def _empty_result(self) -> dict:
        return {
            "is_invoice": False,
            "confidence": "none",
            "field_count": 0,
            "fields": {},
        }
