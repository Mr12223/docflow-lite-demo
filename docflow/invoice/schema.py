"""Invoice field schema registry.

This module centralizes invoice field labels, display order and type-specific
field sets so new invoice-like documents can be added without touching the
frontend or merge layer.
"""

from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
import re
import tempfile

from docflow.paths import INVOICE_TEMPLATES_PATH


@dataclass(frozen=True)
class InvoiceTypeSchema:
    key: str
    name: str
    fields: tuple[str, ...]
    aliases: tuple[str, ...] = ()


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
    "passenger_name": "乘车人",
    "train_no": "车次",
    "departure_station": "出发站",
    "arrival_station": "到达站",
    "travel_date": "乘车日期",
    "seat_class": "席别",
    "fare": "票价",
    "taxi_car_no": "车牌号",
    "taxi_start_time": "上车时间",
    "taxi_end_time": "下车时间",
    "quota_amount": "定额金额",
    "fuel_oil_surcharge": "燃油附加费",
}

INVOICE_FIELD_ORDER: tuple[str, ...] = (
    "invoice_type",
    "invoice_code",
    "invoice_number",
    "invoice_date",
    "total",
    "amount",
    "tax",
    "buyer_name",
    "buyer_tax_id",
    "seller_name",
    "seller_tax_id",
    "machine_number",
    "check_code",
    "passenger_name",
    "train_no",
    "departure_station",
    "arrival_station",
    "travel_date",
    "seat_class",
    "fare",
    "taxi_car_no",
    "taxi_start_time",
    "taxi_end_time",
    "quota_amount",
    "fuel_oil_surcharge",
)

BUILTIN_INVOICE_TYPE_SCHEMAS: dict[str, InvoiceTypeSchema] = {
    "vat": InvoiceTypeSchema(
        key="vat",
        name="增值税发票",
        aliases=("增值税", "增值税发票", "增值税专用发票", "增值税普通发票"),
        fields=(
            "invoice_type",
            "invoice_code",
            "invoice_number",
            "invoice_date",
            "total",
            "amount",
            "tax",
            "buyer_name",
            "buyer_tax_id",
            "seller_name",
            "seller_tax_id",
            "machine_number",
            "check_code",
        ),
    ),
    "electronic": InvoiceTypeSchema(
        key="electronic",
        name="电子发票",
        aliases=("电子发票", "全国统一电子发票", "增值税电子普通发票", "增值税电子专用发票"),
        fields=(
            "invoice_type",
            "invoice_number",
            "invoice_date",
            "total",
            "amount",
            "tax",
            "buyer_name",
            "seller_name",
            "machine_number",
            "check_code",
        ),
    ),
    "normal": InvoiceTypeSchema(
        key="normal",
        name="普通发票",
        aliases=("普通发票", "增值税普通发票"),
        fields=("invoice_type", "invoice_number", "invoice_date", "total", "amount", "buyer_name", "seller_name"),
    ),
    "taxi": InvoiceTypeSchema(
        key="taxi",
        name="出租车发票",
        aliases=("出租车发票", "出租车专用发票", "Taxi", "TAXI"),
        fields=("invoice_type", "invoice_number", "invoice_date", "taxi_car_no", "taxi_start_time", "taxi_end_time", "fare", "total", "fuel_oil_surcharge"),
    ),
    "train": InvoiceTypeSchema(
        key="train",
        name="火车票",
        aliases=("火车票", "铁路电子客票", "中国铁路"),
        fields=("invoice_type", "passenger_name", "train_no", "departure_station", "arrival_station", "travel_date", "seat_class", "fare"),
    ),
    "quota": InvoiceTypeSchema(
        key="quota",
        name="定额发票",
        aliases=("定额发票",),
        fields=("invoice_type", "invoice_number", "invoice_date", "quota_amount", "total"),
    ),
    "receipt": InvoiceTypeSchema(
        key="receipt",
        name="收据",
        aliases=("收据", "收款收据", "Receipt", "RECEIPT"),
        fields=("invoice_type", "invoice_number", "invoice_date", "total", "buyer_name", "seller_name"),
    ),
    "toll": InvoiceTypeSchema(
        key="toll",
        name="通行费电子发票",
        aliases=("通行费电子发票", "通行费"),
        fields=("invoice_type", "invoice_number", "invoice_date", "total", "amount", "tax", "buyer_name", "seller_name"),
    ),
    "vehicle": InvoiceTypeSchema(
        key="vehicle",
        name="机动车销售统一发票",
        aliases=("机动车销售统一发票",),
        fields=(
            "invoice_type",
            "invoice_code",
            "invoice_number",
            "invoice_date",
            "total",
            "amount",
            "tax",
            "buyer_name",
            "buyer_tax_id",
            "seller_name",
            "seller_tax_id",
        ),
    ),
    "unknown": InvoiceTypeSchema(
        key="unknown",
        name="其他票据",
        aliases=(),
        fields=("invoice_type", "invoice_number", "invoice_date", "total", "amount"),
    ),
}


def _normalize_schema_key(value: str) -> str:
    normalized = re.sub(r"[^a-zA-Z0-9_]+", "_", str(value or "").strip().lower()).strip("_")
    return normalized[:40]


def _clean_str_list(values) -> tuple[str, ...]:
    if not isinstance(values, (list, tuple)):
        return ()
    cleaned = []
    seen = set()
    for value in values:
        text = str(value or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        cleaned.append(text)
    return tuple(cleaned)


def _build_payload_labels(fields: tuple[str, ...] | list[str], labels: dict | None = None) -> dict[str, str]:
    merged: dict[str, str] = {}
    custom_labels = labels if isinstance(labels, dict) else {}
    for field_name in fields:
        default_label = FIELD_LABELS.get(field_name)
        custom_label = str(custom_labels.get(field_name) or "").strip()
        if custom_label:
            merged[field_name] = custom_label
        elif default_label:
            merged[field_name] = default_label
    return merged


def validate_invoice_schema_payload(payload: dict) -> tuple[dict | None, str | None]:
    if not isinstance(payload, dict):
        return None, "模板数据格式不正确"

    raw_key = str(payload.get("key") or payload.get("name") or "").strip()
    key = _normalize_schema_key(raw_key)
    if not raw_key:
        return None, "请填写模板标识"
    if not key:
        return None, "模板标识只能包含英文、数字和下划线"
    if key in BUILTIN_INVOICE_TYPE_SCHEMAS:
        return None, "不能覆盖内置模板，请更换模板标识"

    name = str(payload.get("name") or "").strip()
    if not name:
        return None, "请填写模板名称"

    fields = _clean_str_list(payload.get("fields"))
    if "invoice_type" not in fields:
        return None, "模板字段必须包含 invoice_type"

    invalid_fields = [field_name for field_name in fields if not re.fullmatch(r"[A-Za-z0-9_]+", field_name)]
    if invalid_fields:
        return None, f"字段标识不合法：{invalid_fields[0]}"

    aliases = _clean_str_list(payload.get("aliases"))
    labels = payload.get("labels") if isinstance(payload.get("labels"), dict) else {}
    cleaned_labels = {
        str(field_name).strip(): str(label).strip()
        for field_name, label in labels.items()
        if str(field_name).strip() in fields and str(label).strip()
    }

    return {
        "key": key,
        "name": name,
        "fields": list(fields),
        "aliases": list(aliases),
        "labels": cleaned_labels,
    }, None


def _schema_to_payload(
    schema: InvoiceTypeSchema,
    *,
    builtin: bool = False,
    labels: dict | None = None,
) -> dict:
    return {
        "key": schema.key,
        "name": schema.name,
        "fields": list(schema.fields),
        "aliases": list(schema.aliases),
        "labels": _build_payload_labels(schema.fields, labels),
        "builtin": builtin,
    }


def _normalize_schema_payload(payload: dict) -> dict | None:
    normalized, _ = validate_invoice_schema_payload(payload)
    return normalized


def _load_custom_schema_payloads() -> list[dict]:
    try:
        if not INVOICE_TEMPLATES_PATH.exists():
            return []
        data = json.loads(INVOICE_TEMPLATES_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []

    payloads = data.get("templates") if isinstance(data, dict) else data
    return payloads if isinstance(payloads, list) else []


def _build_custom_schemas() -> dict[str, InvoiceTypeSchema]:
    schemas: dict[str, InvoiceTypeSchema] = {}
    for payload in _load_custom_schema_payloads():
        normalized = _normalize_schema_payload(payload)
        if not normalized:
            continue
        schemas[normalized["key"]] = InvoiceTypeSchema(
            key=normalized["key"],
            name=normalized["name"],
            fields=tuple(normalized["fields"]),
            aliases=tuple(normalized["aliases"]),
        )
    return schemas


def get_field_labels() -> dict[str, str]:
    labels = dict(FIELD_LABELS)
    for payload in _load_custom_schema_payloads():
        normalized = _normalize_schema_payload(payload)
        if not normalized:
            continue
        for field_name in normalized["fields"]:
            labels.setdefault(field_name, normalized["labels"].get(field_name) or field_name)
    return labels


def get_invoice_field_order() -> tuple[str, ...]:
    order = list(INVOICE_FIELD_ORDER)
    seen = set(order)
    for payload in _load_custom_schema_payloads():
        normalized = _normalize_schema_payload(payload)
        if not normalized:
            continue
        for field_name in normalized["fields"]:
            if field_name in seen:
                continue
            seen.add(field_name)
            order.append(field_name)
    return tuple(order)


def get_invoice_schemas(include_custom: bool = True) -> dict[str, InvoiceTypeSchema]:
    schemas = dict(BUILTIN_INVOICE_TYPE_SCHEMAS)
    if include_custom:
        schemas.update(_build_custom_schemas())
    return schemas


def list_invoice_schema_payloads() -> list[dict]:
    custom_payloads = {}
    for payload in _load_custom_schema_payloads():
        normalized = _normalize_schema_payload(payload)
        if not normalized:
            continue
        custom_payloads[normalized["key"]] = normalized

    custom_keys = set(custom_payloads)
    payloads = []
    for key, schema in get_invoice_schemas().items():
        if key == "unknown":
            continue
        payloads.append(
            _schema_to_payload(
                schema,
                builtin=key not in custom_keys,
                labels=custom_payloads.get(key, {}).get("labels"),
            )
        )
    return payloads


def save_custom_invoice_schemas(payloads: list[dict]) -> list[dict]:
    custom_payloads = []
    for payload in payloads if isinstance(payloads, list) else []:
        normalized = _normalize_schema_payload(payload)
        if not normalized:
            continue
        custom_payloads.append(normalized)

    INVOICE_TEMPLATES_PATH.parent.mkdir(parents=True, exist_ok=True)
    data = {"version": 1, "templates": custom_payloads}
    with tempfile.NamedTemporaryFile(
        "w",
        encoding="utf-8",
        dir=str(INVOICE_TEMPLATES_PATH.parent),
        delete=False,
        suffix=".tmp",
    ) as tmp:
        json.dump(data, tmp, ensure_ascii=False, indent=2)
        tmp.write("\n")
        temp_name = tmp.name
    Path(temp_name).replace(INVOICE_TEMPLATES_PATH)
    return list_invoice_schema_payloads()


def upsert_custom_invoice_schema(payload: dict) -> list[dict]:
    existing = [
        item
        for item in list_invoice_schema_payloads()
        if not item.get("builtin")
    ]
    key = _normalize_schema_key((payload or {}).get("key") or (payload or {}).get("name"))
    updated = False
    next_items = []
    for item in existing:
        if item.get("key") == key:
            next_items.append(payload)
            updated = True
        else:
            next_items.append(item)
    if not updated:
        next_items.append(payload)
    return save_custom_invoice_schemas(next_items)


def delete_custom_invoice_schema(key: str) -> list[dict]:
    normalized_key = _normalize_schema_key(key)
    existing = [
        item
        for item in list_invoice_schema_payloads()
        if not item.get("builtin") and item.get("key") != normalized_key
    ]
    return save_custom_invoice_schemas(existing)


def get_invoice_schema(category: str) -> InvoiceTypeSchema:
    return get_invoice_schemas().get(str(category or "").strip()) or BUILTIN_INVOICE_TYPE_SCHEMAS["unknown"]


def infer_category_from_type_name(type_name: str) -> str:
    text = str(type_name or "").strip()
    if not text:
        return "unknown"
    schemas = get_invoice_schemas()
    ordered_items = [
        (key, schema)
        for key, schema in schemas.items()
        if key not in BUILTIN_INVOICE_TYPE_SCHEMAS and key != "unknown"
    ] + [
        (key, schema)
        for key, schema in schemas.items()
        if key in BUILTIN_INVOICE_TYPE_SCHEMAS and key != "unknown"
    ]
    for key, schema in ordered_items:
        if key == "unknown":
            continue
        if any(alias and alias in text for alias in schema.aliases):
            return key
    return "unknown"
