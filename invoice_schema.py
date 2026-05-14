"""兼容垫片：实际代码已移至 docflow/invoice/schema.py"""
# ruff: noqa: F401, F403
from docflow.invoice.schema import *
from docflow.invoice.schema import (
    InvoiceTypeSchema,
    FIELD_LABELS,
    INVOICE_FIELD_ORDER,
    BUILTIN_INVOICE_TYPE_SCHEMAS,
    validate_invoice_schema_payload,
    get_field_labels,
    get_invoice_field_order,
    get_invoice_schemas,
    list_invoice_schema_payloads,
    save_custom_invoice_schemas,
    upsert_custom_invoice_schema,
    delete_custom_invoice_schema,
    get_invoice_schema,
    infer_category_from_type_name,
)
