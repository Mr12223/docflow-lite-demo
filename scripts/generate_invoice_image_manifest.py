from __future__ import annotations

import argparse
import csv
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

from _bootstrap import ensure_project_root_on_path
import fitz

ensure_project_root_on_path()

from docflow.paths import PROJECT_ROOT


DEFAULT_SOURCE = PROJECT_ROOT / "发票样本" / "invoice_excel_manifest.csv"
DEFAULT_PDF_DIR = PROJECT_ROOT / "发票样本" / "invoices" / "pdf"
DEFAULT_IMAGE_DIR = PROJECT_ROOT / "发票样本" / "invoices" / "images"
DEFAULT_OUTPUT = PROJECT_ROOT / "发票样本" / "invoice_image_manifest.csv"
HEADERS = [
    "sample_id",
    "file_name",
    "source_dir",
    "sample_type",
    "expected_invoice",
    "invoice_code",
    "invoice_number",
    "invoice_date",
    "amount",
    "tax",
    "total",
    "buyer_name",
    "seller_name",
    "buyer_tax_id",
    "seller_tax_id",
    "invoice_type",
    "quality_level",
    "notes",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="generate_invoice_image_manifest",
        description="从 PDF 发票渲染图片并基于 PDF 文本生成图片评测真值清单。",
    )
    parser.add_argument(
        "--source",
        default=str(DEFAULT_SOURCE),
        help="源清单路径，默认 `发票样本/invoice_excel_manifest.csv`。",
    )
    parser.add_argument(
        "--pdf-dir",
        default=str(DEFAULT_PDF_DIR),
        help="PDF 发票目录，默认 `发票样本/invoices/pdf`。",
    )
    parser.add_argument(
        "--image-dir",
        default=str(DEFAULT_IMAGE_DIR),
        help="图片输出目录，默认 `发票样本/invoices/images`。",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT),
        help="输出图片清单路径，默认 `发票样本/invoice_image_manifest.csv`。",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=220,
        help="PDF 首页渲染 DPI，默认 220。",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="重新渲染已存在的 PNG 图片。",
    )
    return parser.parse_args()


def _resolve_path(raw_path: str) -> Path:
    path = Path(raw_path)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path


def _relative_to_base(path: Path, base_dir: Path) -> str:
    try:
        return str(path.relative_to(base_dir)).replace("\\", "/")
    except ValueError:
        try:
            return str(path.relative_to(PROJECT_ROOT)).replace("\\", "/")
        except ValueError:
            return str(path).replace("\\", "/")


def _convert_sample_type(sample_type: str) -> str:
    sample_type = str(sample_type or "").strip()
    if sample_type.startswith("excel_"):
        return "image_" + sample_type[len("excel_") :]
    if sample_type:
        return "image_" + sample_type
    return "image_other"


def _render_first_page(pdf_path: Path, image_path: Path, dpi: int, overwrite: bool) -> bool:
    if image_path.exists() and not overwrite:
        return False

    image_path.parent.mkdir(parents=True, exist_ok=True)
    scale = max(dpi, 72) / 72
    with fitz.open(pdf_path) as document:
        if document.page_count < 1:
            raise ValueError(f"PDF has no pages: {pdf_path}")
        page = document.load_page(0)
        pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        pixmap.save(image_path)
    return True


def _read_pdf_text(pdf_path: Path) -> str:
    with fitz.open(pdf_path) as document:
        return "\n".join(page.get_text("text") for page in document)


def _normalize_text(value: str) -> str:
    return " ".join(str(value or "").strip().split())


def _normalize_date(value: str) -> str:
    text = _normalize_text(value)
    match = re.search(r"((?:19|20)\d{2})[年/\-.](\d{1,2})[月/\-.](\d{1,2})日?", text)
    if not match:
        return text.replace("/", "-").replace(".", "-")
    year, month, day = match.groups()
    return f"{year}-{int(month):02d}-{int(day):02d}"


def _normalize_amount(value: str | Decimal | None) -> str:
    if value in (None, ""):
        return ""

    text = str(value).replace("¥", "").replace("￥", "").replace(",", "").strip()
    try:
        amount = Decimal(text)
    except InvalidOperation:
        return _normalize_text(value)

    amount = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return format(amount, "f")


def _search_value(text: str, pattern: str) -> str:
    match = re.search(pattern, text, re.S)
    if not match:
        return ""
    return _normalize_text(match.group(1))


def _split_lines(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line and line.strip()]


def _compact(line: str) -> str:
    return re.sub(r"\s+", "", line or "")


def _line_after(lines: list[str], label: str) -> str:
    compact_label = _compact(label)
    for index, line in enumerate(lines):
        if _compact(line) != compact_label:
            continue
        for next_line in lines[index + 1 :]:
            if next_line.strip():
                return _normalize_text(next_line)
    return ""


def _sum_grouped_amounts(
    text: str,
    start_marker: str,
    end_marker: str,
    group_size: int,
    amount_index: int,
    tax_index: int | None = None,
) -> tuple[str, str]:
    match = re.search(re.escape(start_marker) + r"(.*?)" + re.escape(end_marker), text, re.S)
    if not match:
        return "", ""

    body = match.group(1)
    values = [
        Decimal(raw.replace(",", ""))
        for raw in re.findall(r"(?<!\d)(\d[\d,]*\.\d{2})(?!\d)", body)
    ]
    if len(values) < group_size:
        return "", ""

    usable_count = len(values) - (len(values) % group_size)
    if usable_count <= 0:
        return "", ""

    values = values[:usable_count]
    amount_total = sum(
        value for index, value in enumerate(values) if index % group_size == amount_index
    )
    tax_total = (
        sum(value for index, value in enumerate(values) if index % group_size == tax_index)
        if tax_index is not None
        else Decimal("0.00")
    )
    return _normalize_amount(amount_total), _normalize_amount(tax_total) if tax_index is not None else ""


def _extract_total(text: str, label: str) -> str:
    pattern = re.escape(label) + r".*?[¥￥]\s*([0-9][0-9,]*\.\d{2})"
    return _normalize_amount(_search_value(text, pattern))


def _derive_total(amount: str, tax: str) -> str:
    if not amount:
        return ""
    amount_value = Decimal(amount)
    tax_value = Decimal(tax or "0.00")
    return _normalize_amount(amount_value + tax_value)


def _parse_electronic(text: str) -> dict[str, str]:
    amount, tax = _sum_grouped_amounts(text, "税额", "价税合计", group_size=4, amount_index=1, tax_index=3)
    total = _extract_total(text, "价税合计")
    if total and amount and not tax:
        tax = _normalize_amount(Decimal(total) - Decimal(amount))

    return {
        "invoice_code": "",
        "invoice_number": _search_value(text, r"发票号码[:：]\s*([A-Z]?\d{8,20})"),
        "invoice_date": _normalize_date(_search_value(text, r"开票日期[:：]\s*([0-9]{4}[年/\-.]\d{1,2}[月/\-.]\d{1,2}日?)")),
        "amount": amount,
        "tax": tax,
        "total": total or _derive_total(amount, tax),
        "buyer_name": _search_value(text, r"购买方[:：]\s*([^\r\n]+)"),
        "seller_name": _search_value(text, r"销售方[:：]\s*([^\r\n]+)"),
        "buyer_tax_id": "",
        "seller_tax_id": "",
        "invoice_type": "电子发票",
    }


def _parse_normal(text: str) -> dict[str, str]:
    total = _extract_total(text, "合计")
    return {
        "invoice_code": "",
        "invoice_number": _search_value(text, r"发票号码[:：]\s*([A-Z]?\d{8,20})"),
        "invoice_date": _normalize_date(_search_value(text, r"开票日期[:：]\s*([0-9]{4}[年/\-.]\d{1,2}[月/\-.]\d{1,2}日?)")),
        "amount": total,
        "tax": "",
        "total": total,
        "buyer_name": _search_value(text, r"(?:购方名称|购买方|购方)[:：]\s*([^\r\n]+)"),
        "seller_name": _search_value(text, r"(?:销方名称|销售方|销方)[:：]\s*([^\r\n]+)"),
        "buyer_tax_id": "",
        "seller_tax_id": "",
        "invoice_type": "增值税普通发票",
    }


def _parse_vat(text: str) -> dict[str, str]:
    buyer_match = re.search(
        r"购买方信息\s+名称[:：]\s*(.+?)\s+纳税人识别号[:：]\s*([A-Z0-9]{15,20})",
        text,
        re.S,
    )
    seller_match = re.search(
        r"销售方信息\s+名称[:：]\s*(.+?)\s+纳税人识别号[:：]\s*([A-Z0-9]{15,20})",
        text,
        re.S,
    )
    amount, tax = _sum_grouped_amounts(text, "税额", "价税合计", group_size=4, amount_index=1, tax_index=3)
    total = _extract_total(text, "价税合计")
    if total and amount and not tax:
        tax = _normalize_amount(Decimal(total) - Decimal(amount))

    return {
        "invoice_code": _search_value(text, r"发票代码[:：]\s*(\d{10,12})"),
        "invoice_number": _search_value(text, r"发票号码[:：]\s*([A-Z]?\d{8,20})"),
        "invoice_date": _normalize_date(_search_value(text, r"开票日期[:：]\s*([0-9]{4}[年/\-.]\d{1,2}[月/\-.]\d{1,2}日?)")),
        "amount": amount,
        "tax": tax,
        "total": total or _derive_total(amount, tax),
        "buyer_name": _normalize_text(buyer_match.group(1)) if buyer_match else "",
        "seller_name": _normalize_text(seller_match.group(1)) if seller_match else "",
        "buyer_tax_id": _normalize_text(buyer_match.group(2)).upper() if buyer_match else "",
        "seller_tax_id": _normalize_text(seller_match.group(2)).upper() if seller_match else "",
        "invoice_type": "增值税专用发票",
    }


def _parse_receipt(text: str) -> dict[str, str]:
    lines = _split_lines(text)
    amount = _normalize_amount(_line_after(lines, "金额"))
    return {
        "invoice_code": "",
        "invoice_number": _line_after(lines, "收据编号"),
        "invoice_date": _normalize_date(_line_after(lines, "日期")),
        "amount": amount,
        "tax": "",
        "total": amount,
        "buyer_name": _line_after(lines, "付款单位"),
        "seller_name": _line_after(lines, "收款单位"),
        "buyer_tax_id": "",
        "seller_tax_id": "",
        "invoice_type": "收据",
    }


def _extract_pdf_fields(pdf_path: Path) -> dict[str, str]:
    text = _read_pdf_text(pdf_path)
    stem = pdf_path.stem

    if stem.startswith("invoice_electronic_"):
        return _parse_electronic(text)
    if stem.startswith("invoice_normal_"):
        return _parse_normal(text)
    if stem.startswith("invoice_vat_"):
        return _parse_vat(text)
    if stem.startswith("invoice_receipt_"):
        return _parse_receipt(text)

    raise ValueError(f"Unsupported invoice pdf: {pdf_path.name}")


def main() -> int:
    args = parse_args()
    source_path = _resolve_path(args.source)
    pdf_dir = _resolve_path(args.pdf_dir)
    image_dir = _resolve_path(args.image_dir)
    output_path = _resolve_path(args.output)
    manifest_base_dir = output_path.parent

    if not source_path.exists():
        print(f"未找到源清单：{source_path}")
        return 1
    if not pdf_dir.exists():
        print(f"未找到 PDF 目录：{pdf_dir}")
        return 1

    with source_path.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = [dict(row) for row in csv.DictReader(handle)]

    if not rows:
        print(f"源清单为空：{source_path}")
        return 1

    output_rows: list[dict[str, str]] = []
    rendered = 0
    reused = 0
    relative_image_dir = _relative_to_base(image_dir, manifest_base_dir)

    for index, row in enumerate(rows, start=1):
        file_name = str(row.get("file_name", "")).strip()
        if not file_name:
            raise ValueError(f"第 {index} 行缺少 file_name")

        pdf_name = Path(file_name).with_suffix(".pdf").name
        image_name = Path(pdf_name).with_suffix(".png").name
        pdf_path = pdf_dir / pdf_name
        image_path = image_dir / image_name

        if not pdf_path.exists():
            raise FileNotFoundError(f"未找到 PDF 文件：{pdf_path}")

        did_render = _render_first_page(
            pdf_path,
            image_path,
            dpi=args.dpi,
            overwrite=args.overwrite,
        )
        if did_render:
            rendered += 1
        else:
            reused += 1

        parsed_fields = _extract_pdf_fields(pdf_path)
        output_row = {
            "sample_id": f"IMG{index:03d}",
            "file_name": image_name,
            "source_dir": relative_image_dir,
            "sample_type": _convert_sample_type(row.get("sample_type", "")),
            "expected_invoice": str(row.get("expected_invoice") or "true"),
            "invoice_code": parsed_fields["invoice_code"],
            "invoice_number": parsed_fields["invoice_number"],
            "invoice_date": parsed_fields["invoice_date"],
            "amount": parsed_fields["amount"],
            "tax": parsed_fields["tax"],
            "total": parsed_fields["total"],
            "buyer_name": parsed_fields["buyer_name"],
            "seller_name": parsed_fields["seller_name"],
            "buyer_tax_id": parsed_fields["buyer_tax_id"],
            "seller_tax_id": parsed_fields["seller_tax_id"],
            "invoice_type": parsed_fields["invoice_type"],
            "quality_level": str(row.get("quality_level") or "high"),
            "notes": "PDF文本真值自动抽取；图片由PDF首页渲染生成",
        }
        output_rows.append({header: str(output_row.get(header, "") or "") for header in HEADERS})

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS)
        writer.writeheader()
        writer.writerows(output_rows)

    print(f"已生成图片评测清单：{output_path}")
    print(f"- 样本数：{len(output_rows)}")
    print(f"- 新渲染图片：{rendered}")
    print(f"- 复用已有图片：{reused}")
    print(f"- 图片目录：{image_dir}")
    print(f"- 清单 source_dir：{relative_image_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
