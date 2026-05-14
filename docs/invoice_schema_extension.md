# 发票类型扩展说明

DocFlow 的发票结构化提取采用“先识别类型，再按类型字段模板提取”的方式。内置类型、字段标签和字段顺序集中维护在 `invoice_schema.py`；用户在网页端新增的自定义模板会保存到 `data/invoice_templates.json`。

## 网页端维护模板

在前端“识别设置 -> 发票模板”面板中可以新增、编辑或删除自定义模板。模板包含：

- 模板标识：英文、数字、下划线，例如 `hotel_receipt`。
- 模板名称：展示给用户看的名称，例如 `酒店收据`。
- 识别关键词/别名：用于根据 OCR 文本识别类型，例如 `酒店收据,住宿费`。
- 字段列表：该类型需要展示和校验的字段，必须包含 `invoice_type`。
- 字段中文名：可选，每行 `field_key=中文名`。

自定义模板会影响：

- `expected_fields`：前端只展示该类型需要的字段。
- `missing_fields`：该类型应有但未识别出的字段。
- `not_applicable_fields`：其他类型字段标记为“不适用”。

注意：新增模板可以控制“展示哪些字段”，但不会自动创造复杂字段抽取能力。如果新增了 `hotel_name` 这类全新字段，通常还需要在 `invoice_extractor.py` 中增加对应 `_find_hotel_name()` 规则。

## 新增一种票据类型

如果需要作为内置模板提交到代码中，可按下面步骤：

1. 在 `invoice_schema.py` 中新增字段标签。

   ```python
   FIELD_LABELS["new_field"] = "新字段"
   ```

2. 将新字段加入 `INVOICE_FIELD_ORDER`，用于前端和导出的稳定展示顺序。

3. 在 `INVOICE_TYPE_SCHEMAS` 中注册新类型。

   ```python
   "new_ticket": InvoiceTypeSchema(
       key="new_ticket",
       name="新票据",
       aliases=("新票据", "关键标题"),
       fields=("invoice_type", "invoice_number", "invoice_date", "new_field"),
   )
   ```

4. 如需专门抽取新字段，在 `invoice_extractor.py` 中增加 `_find_xxx()`，并把字段加入 `_apply_fallbacks()` 的 `ordered_extractors`。

5. 如需更强的类型判定，在 `InvoiceExtractor._detect_invoice_category()` 中增加特征判断。简单标题类票据通常只需要配置 `aliases`，`infer_category_from_type_name()` 会自动识别。

6. 增加测试样例，重点验证：

   - 能识别为新类型；
   - `expected_fields` 只包含该类型需要的字段；
   - 不适用字段进入 `not_applicable_fields`；
   - 缺失但应存在的字段进入 `missing_fields`。

## 数据流

```text
OCR 文本
  -> InvoiceExtractor._detect_invoice_category()
  -> invoice_schema.get_invoice_schema()
  -> 只运行该 schema 需要的字段提取
  -> invoice_merge._attach_invoice_missing_fields()
  -> 前端按 expected_fields 动态展示
```

前端 `frontend/doc_tool.html` 不需要为每种票据单独写展示逻辑。只要后端返回 `expected_fields`、`missing_fields` 和 `not_applicable_fields`，页面和批量导出会自动跟随。
