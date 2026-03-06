# DocFlow 测试样本说明

已生成以下文件：

- `sample.txt`：纯文本样本
- `sample.md`：Markdown 样本
- `sample.json`：JSON 样本
- `sample.csv`：CSV 表格样本
- `sample.docx`：Word 样本
- `sample.doc`：兼容样本（由 docx 改扩展名生成）
- `sample.xlsx`：Excel 样本
- `sample.xls`：兼容样本（由 xlsx 改扩展名生成）
- `sample.pptx`：PPT 样本
- `sample_text.pdf`：文本型 PDF 样本
- `sample_scan.pdf`：扫描型 PDF 样本（适合测试 OCR 兜底）
- `sample.png/.jpg/.bmp/.tiff/.webp`：图片 OCR 样本

说明：

- `sample.doc` 和 `sample.xls` 不是传统二进制老格式，而是兼容测试样本。
- 当前项目对这两类文件采用“优先按新格式内容尝试打开”的策略。