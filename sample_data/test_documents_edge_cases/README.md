# DocFlow 异常 / 边界测试样本

## 文件与预期

| 文件 | 说明 | 预期表现 |
| --- | --- | --- |
| `00_empty.txt` | 空文本文件 | 成功，提取 0 字符 |
| `01_utf8_bom.txt` | UTF-8 BOM 编码 | 成功，中文正常 |
| `02_gbk.txt` | GBK 编码文本 | 成功，自动识别编码 |
| `03_utf16.txt` | UTF-16 文本 | 成功，自动识别编码 |
| `04_long_line.txt` | 超长单行文本 | 成功，测试长文本稳定性 |
| `05_special_chars.md` | 特殊字符 Markdown | 成功，保留内容 |
| `06_带空格 和 中文 @2026!.txt` | 特殊文件名 | 成功，验证文件名兼容 |
| `10_nested.json` | 嵌套 JSON | 成功，格式化输出 |
| `11_invalid.json` | 非法 JSON | 失败，返回解析错误 |
| `12_empty.csv` | 空 CSV | 成功，行列为 0 |
| `13_irregular.csv` | 列数不一致 CSV | 成功，按原样读取 |
| `14_large_table.csv` | 500 行 CSV | 成功，测试大表格 |
| `20_empty.docx` | 空 Word 文档 | 成功或接近空内容 |
| `21_minimal.docx` | 最小 Word 文档 | 成功 |
| `22_compat.doc` | 兼容 .doc 样本 | 成功 |
| `23_fake.docx` | 伪造 docx | 失败，返回错误 |
| `24_fake.doc` | 伪造 doc | 失败，返回错误 |
| `30_empty.xlsx` | 空 Excel | 成功 |
| `31_minimal.xlsx` | 最小 Excel | 成功 |
| `32_compat.xls` | 兼容 .xls 样本 | 成功 |
| `33_fake.xlsx` | 伪造 xlsx | 失败，返回错误 |
| `34_fake.xls` | 伪造 xls | 失败，返回错误 |
| `40_empty.pptx` | 空白 PPT | 成功或近空内容 |
| `41_minimal.pptx` | 最小 PPT | 成功 |
| `42_fake.pptx` | 伪造 pptx | 失败，返回错误 |
| `50_blank.png` | 空白图片 | OCR 结果为空或极少 |
| `51_low_contrast.jpg` | 低对比度图片 | OCR 可能不稳定 |
| `52_tiny_text.png` | 超小字体图片 | OCR 可能部分缺失 |
| `53_dense_text.webp` | 密集文本图片 | OCR 可提取部分内容 |
| `54_tiff_scan.tiff` | TIFF 样本 | OCR 成功 |
| `60_blank_scan.pdf` | 空白扫描 PDF | 失败，提示未识别到内容 |
| `61_corrupt.pdf` | 损坏 PDF | 失败，返回错误 |
| `62_short_text.pdf` | 极短文本 PDF | 成功 |
| `70_unsupported.bin` | 不支持扩展名 | 失败，提示格式不支持 |
| `71_empty.json` | 空 JSON 文件 | 失败，返回解析错误 |

## 说明

- `22_compat.doc` 与 `32_compat.xls` 是兼容测试样本，不是真正旧版二进制格式。
- 这些样本适合用于演示“系统在异常输入下的鲁棒性与错误提示”。
