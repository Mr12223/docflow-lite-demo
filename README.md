# DocFlow

DocFlow 是一个基于 Python 的多格式文档自动化处理与内容提取系统。项目的核心目标不是单纯做“文件摘要工具”，而是将不同来源、不同格式的文档统一转换为可分析的文本、表格和结构化数据，并以**发票识别与字段提取**作为重点应用场景进行验证。

当前项目适合作为本科毕业设计系统进行展示和论文实验，不定位为成熟商用财税产品。

## 项目定位

项目主线可以概括为：

```text
多格式文件上传
    -> 文档解析 / 图片 OCR
    -> 统一文本与表格结果
    -> 发票字段结构化提取
    -> SQLite 存储、查询、导出
    -> 样本集批量评测与报告生成
```

其中，多格式文件提取、OCR、关键词和摘要等能力主要承担“文档处理底座”的作用。它们用于把 PDF、Word、Excel、PPT、图片等不同输入转换成后续可以分析的数据。项目目前最清晰的业务落地点是发票识别，而不是泛泛地做文档摘要。

## 主要功能

### 1. 多格式文档解析

系统提供统一的文档处理入口，支持以下类型文件：

- PDF：支持文本层解析、扫描件 OCR 增强和不同处理模式
- Word：支持 `docx`，并兼容部分 `doc`
- Excel：支持 `xlsx`，并兼容部分 `xls`
- PPT：支持 `pptx`
- 文本类文件：支持 `txt`、`md`、`json`、`csv`
- 图片文件：支持 `jpg`、`jpeg`、`png`、`bmp`、`tiff`、`webp`

解析结果包括正文文本、表格、基础统计信息、元数据和格式化输出。

### 2. OCR 识别

针对图片和扫描类 PDF，系统接入 OCR 能力，将图像内容转换为文本。

当前 OCR 链路包含：

- RapidOCR
- Tesseract
- 多 OCR 候选结果评分
- OCR 结果选择
- 图片预处理和缓存
- 扫描 PDF 的 OCR 回退解析

OCR 的主要价值是为后续发票字段提取、内容检索和批量评测提供文本来源。

默认图片 OCR 使用快速档：RapidOCR 只跑少量高收益预处理变体；多文件处理时会并行运行任务。当前本地演示和发票样本识别推荐使用下面这组配置，优先保证速度稳定，并避免 Tesseract 陪跑拉长耗时：

```powershell
$env:DOCFLOW_IMAGE_OCR_ORDER="rapidocr"
$env:DOCFLOW_IMAGE_OCR_SPEED_PROFILE="fast"
$env:DOCFLOW_RAPIDOCR_PREWARM="1"
$env:DOCFLOW_RAPIDOCR_TAX_ID_CROP_VARIANT_LIMIT="1"
$env:DOCFLOW_RAPIDOCR_IMAGE_OCR_VARIANT_WORKERS="1"
```

说明：

- `DOCFLOW_IMAGE_OCR_ORDER="rapidocr"`：图片 OCR 只使用 RapidOCR，减少 Tesseract 额外耗时。
- `DOCFLOW_IMAGE_OCR_SPEED_PROFILE="fast"`：只启用高收益预处理变体。
- `DOCFLOW_RAPIDOCR_PREWARM="1"`：应用启动时预热模型，降低第一次识别的等待感。
- `DOCFLOW_RAPIDOCR_TAX_ID_CROP_VARIANT_LIMIT="1"`：税号补强只跑首个裁剪变体，减少额外 OCR 次数。
- `DOCFLOW_RAPIDOCR_IMAGE_OCR_VARIANT_WORKERS="1"`：避免变体并行和多文件并行抢 CPU / 内存；如果机器性能较好，可以临时改为 `2` 做对比。

如需调试并行链路，可通过环境变量调整：

```powershell
$env:DOCFLOW_IMAGE_OCR_PARALLEL_ENGINES="1"
$env:DOCFLOW_IMAGE_OCR_ENGINE_WORKERS="2"
$env:DOCFLOW_RAPIDOCR_IMAGE_OCR_VARIANT_WORKERS="1"
```

图片发票默认启用“自适应补强”：如果主 OCR 已经识别出完整购销方税号，就跳过额外的税号裁剪 OCR；如果税号缺失、长度异常或购销方税号重复，才补跑裁剪识别。这样可以减少常见清晰发票的耗时，同时保留复杂样本的兜底能力。

如果需要更高精度、愿意接受更长耗时，可以切换到完整变体档，并打开 Tesseract 税号补强：

```powershell
$env:DOCFLOW_IMAGE_OCR_SPEED_PROFILE="accurate"
$env:DOCFLOW_IMAGE_OCR_TESSERACT_TAX_ID_REFINEMENT="1"
$env:DOCFLOW_RAPIDOCR_TAX_ID_CROP_VARIANT_LIMIT="3"
```

如果要回到最稳妥的完整补强策略，可关闭自适应跳过：

```powershell
$env:DOCFLOW_IMAGE_OCR_ADAPTIVE_INVOICE_REFINEMENT="0"
```

默认情况下，税号精修会更保守：缺买方税号只裁买方区域，缺销售方税号只裁销售方区域；买卖方税号重复时不再默认补跑裁剪 OCR。如果需要恢复重复税号也补跑，可开启：

```powershell
$env:DOCFLOW_IMAGE_OCR_DUPLICATE_TAX_ID_REFINEMENT="1"
```

如需排查兼容性问题，可临时关闭引擎并行：

```powershell
$env:DOCFLOW_IMAGE_OCR_PARALLEL_ENGINES="0"
```

### 3. 发票结构化提取

系统可从文档文本或 OCR 文本中识别发票，并提取关键字段：

- 发票代码
- 发票号码
- 开票日期
- 不含税金额
- 税额
- 价税合计
- 购买方名称
- 销售方名称
- 购买方税号
- 销售方税号
- 发票类型

发票识别是当前项目最主要的应用场景，也是论文和答辩中最适合重点展示的部分。

发票字段提取已改为可扩展的“类型识别 + 字段模板”机制：系统会先判断票据类型，再按该类型需要的字段进行结构化提取。不同类型票据不会强行套用同一套增值税发票字段；该类型本来不需要的字段会标记为“该类型不适用”，应有但未识别出的字段会保留空值并标记为“原图疑似为空/未识别”。

当前字段模板集中维护在 `invoice_schema.py`，新增票据类型时优先扩展该文件中的字段标签、字段顺序、类型别名和字段集合；如需新增字段抽取逻辑，再在 `invoice_extractor.py` 中增加对应的 `_find_xxx()` 方法。扩展流程见 `docs/invoice_schema_extension.md`。

### 4. 发票记录管理

识别出的发票信息可以保存到 SQLite 数据库，并支持：

- 发票记录分页查询
- 发票详情查看
- 单条删除
- 全部清空
- CSV 导出

这部分使系统具备从“识别结果展示”到“数据管理”的完整闭环。

### 5. 辅助内容分析

系统在完成正文提取后，会提供关键词、摘要和基础统计信息。这些功能主要用于帮助用户快速预览文档内容，属于辅助展示能力，不是本项目的核心创新点。

### 6. 批量测试与评测

项目提供样本集和脚本化评测能力，用于验证系统处理效果：

- 通用文档批量测试
- 边界样本测试
- 发票字段级准确率评测
- 图片 OCR 发票评测
- Markdown、JSON、CSV、HTML 报告输出

评测结果可用于论文实验章节和答辩说明。

## 技术栈

| 类型 | 技术 |
| --- | --- |
| 后端框架 | Flask |
| 前端 | HTML / CSS / JavaScript 单页应用 |
| 文档解析 | PyMuPDF、pdfplumber、python-docx、openpyxl、python-pptx |
| OCR | RapidOCR、Tesseract |
| 数据库 | SQLite |
| 评测与脚本 | Python 脚本、CSV/JSON/Markdown/HTML 报告 |
| 部署 | Docker、Render 配置 |

## 目录结构

```text
ShiXunClaud/
├─ app.py                         # Flask 启动入口
├─ docflow_core.py                # 多格式文档解析核心
├─ docflow_support.py             # 依赖检查、错误增强和辅助函数
├─ invoice_extractor.py           # 发票字段提取规则
├─ invoice_db.py                  # 发票 SQLite 持久化
├─ docflow/                       # 公共配置、路径、运行时和 Web 应用
│  ├─ paths.py
│  ├─ settings.py
│  ├─ runtime.py
│  └─ webapp/
│     ├─ routes/                  # Flask 路由层
│     └─ services/                # 业务服务层
├─ frontend/
│  └─ doc_tool.html               # 前端页面
├─ scripts/                       # 批测、评测、样本生成和论文材料脚本
├─ tests/                         # 单元测试与回归测试
├─ sample_data/                   # 通用文档测试样本
├─ 发票样本/                       # 发票专项样本、清单和说明
├─ docs/                          # 架构、部署、毕业设计相关文档
├─ reports/                       # 批测和评测报告输出目录
│  └─ archive/                    # 历史报告归档目录
├─ uploads_temp/                  # 上传文件、临时文件和 OCR 缓存
└─ data/                          # SQLite 数据文件
```

### 源码与生成物说明

项目中的目录大致分为三类：

- 源码与配置：`app.py`、`docflow/`、`docflow_core.py`、`docflow_support.py`、`invoice_extractor.py`、`invoice_db.py`、`frontend/`、`scripts/`、`tests/`
- 样本与文档：`sample_data/`、`发票样本/`、`docs/`
- 运行生成物：`reports/`、`uploads_temp/`、`data/`、`tmp/`、`debug.log`、`.tmp_*.json`

运行生成物主要用于本地调试、缓存、数据库和测试报告。旧报告建议放到 `reports/archive/`，日常只保留最新的批测或评测报告，便于论文和答辩引用。

## 本地启动

项目默认使用 Windows PowerShell 命令示例。

```powershell
.\.venv\Scripts\python.exe app.py
```

启动后浏览器打开：

```text
http://127.0.0.1:5000
```

如果需要重新安装依赖，可参考：

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-cloud.txt
```

注意：OCR 和旧版 Office 文件兼容能力可能依赖本机环境，例如 Tesseract、LibreOffice 等外部程序。

## 常用命令

### 1. 语法检查

```powershell
.\.venv\Scripts\python.exe -m py_compile app.py docflow_core.py docflow_support.py invoice_extractor.py invoice_db.py
```

### 2. 运行单元测试

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

### 3. 运行通用批量测试

```powershell
.\.venv\Scripts\python.exe .\scripts\run_batch_tests.py test_documents test_documents_edge_cases --strict
```

批测报告会输出到 `reports/batch_test_<时间戳>/` 目录，通常包含：

- `report.md`
- `summary.html`
- `results.json`
- `results.csv`

其中 `--strict` 表示如果样本处理结果与预期不一致，命令会返回失败状态，适合用来判断当前代码是否有回归。

### 4. 运行发票字段评测

评测 Excel 真值集：

```powershell
.\.venv\Scripts\python.exe .\scripts\run_invoice_evaluation.py .\发票样本\invoice_excel_manifest.csv
```

评测图片发票集：

```powershell
.\.venv\Scripts\python.exe .\scripts\run_invoice_evaluation.py .\发票样本\invoice_image_manifest.csv
```

评测报告会输出到 `reports/invoice_evaluation_*` 目录。

### 5. 整理历史报告

如果 `reports/` 下历史报告过多，可以把旧报告移动到归档目录，只保留最新结果：

```powershell
New-Item -ItemType Directory -Force .\reports\archive
```

然后将旧的 `batch_test_*`、`evaluation_*`、`invoice_evaluation_*` 目录移动到 `reports/archive/`。建议保留最新一份通用批测报告和最新一份发票评测报告在 `reports/` 根目录下。

## 答辩演示建议

建议演示时重点展示下面几条链路：

1. 上传普通文档，展示系统能够解析文本、表格和元数据。
2. 上传图片或扫描类文件，展示 OCR 识别结果。
3. 上传发票样本，展示发票字段自动提取。
4. 查看发票历史记录，演示查询、删除和 CSV 导出。
5. 打开发票评测报告，展示字段级准确率统计。

不建议把摘要和关键词作为核心卖点。它们更适合作为“提取结果预览”的辅助功能。

## 当前完成度

从本科毕业设计角度看，当前项目已经具备：

- 可运行的 Web 系统
- 多格式文档处理能力
- 图片 OCR 能力
- 发票结构化提取场景
- SQLite 数据持久化
- 前后端演示闭环
- 样本集与评测报告
- 论文和答辩材料基础

当前完成度可评估为约 **85%**。后续最重要的工作不是继续盲目增加功能，而是补齐最新批测报告、接口测试、README 说明和论文实验口径。

详细分析见：

- `docs/graduation/当前项目完成度分析与后续完善.md`
- `docs/项目架构说明.md`
- `发票样本/评测说明.md`

## 当前限制

项目仍存在一些限制：

- 任务状态主要保存在进程内存中，服务重启后任务状态会丢失
- 没有用户登录、权限控制和多用户数据隔离
- 部分旧版 `doc`、`xls` 文件依赖外部工具和文件本身质量
- 发票字段提取主要基于规则和启发式策略，不是训练型模型
- 前端页面和部分核心模块仍较大，后续可继续拆分
- 当前更适合毕设演示和实验验证，不适合直接作为生产系统上线

## 推荐后续完善方向

优先级较高的收尾工作：

1. 重新生成最新通用批测报告。
2. 整理发票评测结果，明确论文中的实验口径。
3. 补充 Flask 接口级测试。
4. 将测试、评测和演示流程写入论文与答辩材料。
5. 逐步拆分 `docflow_core.py`、`invoice_extractor.py` 和前端页面。

## 适合论文中的项目表述

可以将本项目概括为：

> 本系统以 Python 为主要开发语言，基于 Flask 构建 Web 服务，围绕多格式文档自动化处理场景，实现了 PDF、Word、Excel、PPT、文本文件和图片文件的统一解析流程；同时结合 OCR 技术和规则化信息抽取方法，实现了发票关键信息的结构化提取、存储、查询和导出，并通过样本集和评测脚本对系统效果进行了验证。

这个表述比“文档摘要工具”更准确，也更能体现项目的实际工作量。
