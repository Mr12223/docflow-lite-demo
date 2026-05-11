# 基于Python的多格式文档自动化处理与发票信息提取系统设计与实现

> 初稿说明：本文档根据当前 DocFlow 项目代码、评测报告和软件学院 2026 届本科毕业设计（论文）模板生成。学生姓名、学号、指导教师、班级、部分图表截图和最终实验数据需要在提交前补齐。本文中的实验数据优先引用仓库现有报告，后续如重新运行评测，应同步更新第 5 章。

# 摘要

随着办公数字化、电子票据和档案管理系统的普及，企事业单位在日常业务中积累了大量 PDF、Word、Excel、PPT、图片和扫描件等异构文档。传统人工录入方式存在效率低、重复劳动多、字段容易漏填和结果格式不统一等问题。针对上述问题，本文设计并实现了一套基于 Python 的多格式文档自动化处理与发票信息提取系统 DocFlow。系统面向常见办公文档和发票处理场景，提供统一上传、自动解析、OCR 识别、结构化字段抽取、结果存储、历史查询和 CSV 导出等功能。

系统采用“前端交互层、Web 接口层、业务服务层、文档解析核心层和数据持久化层”的分层架构。前端基于原生 HTML、CSS 和 JavaScript 实现文件上传、任务轮询、结果展示、批量测试和发票记录管理；后端基于 Flask 提供同步处理、异步任务、依赖检测、批量评测和发票管理接口；核心处理层通过解析器抽象统一支持 PDF、Word、Excel、PPT、TXT、Markdown、JSON、CSV 与图片等多种输入格式。针对图片和扫描件，系统集成 RapidOCR、Tesseract 等 OCR 能力，并设计了面向发票场景的候选结果评分、字段级合并和税号修复策略。针对发票文本，系统通过规则匹配、字段清洗和置信度评估提取发票代码、发票号码、开票日期、金额、税额、价税合计、购销方名称及税号等关键字段，并将结果保存到 SQLite 数据库。

实验部分采用“样本、清单、报告”的数据驱动方式进行验证。通用批量测试报告覆盖 51 个常规与边界样本，其中 35 个带预期样本全部符合预期；发票图片字段评测集包含 30 张由 PDF 首页渲染生成的发票图片，共 249 个已标注字段，系统在该评测集上取得 100.00% 的发票判定匹配率、100.00% 的字段覆盖率和 100.00% 的字段准确率，平均处理耗时为 15050.54 ms。实验结果表明，DocFlow 能够完成从多格式文档输入、文本获取、发票字段抽取到结构化结果管理的完整处理流程，具备较好的工程完整性和毕业设计演示价值。

**关键词**：文档处理；OCR；发票识别；结构化信息提取；Flask；SQLite

# ABSTRACT

With the rapid development of digital office systems, electronic invoices and archive management platforms, enterprises and institutions have accumulated a large number of heterogeneous documents, such as PDF files, Word documents, Excel spreadsheets, PowerPoint files, images and scanned documents. Manual information entry is inefficient and error-prone, and it is difficult to maintain consistent structured results. To address these problems, this thesis designs and implements DocFlow, a Python-based multi-format document processing and invoice information extraction system. The system supports unified file uploading, automatic parsing, OCR recognition, structured field extraction, record storage, historical query and CSV export.

The system adopts a layered architecture consisting of the front-end interaction layer, Web API layer, business service layer, document processing core layer and data persistence layer. The front end is implemented with native HTML, CSS and JavaScript, providing file uploading, task polling, result visualization, batch testing and invoice record management. The back end is built with Flask and provides synchronous processing, asynchronous task management, dependency checking, batch evaluation and invoice management APIs. The core processing layer uses parser abstraction to support PDF, Word, Excel, PowerPoint, TXT, Markdown, JSON, CSV and image files in a unified way. For images and scanned documents, the system integrates OCR engines such as RapidOCR and Tesseract, and designs invoice-oriented candidate scoring, field-level merging and tax identification number repairing strategies. For invoice texts, key fields including invoice code, invoice number, invoice date, amount, tax, total, buyer and seller information are extracted through rule matching, field normalization and confidence evaluation, and the results are persisted in a SQLite database.

The system is evaluated in a data-driven manner based on samples, manifests and generated reports. The general batch test covers 51 normal and edge-case samples, and all 35 samples with expected results match their expectations. The invoice image evaluation set contains 30 invoice images rendered from PDF first pages and 249 annotated fields. On this dataset, the system achieves 100.00% invoice classification accuracy, 100.00% field coverage and 100.00% field accuracy, with an average processing time of 15050.54 ms. The results show that DocFlow can complete an end-to-end workflow from multi-format document input and text extraction to invoice field extraction and structured result management, demonstrating practical engineering value for an undergraduate software engineering project.

**KEY WORDS**: Document Processing; OCR; Invoice Recognition; Structured Information Extraction; Flask; SQLite

# 1 绪论

## 1.1 研究背景与意义

在企业财务、行政办公、档案管理和报销审核等业务场景中，文档是信息流转的重要载体。传统办公文档既包括 Word、Excel、PPT 等结构化程度较高的电子文件，也包括 PDF、扫描件、图片、票据截图等结构化程度较低的文件。随着电子发票、在线报销和数字档案系统的普及，单位内部需要处理的文档数量持续增加，人工逐份打开、复制、录入和核对的方式已经难以满足效率和准确性要求。特别是在发票处理场景中，工作人员往往需要从票据中提取发票号码、日期、金额、税额、价税合计、购销方名称及税号等字段，如果完全依赖人工录入，不仅耗时较长，而且容易因为图片质量、排版差异和视觉疲劳产生错录与漏录。

从技术发展角度看，文档自动化处理已经从简单的文件读取逐步发展为集格式解析、OCR 识别、版面理解、结构化抽取和结果管理于一体的综合任务。对于带有可检索文本层的 PDF 或 Office 文档，系统可以优先通过解析库直接读取内容，以获得更高的效率和字符保真度；对于扫描件和图片类文档，则需要借助 OCR 技术完成文字识别；对于发票、合同和报销单等业务文档，还需要进一步将非结构化文本转换为可查询、可统计、可导出的结构化字段。因此，一个实用的文档处理系统不能只解决“识别文字”这一单点问题，还需要围绕输入格式、处理流程、异常情况、字段规则和结果管理形成完整闭环。

当前常见的商业 OCR 平台虽然具备较强的识别能力，但在毕业设计和中小型私有化场景中仍存在一定限制：一是调用费用和网络依赖会增加部署成本；二是业务数据上传到第三方平台可能带来隐私和合规顾虑；三是商业接口通常更偏向通用能力，针对本地项目的流程、字段和数据库管理仍需要二次开发。与此同时，单一开源 OCR 或解析库又难以覆盖所有真实文档场景，例如 PDF 文本层质量不稳定、扫描件分辨率差异较大、旧版 Office 文件需要外部工具辅助等。因此，本文选择以轻量化、可本地运行、可扩展和可演示为目标，设计并实现 DocFlow 多格式文档自动化处理与发票信息提取系统。

本课题的意义主要体现在三个方面。首先，从应用价值看，系统能够将多格式文件统一接入并输出结构化结果，能够降低人工整理文档和录入发票字段的工作量。其次，从工程实践看，系统覆盖了需求分析、架构设计、接口开发、OCR 调度、字段抽取、数据库持久化、前端展示、测试评测和 Docker 部署等软件工程环节，具有较完整的毕业设计工作量。最后，从扩展价值看，发票识别只是系统在票据场景下的一个落地模块，底层多格式文档处理框架后续还可以扩展到合同审核、报销单识别、档案材料提取和文档问答等方向。

## 1.2 国内外研究现状

OCR 技术是扫描文档和图片文字识别的基础。传统 OCR 系统通常包括图像预处理、文本检测、字符识别和识别后处理等环节，其中 Tesseract 作为典型开源 OCR 工具，长期被用于印刷体识别和离线部署场景。随着深度学习发展，基于卷积神经网络和循环神经网络的文本检测与识别方法显著提升了复杂背景、弯曲文字和多语言场景下的识别效果。近年来，PaddleOCR、RapidOCR、EasyOCR 等开源工具进一步降低了 OCR 工程落地门槛，使开发者能够在普通服务器或个人电脑上完成中文票据、文档截图和扫描件识别。

在文档理解研究方面，单纯的字符识别已经无法满足复杂业务需求。LayoutLM、LayoutLMv2 和 LayoutLMv3 等模型将文本内容、页面布局和视觉特征联合建模，推动文档理解从“读出文字”向“理解版面和字段关系”演进。Donut、Nougat 等端到端文档理解方法尝试直接从页面图像生成结构化结果，体现了多模态模型在复杂文档场景中的潜力。但是，这类模型通常需要较高算力、训练数据和部署成本，对于本科毕业设计和轻量私有化系统而言，直接引入大模型并不是最稳妥的主体路线。本文采用“解析库优先、OCR 兜底、规则抽取和字段融合”的工程方案，在可控复杂度内实现较完整的系统能力。

在多格式文档解析方面，国内外工程实践通常优先使用文件格式对应的成熟解析库。PDF 文档可以通过 pdfplumber、PyMuPDF 和 pypdfium2 获取文本层、页面对象或渲染图像；DOCX、XLSX 和 PPTX 等 Office Open XML 文件可以分别使用 python-docx、openpyxl 和 python-pptx 读取段落、表格、工作表和幻灯片内容；TXT、Markdown、JSON 和 CSV 等文本类文件则需要重点处理编码、非法格式、空文件和字段分隔问题。由于不同格式的内部结构和错误模式差异明显，系统设计时需要通过解析器抽象降低各类文件处理逻辑之间的耦合。

在票据识别和发票结构化抽取方面，已有研究和商业系统通常围绕票面文字识别、关键信息定位、字段分类和结果校验展开。由于我国发票版式相对规范，发票代码、号码、税号、日期、金额和价税合计等字段具有较明显的格式特征，因此规则匹配、正则表达式、上下文关键词定位和字段清洗仍然具有较高工程价值。对于毕业设计项目而言，相比训练专用深度学习模型，基于 OCR 文本和业务规则构建可解释、可调试、可评测的字段抽取模块，更能体现系统分析和工程实现能力。

## 1.3 本文主要研究内容

本文围绕 DocFlow 项目的设计与实现展开，主要完成以下工作。

第一，设计并实现多格式文档统一处理框架。系统面向 PDF、Word、Excel、PPT、TXT、Markdown、JSON、CSV 与图片等输入格式，通过统一处理入口和解析器调度机制屏蔽底层文件差异，并将解析结果转换为统一的文本、表格、统计信息和格式化输出结构。

第二，设计并实现 OCR 识别和多引擎回退机制。针对图片和扫描型文档，系统集成 RapidOCR、Tesseract 等 OCR 能力，并在服务层提供统一的 OCR 门面。系统能够根据配置顺序尝试不同 OCR 引擎，对候选结果进行发票字段评分，并选择更适合发票抽取的识别结果。

第三，设计并实现发票结构化信息提取模块。系统基于规则匹配、上下文分析、字段清洗和置信度评估方法，从 OCR 文本或文档解析文本中抽取发票代码、发票号码、日期、金额、税额、合计、购销方名称、购销方税号、发票类型、机器编号和校验码等字段。

第四，设计并实现发票记录管理和 Web 展示功能。系统基于 Flask 提供文件处理、任务轮询、批量评测、发票记录查询、详情查看、删除和 CSV 导出接口，基于 SQLite 保存结构化发票记录，前端页面能够展示处理进度、识别结果和历史记录。

第五，构建数据驱动的测试与评测体系。系统提供通用批量测试、清单式评测和发票专项评测脚本，能够生成 Markdown、HTML、JSON 和 CSV 报告，为论文实验章节提供可复现的数据来源。

## 1.4 论文组织结构

全文共分为六个部分。第 1 章为绪论，介绍课题背景、研究意义、国内外研究现状和本文主要工作。第 2 章介绍系统涉及的相关技术，包括 OCR、PDF 解析、Office 文档解析、Flask Web 框架、SQLite 数据库和 Docker 部署。第 3 章进行系统需求分析与总体设计，说明功能需求、非功能需求、系统架构、模块划分、数据库设计和关键流程。第 4 章阐述系统实现，重点说明多格式文档处理引擎、OCR 多引擎调度、发票字段抽取、Web 接口、前端展示和持久化模块的实现。第 5 章进行系统测试与结果分析，给出测试环境、测试用例、批量测试、发票专项评测和问题分析。最后为结束语、致谢、参考文献和附录，总结本文工作并展望后续改进方向。

# 2 相关技术介绍

## 2.1 OCR 技术概述

OCR 即光学字符识别，其目标是将图像中的文字区域转换为可编辑、可检索和可结构化处理的文本。典型 OCR 流程包括图像预处理、文本检测、文本识别和后处理四个环节。图像预处理通常包括灰度化、二值化、去噪、旋转校正和缩放增强，用于提升后续检测和识别效果。文本检测负责从图像中定位文字区域，文本识别负责将检测到的文字图像转换为字符序列，后处理则结合语言规则、业务词典和字段格式对识别结果进行修正。

在本文系统中，OCR 不是独立功能，而是文档处理链路中的补充能力。当文档本身含有高质量文本层时，系统优先使用解析库直接读取文本；当输入为图片、扫描件或文本层质量不足的 PDF 时，系统才调用 OCR。该策略能够避免对所有文件都进行高成本图像识别，从而兼顾处理速度与识别效果。

| OCR 引擎 | 技术路线 | 离线支持 | 中文支持 | 本文使用方式 |
| --- | --- | --- | --- | --- |
| Tesseract | 传统 OCR 与 LSTM 识别结合 | 支持 | 一般，依赖语言包 | 作为本地兜底 OCR 引擎 |
| RapidOCR | 检测识别模型与 ONNX 推理 | 支持 | 较好 | 作为轻量云部署默认 OCR 引擎 |
| PaddleOCR | 深度学习检测识别框架 | 支持 | 较好 | 代码预留适配，可按环境启用 |
| EasyOCR | PyTorch OCR 工具链 | 支持 | 较好 | 代码预留适配，用于本地增强 |

## 2.2 PDF 与 Office 文档解析技术

PDF 文档可以分为文本型 PDF 和扫描型 PDF。文本型 PDF 内部包含可检索文本层，适合使用 pdfplumber 或 PyMuPDF 直接提取文本与表格信息；扫描型 PDF 实质上是图片页面的封装，无法直接读取文字，需要先将页面渲染为图像，再交由 OCR 引擎识别。DocFlow 在 PDF 处理模块中设计了 accurate、balanced 和 fast 三种模式，以便在精度、速度和资源消耗之间进行权衡。

Office 文档方面，DOCX、XLSX 和 PPTX 均属于 Office Open XML 体系，内部结构相对规范。python-docx 能读取 Word 段落和表格，openpyxl 能读取 Excel 工作表单元格，python-pptx 能读取 PPT 幻灯片文本框和表格内容。对于旧版 .doc 和 .xls 文件，系统需要借助 LibreOffice 等外部工具进行转换或降级处理，因此在依赖检查模块中提供了相关诊断和错误提示。

| 文件类型 | 主要解析工具 | 处理内容 | 主要风险 |
| --- | --- | --- | --- |
| PDF | pdfplumber、PyMuPDF、pypdfium2 | 文本层、表格、页面渲染 | 扫描件无文本层、字体映射异常 |
| DOCX | python-docx | 段落、表格、元数据 | 嵌入对象和复杂排版覆盖有限 |
| XLSX | openpyxl | 工作表、单元格、公式值 | 公式结果和复杂格式处理有限 |
| PPTX | python-pptx | 幻灯片文本、表格 | 图片中文字需要 OCR |
| TXT/MD/JSON/CSV | 标准库与文本解析 | 文本、结构化字段 | 编码、非法格式、空文件 |

## 2.3 Flask Web 框架与异步任务

Flask 是 Python 生态中常用的轻量 Web 框架，具有路由定义简单、扩展灵活、适合快速构建接口服务等特点。本文系统使用 Flask 提供页面访问、文件处理、任务轮询、批量测试、依赖检查和发票记录管理接口。由于文档解析和 OCR 可能耗时较长，如果直接在同步请求中完成所有处理，前端会长时间等待，用户体验较差。因此，系统同时提供同步接口和异步任务接口。前端上传文件后，后端保存临时文件、创建任务编号并启动后台处理线程，前端再根据任务编号轮询处理状态。

当前系统采用进程内存保存任务状态，优点是实现简单、依赖少、适合毕业设计演示和单机部署；不足是服务重启后任务状态会丢失，不适合多实例生产环境。后续若面向真实生产部署，可以进一步引入 Redis、Celery、消息队列和对象存储。

## 2.4 SQLite 数据库与 CSV 导出

SQLite 是轻量级关系型数据库，具有无需独立服务、部署简单、适合本地单机应用等特点。DocFlow 使用 SQLite 保存发票结构化识别结果，数据库文件位于项目 data 目录。系统将发票代码、号码、日期、金额、税额、合计、购销方信息、置信度、字段数量、原始文本和创建时间等信息保存到 invoice_records 表中，并提供分页查询、详情查看、删除、清空和 CSV 导出功能。对于毕业设计系统而言，SQLite 能够满足演示和小规模数据管理需求，同时避免引入复杂数据库运维。

## 2.5 本章小结

本章介绍了系统实现所依赖的核心技术。OCR 技术解决图片和扫描件中的文字识别问题，PDF 与 Office 解析库解决结构化电子文档中的内容提取问题，Flask 提供 Web 接口与任务调度能力，SQLite 负责发票识别结果的轻量化持久化。这些技术共同构成了 DocFlow 系统的基础支撑。

# 3 系统需求分析与总体设计

## 3.1 需求分析

系统的目标用户主要是需要处理办公文档和票据材料的普通业务人员、财务人员以及系统演示人员。用户希望通过浏览器上传文档，系统自动识别文件类型并提取文本内容；当文件为发票图片或包含发票内容的文档时，系统能够进一步识别发票并提取关键字段；识别结果需要能够在页面上查看，并支持保存、查询和导出。

功能需求包括：文件上传与格式识别、多格式文档解析、图片 OCR 识别、PDF 模式选择、发票字段提取、发票记录保存、历史记录查询、记录删除、CSV 导出、批量测试、依赖检查和错误提示。非功能需求包括：界面操作简单、处理结果可追溯、模块具备可扩展性、系统可本地运行、支持 Docker 部署、对损坏文件和不支持格式给出清晰提示。

| 需求类别 | 具体需求 | 对应模块 |
| --- | --- | --- |
| 文档处理 | 支持 PDF、Word、Excel、PPT、文本、图片等格式 | docflow_core.py、services/ocr.py |
| 发票识别 | 判断是否为发票并提取关键字段 | invoice_extractor.py |
| 数据管理 | 保存、查询、删除和导出发票记录 | invoice_db.py、routes/invoice.py |
| Web 交互 | 上传文件、查看进度、展示结果 | frontend/doc_tool.html、routes/process.py |
| 测试评测 | 批量样本测试和发票字段评测 | scripts/run_batch_tests.py、scripts/run_invoice_evaluation.py |
| 部署运维 | 本地启动和 Docker 轻量部署 | app.py、Dockerfile、render.yaml |

图 3-1 系统用例图可在定稿中绘制为：用户连接“上传文档”“查看处理结果”“启用发票识别”“查看历史记录”“导出 CSV”“运行批量测试”“检查系统依赖”等用例；系统管理员或演示人员可以额外连接“查看报告”“调试依赖环境”等用例。

## 3.2 系统总体架构

DocFlow 采用分层架构设计。最上层是前端交互层，负责文件上传、模式选择、任务状态显示、结果展示和发票记录操作；第二层是 Web 接口层，基于 Flask 提供 HTTP 路由；第三层是业务服务层，负责异步任务管理、批量测试任务、OCR 调度和发票字段融合；第四层是文档解析核心层，负责按照文件类型调用不同解析器；第五层是数据持久化层，负责将发票记录保存到 SQLite 数据库。

图 3-2 系统五层架构图可按如下结构绘制：

```text
前端交互层: frontend/doc_tool.html
        ↓
Web 接口层: docflow/webapp/routes/*
        ↓
业务服务层: process_jobs、batch_jobs、ocr、ocr_engines、invoice_merge
        ↓
核心处理层: docflow_core.py、invoice_extractor.py
        ↓
数据与文件层: uploads_temp、reports、data/invoices.db
```

这种设计将请求处理、业务执行、文档解析和数据存储分开，降低了模块之间的耦合。与早期把大量逻辑集中在 app.py 的方式相比，当前结构更便于论文阐述，也更利于后续维护。

## 3.3 模块划分

系统主要模块包括入口模块、共享基础模块、Web 应用模块、业务服务模块、文档解析模块、发票提取模块、数据库模块、前端模块和脚本评测模块。入口模块 app.py 作为兼容启动入口；docflow/paths.py、settings.py 和 runtime.py 统一处理路径、配置和运行环境；docflow/webapp/routes 负责 HTTP 请求；docflow/webapp/services 负责业务执行；docflow_core.py 负责多格式文档解析；invoice_extractor.py 负责发票字段抽取；invoice_db.py 负责发票记录持久化；scripts 目录提供批测和评测能力。

| 模块 | 主要文件 | 职责 |
| --- | --- | --- |
| 启动入口 | app.py | 启动 Flask 应用，兼容旧入口 |
| 基础配置 | docflow/paths.py、settings.py、runtime.py | 路径、环境变量、日志和线程限制 |
| Web 路由 | docflow/webapp/routes/*.py | 接收请求、整理参数、返回 JSON |
| 业务服务 | docflow/webapp/services/*.py | 任务管理、OCR 调度、批测控制、字段融合 |
| 文档解析 | docflow_core.py | 多格式解析、摘要、关键词、格式化输出 |
| 发票抽取 | invoice_extractor.py | 发票判定和字段规则抽取 |
| 数据存储 | invoice_db.py | SQLite 保存、查询、删除、CSV 导出 |
| 前端页面 | frontend/doc_tool.html | 用户交互和结果展示 |

## 3.4 数据库设计

系统使用 SQLite 数据库保存发票记录，核心表为 invoice_records。该表以自增 id 作为主键，保存文件信息、发票字段、置信度、字段数量、原始文本和创建时间。数据库设计遵循轻量够用原则，既能满足毕业设计演示中的查询和导出需求，又不引入额外数据库服务。

| 字段名 | 类型 | 说明 | 是否主键 |
| --- | --- | --- | --- |
| id | INTEGER | 自增记录编号 | 是 |
| file_name | TEXT | 原始文件名 | 否 |
| file_type | TEXT | 文件类型 | 否 |
| file_size | INTEGER | 文件大小 | 否 |
| invoice_code | TEXT | 发票代码 | 否 |
| invoice_number | TEXT | 发票号码 | 否 |
| invoice_date | TEXT | 开票日期 | 否 |
| amount | TEXT | 不含税金额 | 否 |
| tax | TEXT | 税额 | 否 |
| total | TEXT | 价税合计 | 否 |
| buyer_name | TEXT | 购买方名称 | 否 |
| seller_name | TEXT | 销售方名称 | 否 |
| buyer_tax_id | TEXT | 购买方税号 | 否 |
| seller_tax_id | TEXT | 销售方税号 | 否 |
| confidence | TEXT | 识别置信度 | 否 |
| field_count | INTEGER | 已提取字段数量 | 否 |
| raw_text | TEXT | OCR 或解析得到的原始文本片段 | 否 |
| created_at | TIMESTAMP | 创建时间 | 否 |

## 3.5 关键流程设计

普通文档处理流程为：用户在前端选择文件并上传；后端根据扩展名识别文件类型；非图片文件进入 DocFlowProcessor，由对应解析器提取文本、表格和元数据；服务层补充统一统计信息；如果用户启用发票识别，则将提取文本交给 InvoiceExtractor；若识别为发票，则将结构化字段写入 SQLite；前端通过任务接口轮询结果并展示。

图片 OCR 与发票识别流程为：用户上传图片；后端识别为图片扩展名后调用 process_image_ocr；OCR 服务根据配置顺序尝试 RapidOCR、Tesseract 等引擎；每个候选 OCR 文本都会经过发票字段评分；系统选择得分较高的文本作为主结果，并在必要时对税号等关键字段进行补强；最终返回 OCR 文本、发票字段、统计信息和保存状态。

批量测试流程为：用户或开发者启动批量测试脚本；脚本遍历样本目录或评测清单；对每个样本调用文档处理或 OCR 流程；将结果与预期进行比对；最后在 reports 目录生成 Markdown、HTML、JSON 和 CSV 报告。该流程使系统质量验证从人工观察变为可复现的数据统计。

# 4 系统实现

## 4.1 开发环境与技术选型

本文系统使用 Python 作为主要开发语言，后端 Web 框架选用 Flask，数据库选用 SQLite，前端采用原生 HTML、CSS 和 JavaScript，部署支持 Docker 和 Render。OCR 方面，轻量云部署默认依赖 RapidOCR 与 Tesseract，本地环境可按需要启用 PaddleOCR 或 EasyOCR。文档解析方面，系统使用 pdfplumber、PyMuPDF、pypdfium2、python-docx、openpyxl 和 python-pptx 等库。

| 类别 | 技术选型 | 项目中的作用 |
| --- | --- | --- |
| 开发语言 | Python 3.9+ | 后端服务、文档解析、评测脚本 |
| Web 框架 | Flask | HTTP 接口、页面服务、任务接口 |
| OCR | RapidOCR、Tesseract | 图片和扫描件文字识别 |
| PDF | pdfplumber、PyMuPDF、pypdfium2 | PDF 文本提取和页面渲染 |
| Office | python-docx、openpyxl、python-pptx | Word、Excel、PPT 内容解析 |
| 数据库 | SQLite | 发票记录持久化 |
| 前端 | HTML、CSS、JavaScript | 上传、轮询、展示和记录管理 |
| 部署 | Docker、gunicorn、Render | 轻量云部署和演示 |

## 4.2 多格式文档处理引擎实现

多格式文档处理引擎位于 docflow_core.py，其中 DocFlowProcessor 是统一处理入口。系统首先根据文件扩展名选择对应解析器，然后由解析器返回文本、表格和元数据。解析完成后，核心处理器会生成关键词、摘要和统计信息，并根据输出参数转换为 TXT、JSON、Markdown 或 CSV 等格式。

```python
class DocFlowProcessor:
    def process(self, file_path, output_format="json", pdf_mode=DEFAULT_PDF_MODE):
        ext = Path(file_path).suffix.lower()
        parser = self.parsers.get(ext)
        if not parser:
            raise ValueError(f"不支持的文件格式: {ext}")
        text, tables, metadata = parser.parse(file_path, pdf_mode=pdf_mode)
        return self._build_result(text, tables, metadata, output_format)
```

PDFParser 是系统中最复杂的解析器。针对不同场景，系统提供 accurate、balanced 和 fast 三种 PDF 模式。accurate 模式优先保留较完整的表格和文本提取过程，适合质量要求较高的文档；balanced 模式在速度和准确性之间折中，是默认模式；fast 模式倾向于使用更快的解析路径和更轻量的 OCR 策略，适合大文件或快速预览场景。三种模式的存在使用户可以根据文档规模和处理目标进行选择。

Office 解析模块分别针对 Word、Excel 和 PPT 实现。WordParser 负责读取段落和表格，ExcelParser 负责遍历工作表和单元格，PPTXParser 负责读取幻灯片文本框和表格内容。TextParser 则统一处理 TXT、Markdown、JSON 和 CSV 等文本类文件，同时对编码和非法格式做异常处理。

## 4.3 OCR 多引擎调度与字段融合实现

图片 OCR 逻辑位于 docflow/webapp/services/ocr.py 和 ocr_engines.py，发票场景下的字段评分和融合逻辑位于 invoice_merge.py。系统没有把 OCR 调用直接写在路由中，而是通过服务层进行封装，使路由只需要关心请求和响应，OCR 服务负责引擎选择、缓存、候选结果评分和字段合并。

系统对 OCR 候选文本进行发票字段评分。评分时不仅考虑 OCR 文本长度，还会调用 InvoiceExtractor 抽取发票字段，并根据发票代码、发票号码、日期、金额、合计、购销方等字段的出现情况累计权重。如果候选文本被判断为发票，或者同时包含号码、金额、购销方等关键字段，则得分会增加。该策略使系统在多 OCR 结果中优先选择更有利于发票结构化抽取的文本，而不只是选择字符数量最多的文本。

```python
_OCR_INVOICE_FIELD_WEIGHTS = {
    "invoice_code": 30,
    "invoice_number": 30,
    "invoice_date": 14,
    "amount": 12,
    "tax": 8,
    "total": 14,
    "buyer_name": 10,
    "seller_name": 10,
}
```

税号是发票识别中的关键字段，通常由 18 位数字和大写字母组成。OCR 在识别税号时容易将 O 与 0、I 与 1、S 与 5、B 与 8 等字符混淆。系统在 invoice_merge.py 中维护了模糊字符映射，并对不同 OCR 候选结果中的税号进行比较和修复。当主结果中税号长度不足或存在疑似混淆字符时，系统会尝试从其他候选结果和局部裁剪识别结果中补充字符，从而提升关键字段的完整性。

## 4.4 发票字段提取模块实现

发票字段提取模块位于 invoice_extractor.py，是本文系统面向业务场景的核心模块。该模块首先对输入文本进行清洗和分行处理，然后依次尝试提取发票代码、发票号码、日期、金额、税额、价税合计、购销方名称、购销方税号、发票类型、机器编号和校验码等字段。每个字段提取函数都返回字段值、中文标签、置信度和来源信息，便于前端展示和后续评测。

发票判定并不是简单依赖单个关键词，而是综合考虑票头、号码、日期、金额、税号和购销方等特征。当文本中同时出现多个关键字段时，系统将其判定为发票并提升置信度。对于金额字段，系统需要区分“不含税金额”“税额”和“价税合计”；对于名称字段，系统需要结合“购买方”“销售方”等上下文位置，避免把商品名称误识别为公司名称。

```python
extractors = [
    ("invoice_code", self._find_invoice_code),
    ("invoice_number", self._find_invoice_number),
    ("invoice_date", self._find_invoice_date),
    ("amount", self._find_amount),
    ("tax", self._find_tax),
    ("total", self._find_total),
    ("buyer_name", self._find_buyer_name),
    ("seller_name", self._find_seller_name),
]
```

字段提取后的结果会被标准化为统一字典结构。该结构既能被前端直接展示，也能被 invoice_db.py 保存到数据库，还能被评测脚本与真值清单逐字段比对。统一的数据结构降低了 Web 展示、数据库持久化和实验评测之间的适配成本。

## 4.5 Web 接口与异步任务实现

Web 应用实际位于 docflow/webapp 包。core.py 负责创建 Flask app、启用 CORS、初始化日志和运行目录；routes 目录按功能拆分为 common、process、batch 和 invoice；services 目录实现具体业务。这样的拆分使 HTTP 层和业务层职责更加清晰。

| 接口 | 方法 | 功能 |
| --- | --- | --- |
| / | GET | 返回前端页面 |
| /process/start | POST | 启动异步文档处理任务 |
| /process/<job_id> | GET | 查询任务状态和处理结果 |
| /process/<job_id>/cancel | POST | 取消任务 |
| /process | POST | 同步处理单个文件 |
| /run-batch-tests | POST | 启动批量测试 |
| /invoices | GET | 分页查询发票记录 |
| /invoices/<id> | GET/DELETE | 查询或删除单条发票记录 |
| /invoices/export | GET | 导出 CSV |

异步任务实现采用后台线程和内存字典保存任务状态。处理开始时，系统生成 job_id，并将状态置为 running；处理过程中更新进度、日志和阶段信息；处理完成后写入 result；出现异常时写入 error。前端通过轮询接口获取任务状态，从而避免长时间阻塞上传请求。

## 4.6 数据持久化与导出实现

invoice_db.py 封装了发票记录的 CRUD 操作。初始化时自动创建 invoice_records 表；保存记录时从 InvoiceExtractor 的 fields 字典中提取各列值；查询时支持分页和关键词搜索；导出时将数据库记录转换为 CSV 字符串。由于 CSV 是办公场景中常见的数据交换格式，导出功能可以方便用户将识别结果导入 Excel 或后续财务系统。

```python
CREATE TABLE IF NOT EXISTS invoice_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name TEXT NOT NULL,
    invoice_code TEXT,
    invoice_number TEXT,
    invoice_date TEXT,
    amount TEXT,
    tax TEXT,
    total TEXT,
    buyer_name TEXT,
    seller_name TEXT,
    raw_text TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
```

## 4.7 前端页面实现

前端入口为 frontend/doc_tool.html，采用单页应用方式组织。页面包含文件上传区、PDF 模式选择、发票识别开关、处理日志、结果展示、批量测试入口、依赖检查和发票历史记录面板。前端通过 fetch 调用后端接口，并根据 job_id 周期性轮询处理进度。识别完成后，页面会将文本、统计信息和发票字段分区展示，用户可以直观看到系统是否识别为发票以及各字段的提取结果。

前端没有引入复杂框架，主要原因是本课题重点在文档处理和发票抽取，原生 HTML 与 JavaScript 已能满足演示需要。这样既降低了部署复杂度，也便于将系统打包到轻量 Docker 镜像中。

## 4.8 部署实现

系统支持本地启动和轻量云部署。本地开发时可以直接执行 python app.py，然后在浏览器访问 http://127.0.0.1:5000。云部署方面，项目提供 Dockerfile 和 render.yaml，生产入口使用 gunicorn，依赖集合使用 requirements-cloud.txt。云环境下默认采用 RapidOCR 和 Tesseract 作为 OCR 组合，并通过线程限制和图片降采样降低资源占用。

# 5 系统测试与结果分析

## 5.1 测试环境与测试方法

系统测试采用功能测试、边界测试和专项评测结合的方法。功能测试关注常见文档是否能够成功解析，边界测试关注空文件、损坏文件、伪造扩展名和不支持格式是否能给出合理错误，专项评测关注发票判定和字段抽取准确性。测试数据主要来自 sample_data 目录下的通用样本、边界样本，以及 发票样本 目录下的发票 PDF、图片和 Excel 真值清单。

测试报告由脚本自动生成，避免只依赖人工观察。run_batch_tests.py 用于通用批量测试，run_evaluation_set.py 用于清单式规则校验，run_invoice_evaluation.py 用于发票专项评测。每次测试会在 reports 目录下生成时间戳目录，并输出 Markdown、HTML、JSON 和 CSV 结果。

## 5.2 通用批量测试结果

根据 reports/batch_test_20260306_203110/report.md，通用批量测试共覆盖 51 个样本，其中常规样本 16 个、边界样本 35 个。测试总成功数为 41，失败数为 10；35 个带预期样本全部符合预期，不符合预期项为 0。失败样本主要是非法 JSON、伪造 Office 文件、损坏 PDF、空扫描 PDF 和不支持的 bin 文件，这些失败与测试预期一致，说明系统能够对异常输入进行识别和分类。

| 指标 | 数值 |
| --- | --- |
| 测试总数 | 51 |
| 成功 | 41 |
| 失败 | 10 |
| 已校验预期 | 35 |
| 符合预期 | 35 |
| 不符合预期 | 0 |

从分格式结果看，TXT、CSV、Markdown、图片类样本成功率较高；旧版 Office 文件、损坏文件和伪造扩展名文件存在失败，但错误信息能够说明原因。该结果符合系统定位：DocFlow 是一个面向常见办公格式和发票场景的自动化处理系统，而不是能无条件修复所有损坏文件的文件恢复工具。

## 5.3 发票专项评测结果

发票专项评测采用 invoice_image_manifest.csv，对 30 张由 PDF 首页渲染生成的发票图片进行 OCR 与字段提取验证。根据 reports/invoice_evaluation_20260314_140004/invoice_evaluation.md，评测样本总数为 30，已标注字段样本 30，已标注字段总数 249。系统在该评测集上取得 100.00% 的发票判定匹配率、100.00% 的字段覆盖率和 100.00% 的字段准确率，平均处理耗时为 15050.54 ms。

| 指标 | 结果 |
| --- | --- |
| 样本总数 | 30 |
| 已标注字段总数 | 249 |
| 发票判定匹配率 | 100.00% |
| 字段覆盖率 | 100.00% |
| 字段准确率 | 100.00% |
| 已标注样本通过率 | 100.00% |
| 平均处理耗时 | 15050.54 ms |

分类统计显示，image_electronic、image_normal、image_receipt 和 image_vat 四类样本均达到 100.00% 字段准确率。其中 image_vat 平均耗时最高，为 21027.78 ms，主要原因是增值税专用发票字段更多、版式信息更密集，OCR 和字段处理过程耗时更长。

| 样本类别 | 总数 | 字段准确率 | 平均耗时(ms) |
| --- | ---: | ---: | ---: |
| image_electronic | 7 | 100.00% | 16393.60 |
| image_normal | 8 | 100.00% | 11616.98 |
| image_receipt | 7 | 100.00% | 10800.42 |
| image_vat | 8 | 100.00% | 21027.78 |

需要说明的是，该评测集图片由 PDF 首页渲染生成，图像质量相对稳定，且真值来自 PDF 文本自动抽取或 Excel 清单，因此该结果能够证明系统在标准电子发票图片场景下的字段提取能力，但不能简单等同于所有真实拍照票据、低清扫描票据和严重倾斜票据的准确率。论文定稿时可以补充低质量图片或公开图片集上的鲁棒性测试，以增强结论说服力。

## 5.4 字段级结果分析

字段级统计显示，invoice_number、invoice_date、amount、total、buyer_name、seller_name 和 invoice_type 等字段在 30 个样本中均被完整提取并命中；invoice_code、buyer_tax_id 和 seller_tax_id 主要出现在增值税专用发票样本中，也全部命中；tax 字段在 15 个有标注样本中全部命中。该结果说明系统的规则抽取对标准票面中的关键字段具有较好的适应性。

| 字段 | 已标注数 | 已提取数 | 命中数 | 准确率 |
| --- | ---: | ---: | ---: | ---: |
| invoice_code | 8 | 8 | 8 | 100.00% |
| invoice_number | 30 | 30 | 30 | 100.00% |
| invoice_date | 30 | 30 | 30 | 100.00% |
| amount | 30 | 30 | 30 | 100.00% |
| tax | 15 | 15 | 15 | 100.00% |
| total | 30 | 30 | 30 | 100.00% |
| buyer_name | 30 | 30 | 30 | 100.00% |
| seller_name | 30 | 30 | 30 | 100.00% |
| buyer_tax_id | 8 | 8 | 8 | 100.00% |
| seller_tax_id | 8 | 8 | 8 | 100.00% |

## 5.5 系统不足与原因分析

尽管系统已经能够完成主要功能，但仍存在一些不足。首先，当前发票抽取主要依赖规则和启发式逻辑，对于严重倾斜、遮挡、低分辨率、强噪声和手写内容的票据，准确率可能下降。其次，异步任务状态保存在进程内存中，服务重启后任务状态会丢失，不适合多实例部署。再次，SQLite 适合本地演示和小规模数据管理，但如果未来需要支持多用户、权限控制和大规模并发，需要改造为 MySQL、PostgreSQL 或其他服务型数据库。最后，当前系统没有实现用户登录和权限隔离，真实生产环境还需要补充鉴权、审计日志、对象存储和监控告警。

## 5.6 本章小结

本章从通用批量测试和发票专项评测两个角度验证了系统功能。结果表明，DocFlow 能够正确处理常见样本，对异常输入给出符合预期的失败结果，并在标准发票图片评测集上完成发票判定和字段级抽取。测试结果为系统设计和实现提供了数据支撑，也指出了后续在复杂图像、任务持久化和生产化能力方面的改进方向。

# 结束语

本文围绕多格式文档自动化处理和发票信息提取需求，设计并实现了 DocFlow 系统。系统采用 Python 技术栈和分层架构，集成多格式文档解析、图片 OCR、发票字段抽取、SQLite 持久化、Web 前端展示、批量测试和 Docker 部署等能力，形成了从文档输入到结构化结果管理的完整闭环。

通过项目实现可以看出，轻量解析库与 OCR 引擎结合的路线适合本科毕业设计和中小型私有化应用场景。一方面，系统不依赖昂贵的商业 OCR 平台，能够在本地或轻量云环境运行；另一方面，规则抽取和字段融合逻辑具有较强可解释性，便于针对发票版式持续优化。实验结果表明，系统在已有样本集上能够达到较好的处理效果，具备较强的工程完整性和演示价值。

后续工作可以从四个方向继续改进。第一，补充真实拍照票据、低质量扫描件和复杂背景图片，提升评测样本多样性。第二，引入更稳定的任务队列和任务持久化机制，提升长任务处理能力。第三，增加用户认证、权限控制和多用户数据隔离，使系统更接近实际应用。第四，在现有多格式处理框架基础上扩展合同、报销单、档案材料等更多业务场景，进一步提升系统通用价值。

# 致谢

本课题从需求分析、系统设计、编码实现到测试评测的过程中，得到了指导教师、同学和开源社区的帮助。感谢指导教师在课题方向、系统设计和论文写作方面给予的指导；感谢同学在系统测试、样本整理和演示反馈方面提供的建议；感谢 Python、Flask、RapidOCR、Tesseract、PyMuPDF、pdfplumber、python-docx、openpyxl、python-pptx 等开源项目为本文系统实现提供了可靠基础。

同时，感谢学校和学院提供毕业设计实践机会，使本人能够将软件工程课程中学习到的需求分析、模块设计、接口开发、数据管理、测试验证和部署运维等知识应用到一个完整项目中。通过本课题的完成，本人对文档处理、OCR 识别、结构化信息提取和 Web 系统开发有了更加系统的认识。

# 参考文献

[1] Devlin J, Chang M W, Lee K, et al. BERT: Pre-training of Deep Bidirectional Transformers for Language Understanding[C]//Proceedings of NAACL-HLT 2019. Minneapolis: ACL, 2019: 4171-4186.

[2] Xu Y, Li M, Cui L, et al. LayoutLM: Pre-training of Text and Layout for Document Image Understanding[C]//Proceedings of KDD 2020. New York: ACM, 2020: 1192-1200.

[3] Xu Y, Lv T, Cui L, et al. LayoutLMv2: Multi-modal Pre-training for Visually-Rich Document Understanding[C]//Proceedings of ACL-IJCNLP 2021. Online: ACL, 2021: 2579-2591.

[4] Huang Y, Lv T, Cui L, et al. LayoutLMv3: Pre-training for Document AI with Unified Text and Image Masking[J/OL]. arXiv:2204.08387, 2022.

[5] Kim G, Hong T, Yim M, et al. OCR-free Document Understanding Transformer[J/OL]. arXiv:2111.15664, 2022.

[6] Blecher L, Cucurull G, Scialom T, et al. Nougat: Neural Optical Understanding for Academic Documents[J/OL]. arXiv:2308.13418, 2023.

[7] Lyu C, Luo J, Huang T, et al. DocLLM: A Layout-Aware Generative Language Model for Multimodal Document Understanding[J/OL]. arXiv:2401.00908, 2024.

[8] jsvine. pdfplumber Documentation[EB/OL]. [2026-05-06]. https://github.com/jsvine/pdfplumber.

[9] Artifex. PyMuPDF Documentation[EB/OL]. [2026-05-06]. https://pymupdf.readthedocs.io/.

[10] python-openxml. python-docx Documentation[EB/OL]. [2026-05-06]. https://python-docx.readthedocs.io/.

[11] openpyxl. openpyxl Documentation[EB/OL]. [2026-05-06]. https://openpyxl.readthedocs.io/.

[12] scanny. python-pptx Documentation[EB/OL]. [2026-05-06]. https://python-pptx.readthedocs.io/.

[13] Pallets. Flask Documentation[EB/OL]. [2026-05-06]. https://flask.palletsprojects.com/.

[14] Tesseract OCR. User Manual[EB/OL]. [2026-05-06]. https://tesseract-ocr.github.io/tessdoc/.

[15] RapidAI. RapidOCR[EB/OL]. [2026-05-06]. https://github.com/RapidAI/RapidOCR.

[16] Docker Inc. Docker Documentation[EB/OL]. [2026-05-06]. https://docs.docker.com/.

[17] McKinney W. Python for Data Analysis[M]. 3rd ed. Sebastopol: O'Reilly Media, 2022.

[18] Géron A. Hands-On Machine Learning with Scikit-Learn, Keras, and TensorFlow[M]. 3rd ed. Sebastopol: O'Reilly Media, 2022.

[19] 李航. 统计学习方法[M]. 2版. 北京: 清华大学出版社, 2019.

[20] 王海峰, 吴华. 自然语言处理研究综述[J]. 中国计算机学会通讯, 2021, 17(10): 14-23.

[21] 张子良, 陈文锋. 基于深度学习的文档版面分析方法研究[J]. 计算机工程与应用, 2022, 58(3): 112-120.

[22] Shi B, Bai X, Yao C. An End-to-End Trainable Neural Network for Image-Based Sequence Recognition and Its Application to Scene Text Recognition[J]. IEEE Transactions on Pattern Analysis and Machine Intelligence, 2017, 39(11): 2298-2304.

# 附录

## 附录 A 主要接口清单

| 接口 | 方法 | 说明 |
| --- | --- | --- |
| /process/start | POST | 启动异步文档处理任务 |
| /process/<job_id> | GET | 查询任务进度与结果 |
| /process | POST | 同步处理单个文件 |
| /run-batch-tests | POST | 启动批量测试 |
| /invoices | GET | 查询发票记录 |
| /invoices/export | GET | 导出发票记录 CSV |

## 附录 B 论文定稿待补充清单

1. 补齐封面中的姓名、学号、专业班级、指导教师和职称。
2. 在 Word 中根据学校要求更新目录页码。
3. 根据第 3 章文字绘制用例图、系统架构图、模块图、处理流程图和数据库表结构图。
4. 补充前端页面截图、识别结果截图、发票历史记录截图和测试报告截图。
5. 如重新运行评测脚本，需要同步更新第 5 章中的测试时间、样本数量、准确率和耗时。
6. 按导师意见调整题目，当前题目可改为“DocFlow多格式文档处理工具设计与实现”或“基于多引擎OCR融合的智能文档处理系统设计与实现”。
