# DocFlow 

DocFlow 是一个基于 Python 的多格式文档自动化处理与内容提取工具，当前重点支持：

- 多格式文档解析：`pdf`、`docx/doc`、`xlsx/xls`、`pptx`、图片
- OCR 文本识别：`RapidOCR`、`Tesseract` 等组合链路
- 发票结构化提取：发票号码、日期、金额、税额、购销方、税号等字段
- 发票记录管理：SQLite 存储、列表查询、删除、CSV 导出
- 批量评测：字段级评测、报告导出

## 目录结构

```text
ShiXunClaud/
├─ app.py                      # Flask Web 入口与接口编排
├─ docflow_core.py             # 文档解析核心流程
├─ docflow_support.py          # 依赖检查、错误增强、辅助函数
├─ invoice_extractor.py        # 发票字段提取规则
├─ invoice_db.py               # 发票持久化与导出
├─ docflow/                    # 新增：共享配置、路径与运行时工具
│  ├─ paths.py
│  ├─ settings.py
│  └─ runtime.py
├─ frontend/                   # 前端页面
├─ scripts/                    # 批测与评测脚本
├─ docs/                       # 部署、论文、协作说明
├─ sample_data/                # 通用测试样本
├─ 发票样本/                    # 发票专项测试与说明
├─ uploads_temp/               # 运行时上传与 OCR 缓存
├─ reports/                    # 评测与批测报告
└─ data/                       # SQLite 数据文件
```

## 本地启动

```powershell
.\.venv\Scripts\python.exe app.py
```

浏览器打开：`http://127.0.0.1:5000`

## 推荐开发流程

1. 新业务规则优先落到对应模块，不要直接堆到路由函数里
2. 公共路径、环境变量、运行时初始化统一放在 `docflow/`
3. 新增脚本放在 `scripts/`，新增说明放在 `docs/`
4. 运行产物只写入 `uploads_temp/`、`reports/`、`data/`

## 常用命令

```powershell
.\.venv\Scripts\python.exe -m py_compile app.py docflow_core.py invoice_extractor.py invoice_db.py
.\.venv\Scripts\python.exe .\scripts\run_invoice_evaluation.py .\发票样本\invoice_test_manifest.csv
```

## 协作规范

详细协作规范见 `docs/开发协作规范.md`
项目分层与目录职责见 `docs/项目架构说明.md`
