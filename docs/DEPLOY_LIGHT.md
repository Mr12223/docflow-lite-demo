# DocFlow 轻量部署说明

## 1. 适用场景

当前这套部署方案适合：

- 功能演示
- 小规模在线试用
- 毕设验收或答辩环境
- 本地 Docker 复现实验

它的目标不是完整生产架构，而是在尽量少的系统依赖下，把当前仓库能跑通的核心能力部署出来。

## 2. 当前轻量镜像实际包含的内容

仓库根目录 `Dockerfile` 基于 `python:3.11-slim` 构建，当前默认包含：

- Flask 服务
- Gunicorn 生产启动
- PDF 相关依赖：`pdfplumber`、`PyMuPDF`、`pypdfium2`
- Office 解析依赖：`python-docx`、`openpyxl`、`python-pptx`
- OCR 依赖：`rapidocr`、`onnxruntime`、`pytesseract`
- 系统级 Tesseract OCR
- 中文 `chi_sim` 与英文 `eng` 语言包

当前轻量镜像默认不包含：

- `EasyOCR`
- `PaddleOCR`
- `LibreOffice`

因此，代码层虽然保留了多 OCR 引擎和旧版 Office 转换接口，但在默认云镜像中，实际主打能力是：

- `RapidOCR`
- `Tesseract`
- 常见 Office Open XML 文档解析

## 3. 默认运行策略

当前 Dockerfile 为云环境做了几项默认优化：

- `PORT=8000`
- `DOCFLOW_DEFAULT_PDF_MODE=fast`
- `DOCFLOW_DISABLE_PDF_TABLES=1`
- `DOCFLOW_IMAGE_OCR_ORDER=rapidocr,tesseract,easyocr`
- `DOCFLOW_CLOUD_DEPLOYMENT=1`
- 图片 OCR 默认启用降采样和更激进的长边限制
- 线程相关环境变量被压到较低水平，降低小机器上的资源争抢
- `WEB_CONCURRENCY=1`

需要注意：

- 虽然 OCR 顺序里保留了 `easyocr`，但轻量镜像默认没有安装它
- 所以容器中的实际回退链通常是 `RapidOCR -> Tesseract`

## 4. 本地 Docker 运行

在项目根目录执行：

```bash
docker build -t docflow-lite .
docker run --rm -p 8000:8000 docflow-lite
```

启动后访问：

```text
http://127.0.0.1:8000
```

### 4.1 建议的持久化运行方式

如果你希望保留发票记录和测试报告，建议挂载卷：

```bash
docker run --rm -p 8000:8000 \
  -v ${PWD}/data:/app/data \
  -v ${PWD}/reports:/app/reports \
  -v ${PWD}/uploads_temp:/app/uploads_temp \
  docflow-lite
```

否则容器重建后，下面这些内容都会丢失：

- `data/invoices.db`
- `reports/`
- `uploads_temp/`

## 5. Render 部署

仓库已经提供 `render.yaml`，当前配置要点是：

- `runtime: docker`
- `healthCheckPath: /`
- 通过 `Dockerfile` 构建
- 构建时忽略 `reports/**`、`uploads_temp/**`、`__pycache__/**` 等目录

适合的部署方式：

1. 把仓库推到 Git 平台
2. 在 Render 中使用 Blueprint 或直接创建 Docker Web Service
3. 选择仓库根目录部署
4. 等待平台完成构建并分配 URL

当前 `render.yaml` 没有声明持久化磁盘，所以 Render 上的报告和 SQLite 数据默认也是临时的。

## 6. Railway / 通用 Docker 平台

Railway、Zeabur、普通 Docker 主机都可以直接复用同一个 `Dockerfile`。

通用思路：

1. 连接仓库
2. 让平台识别根目录 `Dockerfile`
3. 暴露 `PORT`
4. 部署完成后访问首页 `/`

只要平台支持 Docker，并且能给容器分配至少一个对外端口，这个方案就可以运行。

## 7. Ubuntu 云服务器部署

服务器安装 Docker 后执行：

```bash
git clone <你的仓库地址>
cd ShiXunClaud
docker build -t docflow-lite .
docker run -d \
  --name docflow-lite \
  --restart unless-stopped \
  -p 8000:8000 \
  -v $(pwd)/data:/app/data \
  -v $(pwd)/reports:/app/reports \
  -v $(pwd)/uploads_temp:/app/uploads_temp \
  docflow-lite
```

访问方式：

```text
http://服务器IP:8000
```

## 8. 当前部署限制

这套轻量方案当前的限制比较明确：

- 没有外部任务队列，长耗时任务仍在 Web 进程侧完成
- 没有对象存储，上传文件和报告默认落在本地磁盘
- 没有鉴权、限流、多租户隔离
- 大 PDF 在云端仍可能较慢
- 旧版 `.doc/.xls` 在未安装 LibreOffice 时只能部分兼容或降级处理

如果你把它当成正式生产方案，风险主要在“长任务阻塞”和“数据持久化”两点。

## 9. 推荐的云端使用方式

如果目标是“尽快稳定跑起来”，建议保持当前默认配置，只做两类增强：

### 9.1 先补持久化

至少持久化下面两个目录：

- `/app/data`
- `/app/reports`

如果不关心临时文件历史，`/app/uploads_temp` 可以不持久化。

### 9.2 按需要调 PDF 模式

- 追求速度：保持 `DOCFLOW_DEFAULT_PDF_MODE=fast`
- 追求稳妥：改成 `balanced`
- 追求尽量高的 OCR 覆盖：改成 `accurate`

在小规格云主机上，不建议把默认模式直接调成 `accurate` 再跑大批量任务。

## 10. 排障建议

部署后可以按这个顺序检查：

1. 打开 `/`，确认前端页面能正常返回
2. 请求 `/system/dependencies`，确认运行时依赖状态
3. 上传一张图片，验证 `RapidOCR / Tesseract` 是否可用
4. 上传一个 PDF，确认 `fast` 模式下能返回文本
5. 如果旧版 Office 文件异常，再访问 `/debug-doc` 检查 `LibreOffice / soffice` 状态

常见结论：

- 图片 OCR 不可用：优先检查 `RapidOCR` 和 `Tesseract`
- 旧版 `.doc/.xls` 失败：通常是缺少 LibreOffice
- 发票记录丢失：通常是容器没有挂载 `/app/data`

## 11. 适合继续升级的方向

如果后续要从“演示版”走向“更稳定的服务版”，推荐按这个顺序升级：

1. 增加持久化存储
2. 接入任务队列
3. 增加鉴权和任务隔离
4. 为旧版 Office 增加 LibreOffice 容器层
5. 按需扩展 EasyOCR / PaddleOCR 的独立镜像
