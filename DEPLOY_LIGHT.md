# DocFlow 轻量云部署演示

这套方案用于**毕业设计演示 / 功能验收 / 小规模在线展示**，目标是：

- 部署简单
- 体积尽量轻
- 保留 PDF / 图片 OCR 能力
- 不强依赖本地 Windows 环境

## 方案说明

本项目更适合用 **Docker 容器** 做轻量云演示，而不是纯 Serverless：

- 文档解析和 OCR 属于**长耗时任务**
- 需要 **Tesseract** 这类系统级依赖
- 运行过程中会产生临时文件和测试报告

因此，这里给的是一套**通用 Docker 部署方案**，可用于：

- Render 的 Docker Web Service
- Railway 的 Docker 部署
- 小型 Ubuntu 云服务器
- 本地 Docker 演示

## 云演示版默认能力

当前 `Dockerfile` 默认包含：

- Flask 后端服务
- Gunicorn 生产启动
- Tesseract OCR
- 中文 `chi_sim` + 英文 `eng` 语言包
- 常见文档解析依赖

可直接支持或较稳定支持：

- `pdf`
- `docx`
- `xlsx`
- `pptx`
- `txt`
- `csv`
- `json`
- 图片 OCR

## 当前刻意保持“轻量”的地方

为了保证云端镜像更轻、部署更快，当前**没有默认内置**：

- `easyocr`
- `LibreOffice`

这意味着：

- PDF / 图片 OCR 默认走 `pytesseract + Tesseract`
- `.doc` 老格式支持为**降级可用**，不是云演示版主打能力

如果你后面要做“增强版部署”，再考虑追加这两部分更合适。

## 一、本地先跑通 Docker 演示

在项目根目录执行：

```bash
docker build -t docflow-lite .
docker run --rm -p 8000:8000 docflow-lite
```

浏览器打开：

```text
http://127.0.0.1:8000
```

如果云平台会自动注入 `PORT`，容器会自动监听对应端口。

## 二、部署到 Render / Railway

### Render

1. 新建 Web Service
2. 选择使用仓库或上传代码
3. 让平台识别根目录 `Dockerfile`
4. 部署完成后直接访问分配的 URL

### Railway

1. 新建项目
2. 连接仓库或上传项目目录
3. 使用 Dockerfile 部署
4. 等待构建完成并打开生成的域名

## 三、部署到 Ubuntu 云服务器

服务器安装 Docker 后执行：

```bash
git clone <你的仓库地址>
cd ShiXunClaud
docker build -t docflow-lite .
docker run -d --name docflow-lite -p 8000:8000 --restart unless-stopped docflow-lite
```

然后访问：

```text
http://服务器IP:8000
```

## 四、云演示版限制

这套方案适合“展示能跑起来、核心功能可演示”，但还不是完整生产架构。

当前限制主要有：

- 上传文件和报告保存在容器内，**重启后可能丢失**
- 大 PDF 的 OCR 仍然会比较慢
- 没有任务队列，长任务仍在 Web 进程内执行
- 没有对象存储、数据库、鉴权和限流

## 五、毕业设计演示时怎么说

你可以这样描述这部分：

> 系统采用 Docker 进行轻量化云部署演示，在 Linux 容器中集成 Flask、Gunicorn 与 Tesseract OCR，实现了 PDF、Office 文档与图片内容提取的在线运行验证。为控制镜像体积与部署复杂度，演示版保留核心识别链路，增强型 OCR 与旧版 Office 转换能力作为可扩展模块按需启用。

## 六、后续可升级方向

如果后面你想把“演示版”继续升级，我建议按这个顺序：

1. 接入对象存储，解决上传文件与报告持久化
2. 接入任务队列，解决大文件阻塞问题
3. 增加登录和任务隔离
4. 增加 LibreOffice 容器层，强化 `.doc` 兼容
5. 单独做“高精度 OCR 增强版”镜像，把 `easyocr` 放进去

## 七、相关平台文档

- Render Docker 部署文档：https://render.com/docs/docker
- Railway Docker 部署文档：https://docs.railway.com/guides/dockerfiles
