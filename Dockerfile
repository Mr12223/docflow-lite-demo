FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=8000 \
    TESSDATA_PREFIX=/usr/share/tesseract-ocr/5/tessdata \
    DOCFLOW_DEFAULT_PDF_MODE=fast \
    DOCFLOW_DISABLE_PDF_TABLES=1 \
    DOCFLOW_IMAGE_OCR_ORDER=rapidocr,tesseract,easyocr \
    DOCFLOW_IMAGE_OCR_MAX_LONG_EDGE=1600 \
    DOCFLOW_IMAGE_OCR_FAST_EDGE_TRIGGER=2400 \
    DOCFLOW_IMAGE_OCR_HUGE_LONG_EDGE=1280 \
    DOCFLOW_IMAGE_OCR_GRAYSCALE=1 \
    DOCFLOW_IMAGE_OCR_AUTOCONTRAST=1 \
    DOCFLOW_IMAGE_OCR_PSM=6 \
    DOCFLOW_ENABLE_RAPIDOCR=1 \
    DOCFLOW_RAPIDOCR_MAX_LONG_EDGE=1400 \
    DOCFLOW_RAPIDOCR_FAST_EDGE_TRIGGER=2200 \
    DOCFLOW_RAPIDOCR_HUGE_LONG_EDGE=1100 \
    DOCFLOW_CLOUD_DEPLOYMENT=1 \
    DOCFLOW_CLOUD_IMAGE_OCR_DOWNSAMPLE=1 \
    DOCFLOW_CLOUD_IMAGE_OCR_MAX_LONG_EDGE=1100 \
    DOCFLOW_CLOUD_IMAGE_OCR_FAST_EDGE_TRIGGER=1800 \
    DOCFLOW_CLOUD_IMAGE_OCR_HUGE_LONG_EDGE=900 \
    WEB_CONCURRENCY=1 \
    GUNICORN_THREADS=1

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1 \
    libglib2.0-0 \
    libgomp1 \
    tesseract-ocr \
    tesseract-ocr-chi-sim \
    tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

COPY requirements-cloud.txt ./

RUN python -m pip install --upgrade pip && \
    pip install -r requirements-cloud.txt && \
    python -c "import cv2, onnxruntime, rapidocr; from rapidocr import RapidOCR; print('RapidOCR build probe OK:', getattr(rapidocr, '__version__', 'unknown'))"

COPY . .

RUN mkdir -p uploads_temp reports && \
    useradd -m -u 10001 appuser && \
    chown -R appuser:appuser /app

USER appuser

EXPOSE 8000

CMD ["sh", "-c", "gunicorn -w ${WEB_CONCURRENCY:-1} --timeout 1800 --bind 0.0.0.0:${PORT:-8000} app:app"]
