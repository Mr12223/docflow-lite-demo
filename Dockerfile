FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=8000 \
    TESSDATA_PREFIX=/usr/share/tesseract-ocr/5/tessdata \
    DOCFLOW_DEFAULT_PDF_MODE=fast \
    DOCFLOW_DISABLE_PDF_TABLES=1 \
    WEB_CONCURRENCY=2 \
    GUNICORN_THREADS=4

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-chi-sim \
    tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

COPY requirements-cloud.txt ./

RUN python -m pip install --upgrade pip && \
    pip install -r requirements-cloud.txt

COPY . .

RUN mkdir -p uploads_temp reports && \
    useradd -m -u 10001 appuser && \
    chown -R appuser:appuser /app

USER appuser

EXPOSE 8000

CMD ["sh", "-c", "gunicorn -w ${WEB_CONCURRENCY:-2} -k gthread --threads ${GUNICORN_THREADS:-4} --timeout 1800 --bind 0.0.0.0:${PORT:-8000} app:app"]
