FROM python:3.9-slim

# Thiết lập biến môi trường
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    RAILWAY_PORT=5003 \
    UPLOAD_FOLDER=/app/uploads \
    GUNICORN_WORKERS=2 \
    GUNICORN_THREADS=2 \
    GUNICORN_TIMEOUT=120

# Cài đặt các phụ thuộc hệ thống
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    libreoffice \
    fonts-liberation \
    fonts-dejavu \
    libsm6 \
    libxext6 \
    libxrender1 \
    poppler-utils \
    wget \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    && rm -rf /var/cache/apt/*

# Tạo và cấu hình thư mục ứng dụng
RUN useradd -m appuser && \
    mkdir -p ${UPLOAD_FOLDER} && \
    chown appuser:appuser ${UPLOAD_FOLDER}

WORKDIR /app

# Copy và cài đặt requirements
COPY --chown=appuser:appuser requirements.txt .
RUN pip install --no-cache-dir --upgrade pip setuptools wheel && \
    pip install --no-cache-dir -r requirements.txt

# Copy mã nguồn ứng dụng
COPY --chown=appuser:appuser . .

# Chuyển sang user không phải root
USER appuser

# Chạy ứng dụng với Gunicorn
CMD gunicorn --bind 0.0.0.0:${RAILWAY_PORT} \
    --workers 2 \
    --threads 2 \
    --timeout 120 \
    --access-logfile - \
    --error-logfile - \
    app:app