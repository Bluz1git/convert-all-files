FROM python:3.9-slim

# Thiết lập biến môi trường
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=8080 \
    UPLOAD_FOLDER=/app/uploads \
    GUNICORN_WORKERS=2 \
    GUNICORN_THREADS=2 \
    GUNICORN_TIMEOUT=120

# Cài đặt các phụ thuộc hệ thống - giảm kích thước image
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
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    && rm -rf /var/cache/apt/*

# Tạo và cấu hình thư mục ứng dụng
RUN useradd -m appuser && \
    mkdir -p ${UPLOAD_FOLDER} && \
    chown appuser:appuser ${UPLOAD_FOLDER}

WORKDIR /app

# Copy và cài đặt requirements trước để tận dụng cache Docker
COPY --chown=appuser:appuser requirements.txt .
RUN pip install --no-cache-dir --upgrade pip setuptools wheel && \
    pip install --no-cache-dir -r requirements.txt

# Copy mã nguồn ứng dụng
COPY --chown=appuser:appuser . .

# Chuyển sang user không phải root
USER appuser

# Kiểm tra sức khỏe ứng dụng (sử dụng wget thay vì curl để giảm dependencies)
HEALTHCHECK --interval=30s --timeout=3s \
    CMD wget --no-verbose --tries=1 --spider http://localhost:$PORT/health || exit 1

# Mở cổng (chỉ mang tính khai báo)
EXPOSE $PORT

# Chạy ứng dụng với Gunicorn (sử dụng biến môi trường cho cấu hình)
CMD gunicorn --bind 0.0.0.0:$PORT \
    --workers ${GUNICORN_WORKERS} \
    --threads ${GUNICORN_THREADS} \
    --timeout ${GUNICORN_TIMEOUT} \
    --access-logfile - \
    --error-logfile - \
    app:app