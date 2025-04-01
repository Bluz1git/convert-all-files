FROM python:3.9-slim

# Tăng cường bảo mật và performance
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PORT=5003

WORKDIR /app

# Cài đặt các gói hệ thống
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    libreoffice \
    fonts-liberation \
    fonts-dejavu \
    curl \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Tạo symbolic link chắc chắn cho soffice
RUN ln -sf $(which soffice) /usr/local/bin/soffice

# Sao chép và cài đặt requirements
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip setuptools wheel && \
    pip install --no-cache-dir -r requirements.txt && \
    pip install --no-cache-dir gunicorn

# Sao chép mã nguồn
COPY . .

# Tạo thư mục uploads với quyền phù hợp
RUN mkdir -p /app/uploads && chmod 755 /app/uploads

# Kiểm tra LibreOffice
RUN soffice --version || echo "WARNING: LibreOffice might not work correctly"

# Healthcheck để kiểm tra ứng dụng
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD curl -f http://localhost:$PORT/ || exit 1

# Sử dụng Gunicorn để quản lý WSGI
CMD ["gunicorn", "--bind", "0.0.0.0:$PORT", "--workers", "4", "--threads", "2", "--timeout", "120", "app:app"]