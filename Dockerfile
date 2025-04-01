FROM python:3.9-slim

# Môi trường và biến
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# Sử dụng port mặc định của Railway
ARG PORT=5003
ENV PORT=${PORT}

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

# Expose port
EXPOSE ${PORT}

# Chi tiết logging và debugging
RUN echo "Port being used: $PORT"

# Sử dụng shell form để đảm bảo biến môi trường được mở rộng
CMD gunicorn --bind 0.0.0.0:$PORT --workers 4 --threads 2 --timeout 120 app:app