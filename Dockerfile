FROM python:3.9-slim

WORKDIR /app

# Cài đặt các gói cần thiết từng bước
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    wget \
    ca-certificates \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Cài đặt libreoffice riêng
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Cài đặt các thư viện hỗ trợ
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libsm6 \
    libxext6 \
    libxrender1 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Tạo thư mục uploads
RUN mkdir -p /app/uploads

# Cài đặt Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Sao chép mã nguồn
COPY . .

# Chạy ứng dụng
CMD ["python", "app.py"]