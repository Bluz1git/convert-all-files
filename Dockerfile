FROM python:3.9-slim-bullseye

WORKDIR /app

# Cài đặt các phụ thuộc hệ thống
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    wget \
    ca-certificates \
    libsm6 \
    libxext6 \
    libxrender1 \
    libreoffice-writer \
    libreoffice-headless \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Tạo thư mục uploads
RUN mkdir -p /app/uploads && chmod 777 /app/uploads

# Cài đặt thư viện Python
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Sao chép mã nguồn
COPY . .

# Giảm kích thước image
RUN apt-get remove -y build-essential python3-dev && \
    apt-get autoremove -y && \
    rm -rf /var/lib/apt/lists/*

# Chạy ứng dụng
CMD ["python", "app.py"]