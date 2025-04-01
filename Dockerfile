FROM python:3.9-slim

WORKDIR /app

# Cài đặt các gói cần thiết
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    libreoffice \
    libsm6 \
    libxext6 \
    libxrender1 \
    wget \
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