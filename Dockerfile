FROM python:3.9-slim

WORKDIR /app

# 1. Cài đặt các phụ thuộc hệ thống theo từng bước
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    wget \
    ca-certificates \
    libsm6 \
    libxext6 \
    libxrender1 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# 2. Cài LibreOffice phiên bản nhẹ hơn
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice-writer \  # Chỉ cài bộ Writer thay vì toàn bộ LibreOffice
    libreoffice-headless \  # Chế độ không cần GUI
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# 3. Tạo thư mục uploads với quyền phù hợp
RUN mkdir -p /app/uploads && chmod 777 /app/uploads

# 4. Cài đặt thư viện Python
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# 5. Sao chép mã nguồn
COPY . .

# 6. Chạy ứng dụng
CMD ["python", "app.py"]
# Thêm vào cuối Dockerfile để giảm kích thước
RUN apt-get remove -y build-essential python3-dev && \
    apt-get autoremove -y && \
    rm -rf /var/lib/apt/lists/*