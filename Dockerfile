FROM ubuntu:22.04

WORKDIR /app

# Cài đặt Python và pip trước tiên
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    python3 \
    python3-pip \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Sau đó cài đặt các phụ thuộc khác
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-draw \
    libreoffice-java-common \
    libreoffice-base \
    libreoffice-core \
    libreoffice-common \
    libreoffice-calc \
    unoconv \
    openjdk-11-jre \
    libsm6 \
    libxext6 \
    libxrender1 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Tạo cấu trúc thư mục
RUN mkdir -p /app/templates /app/static /app/uploads

# Cài đặt các Python packages
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy templates và static files
COPY templates/* /app/templates/
COPY static/* /app/static/

# Copy toàn bộ mã nguồn
COPY . .

# Chạy ứng dụng
CMD ["python3", "app.py"]