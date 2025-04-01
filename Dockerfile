FROM python:3.9-slim

WORKDIR /app

# 1. Cài đặt các phụ thuộc cơ bản trước
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    wget \
    ca-certificates \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# 2. Cài LibreOffice phiên bản tối giản
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice-writer \  # Chỉ cài bộ Writer
    libreoffice-headless \  # Không cần GUI
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# 3. Cài các thư viện hỗ trợ
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libsm6 \
    libxext6 \
    libxrender1 \
    default-jre-headless \  # Java runtime nhẹ
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Các bước tiếp theo...