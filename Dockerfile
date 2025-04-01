FROM python:3.9-slim

WORKDIR /app

# Cài đặt phụ thuộc cơ bản
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    wget \
    ca-certificates \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Thêm repository để cài đặt LibreOffice
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    apt-transport-https \
    gnupg \
    && echo "deb http://deb.debian.org/debian buster main contrib non-free" >> /etc/apt/sources.list.d/debian.list \
    && apt-get update

# Cài LibreOffice (bao gồm headless mode)
RUN apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-impress \
    libreoffice \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    # Thêm phần này sau khi cài đặt LibreOffice
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    default-jre \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Cài các thư viện hỗ trợ
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    libsm6 \
    libxext6 \
    libxrender1 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Cài đặt Python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy code
COPY . .

# Chạy ứng dụng
CMD ["python", "app.py"]