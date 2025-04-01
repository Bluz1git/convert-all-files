FROM ubuntu:22.04

WORKDIR /app

# Cài đặt các phụ thuộc cơ bản và Python
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    python3.11 \
    python3-pip \
    python3-dev \
    build-essential \
    wget \
    ca-certificates \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Cài đặt LibreOffice đầy đủ, Java, và các gói phụ trợ
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-draw \
    libreoffice-java-common \
    libreoffice-base \
    openjdk-11-jre \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Cài đặt các thư viện hỗ trợ X11 (cho LibreOffice headless)
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    libsm6 \
    libxext6 \
    libxrender1 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Cài đặt các Python packages
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy toàn bộ mã nguồn
COPY . .

# Chạy ứng dụng
CMD ["python3", "app.py"]