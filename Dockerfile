FROM python:3.9-slim

WORKDIR /app

# Cài đặt các gói cần thiết
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    libreoffice && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Cài đặt Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Sao chép mã nguồn
COPY . .

# Đảm bảo thư mục templates và uploads tồn tại
RUN mkdir -p templates uploads

# Thiết lập biến môi trường
ENV PORT=5003
ENV PYTHONUNBUFFERED=1

# Chạy ứng dụng
CMD ["python", "app.py"]