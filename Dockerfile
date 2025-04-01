FROM python:3.9-slim

WORKDIR /app

# Cài đặt các gói cần thiết, bao gồm libreoffice
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    libreoffice \
    fonts-liberation \
    fonts-dejavu \
    && apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Tạo symbolic link cho soffice để đảm bảo nó nằm trong PATH
RUN ln -sf $(which soffice) /usr/local/bin/soffice

# Cài đặt Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Tạo thư mục uploads và thiết lập quyền
RUN mkdir -p /app/uploads && chmod 755 /app/uploads

# Sao chép mã nguồn
COPY . .

# Thiết lập biến môi trường
ENV PORT=5003

# Kiểm tra libreoffice có hoạt động không
RUN soffice --version

# Chạy ứng dụng
CMD ["python", "app.py"]