FROM python:3.9-slim

WORKDIR /app

# Cài đặt các gói cần thiết, bỏ openjdk-11-jre
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    libreoffice && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Cài đặt Python dependencies
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir Flask==2.1.3 Werkzeug==2.1.2 pdf2docx==0.5.8 img2pdf PyMuPDF

# Sao chép mã nguồn
COPY .. .

# Thiết lập biến môi trường
ENV PORT=5003

# Chạy ứng dụng
CMD ["python", "app.py"]