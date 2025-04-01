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
RUN which soffice && ln -sf $(which soffice) /usr/local/bin/soffice || echo "soffice not found"

# Sao chép requirements trước
COPY requirements.txt .

# Cài đặt Python dependencies với xử lý lỗi tốt hơn - tách riêng các thư viện có thể gây xung đột
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir wheel setuptools && \
    pip install --no-cache-dir flask werkzeug PyPDF2 && \
    pip install --no-cache-dir pdf2docx opencv-python-headless && \
    pip install --no-cache-dir pdfplumber reportlab && \
    pip install --no-cache-dir PyMuPDF img2pdf python-docx

# Tạo thư mục uploads và thiết lập quyền
RUN mkdir -p /app/uploads && chmod 755 /app/uploads

# Sao chép mã nguồn
COPY . .

# Thiết lập biến môi trường
ENV PORT=5003
ENV PYTHONUNBUFFERED=1

# Kiểm tra libreoffice có hoạt động không
RUN soffice --version || echo "WARNING: LibreOffice not working correctly"

# Chạy ứng dụng
CMD ["python", "app.py"]