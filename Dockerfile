# Use Ubuntu 22.04 as base image
FROM ubuntu:22.04

# Set environment variables to prevent interactive prompts
ENV DEBIAN_FRONTEND=noninteractive \
    TZ=Etc/UTC

# Create working directory
WORKDIR /app

# Install system dependencies in a single RUN layer to reduce image size
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    # Python and pip
    python3 \
    python3-pip \
    # python3-venv # Không thực sự cần trong container
    # LibreOffice and dependencies
    libreoffice \
    # Các thành phần này thường được cài cùng 'libreoffice' nhưng để rõ ràng cũng không sao
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-draw \
    libreoffice-java-common \
    # libreoffice-base # Có thể không cần thiết
    # libreoffice-core # Đã bao gồm trong 'libreoffice'
    libreoffice-common \
    # libreoffice-calc # Có thể không cần thiết
    # unoconv # Không còn dùng, xóa đi
    # Java runtime cho LibreOffice (OpenJDK 11 hoặc 17 đều ổn)
    openjdk-17-jre-headless \ # Hoặc openjdk-11-jre-headless
    # PDF and image processing
    poppler-utils \ # Cần cho pdf2image
    # tesseract-ocr # Không còn dùng, xóa đi
    # X11 dependencies for headless LibreOffice (Giữ lại các thư viện này)
    libsm6 \
    libxext6 \
    libxrender1 \
    libgl1 \
    # Curl cần cho HEALTHCHECK
    curl \
    # Clean up apt cache
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    # Create necessary directories
    && mkdir -p /app/templates /app/static /app/uploads \
    # Set proper permissions (777 đơn giản cho container, có thể chặt hơn nếu muốn)
    && chmod 777 /app/uploads

# Copy requirements first to leverage Docker cache
COPY requirements.txt .

# Install Python dependencies (sẽ cài waitress từ requirements.txt đã cập nhật)
RUN pip3 install --no-cache-dir --upgrade pip && \
    pip3 install --no-cache-dir -r requirements.txt

# Copy application files
COPY templates/ /app/templates/
COPY static/ /app/static/
COPY app.py /app/

# Set environment variables for LibreOffice user profile
ENV HOME=/tmp \
    LIBREOFFICE_PROFILE=/tmp/libreoffice_profile
    # PATH="/usr/lib/libreoffice/program:$PATH" # Dòng này vẫn OK, dù find_libreoffice() có thể tự tìm

# Health check (dùng cổng mặc định 5003, Railway sẽ tự điều hướng hoặc inject PORT)
HEALTHCHECK --interval=30s --timeout=15s --start-period=10s --retries=3 \
    CMD curl -f http://localhost:5003/health || exit 1

# Expose the default Flask port (Railway sẽ dùng $PORT được inject)
EXPOSE 5003

# Run the application using waitress (thông qua logic trong app.py)
CMD ["python3", "app.py"]