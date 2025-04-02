# Use Ubuntu 22.04 as base image
FROM ubuntu:22.04

# Set environment variables to prevent interactive prompts
ENV DEBIAN_FRONTEND=noninteractive \
    TZ=Etc/UTC \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

# Create working directory
WORKDIR /app

# Install system dependencies in a single RUN layer to reduce image size
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    # Python and pip
    python3 \
    python3-pip \
    # LibreOffice and dependencies
    libreoffice \
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-draw \
    libreoffice-java-common \
    libreoffice-common \
    # Java runtime for LibreOffice
    openjdk-17-jre-headless \
    # PDF and image processing <--- Dòng comment đứng một mình đã bị xóa
    poppler-utils \ # Cần cho pdf2image
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
    # Set proper permissions
    && chmod 777 /app/uploads

# Copy requirements first to leverage Docker cache
COPY requirements.txt .

# Install Python dependencies
RUN pip3 install --no-cache-dir --upgrade pip && \
    pip3 install --no-cache-dir -r requirements.txt

# Copy application files
COPY templates/ /app/templates/
COPY static/ /app/static/
COPY app.py /app/

# Set environment variables for LibreOffice user profile
ENV HOME=/tmp \
    LIBREOFFICE_PROFILE=/tmp/libreoffice_profile

# Health check
HEALTHCHECK --interval=30s --timeout=15s --start-period=10s --retries=3 \
    CMD curl -f http://localhost:5003/health || exit 1

# Expose the default Flask port
EXPOSE 5003

# Run the application using waitress
CMD ["python3", "app.py"]