# Use Ubuntu 22.04 as base image
FROM ubuntu:22.04

# Set environment variables to prevent interactive prompts
ENV DEBIAN_FRONTEND=noninteractive \
    TZ=Etc/UTC \
    # Set Python to use UTF-8 encoding
    PYTHONIOENCODING=UTF-8 \
    PYTHONUNBUFFERED=1 \
    # LibreOffice environment
    HOME=/tmp \
    LIBREOFFICE_PROFILE=/tmp/libreoffice_profile \
    PATH="/usr/lib/libreoffice/program:$PATH"

# Create working directory
WORKDIR /app

# Install system dependencies in a single RUN layer to reduce image size
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    # Python and pip
    python3 \
    python3-pip \
    python3-venv \
    # LibreOffice and dependencies (minimal installation)
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-calc \
    libreoffice-base \
    libreoffice-java-common \
    # Java runtime
    openjdk-11-jre-headless \
    # PDF and image processing
    poppler-utils \
    ghostscript \
    imagemagick \
    tesseract-ocr \
    # Fonts
    fonts-liberation \
    fonts-dejavu \
    # X11 dependencies for headless LibreOffice
    libsm6 \
    libxext6 \
    libxrender1 \
    libgl1 \
    # Other utilities
    curl \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    # Create necessary directories
    && mkdir -p /app/templates /app/static /app/uploads \
    # Set proper permissions
    && chmod 755 /app/uploads \
    # Clean LibreOffice profile (helps prevent startup issues)
    && rm -rf /tmp/libreoffice_profile \
    # Create symlink for python3 to python
    && ln -s /usr/bin/python3 /usr/bin/python

# Copy requirements first to leverage Docker cache
COPY requirements.txt .

# Install Python dependencies
RUN pip3 install --no-cache-dir --upgrade pip && \
    pip3 install --no-cache-dir -r requirements.txt

# Copy application files
COPY templates/ /app/templates/
COPY static/ /app/static/
COPY app.py /app/

# Health check
HEALTHCHECK --interval=30s --timeout=30s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:5003/health || exit 1

# Expose the Flask port
EXPOSE 5003

# Run the application as non-root user for security
RUN useradd -m appuser && chown -R appuser:appuser /app
USER appuser

# Run the application
CMD ["python3", "app.py"]