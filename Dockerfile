# Use Ubuntu 22.04 as base image
FROM ubuntu:22.04

# Set environment variables to prevent interactive prompts
ENV DEBIAN_FRONTEND=noninteractive \
    TZ=Etc/UTC

# Create working directory
WORKDIR /app

# Install system dependencies, including fonts, in a single RUN layer
RUN apt-get update && \
    # Pre-accept the Microsoft EULA for ttf-mscorefonts-installer
    echo ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true | debconf-set-selections && \
    # Install packages
    apt-get install -y --no-install-recommends \
    # Python and pip
    python3 \
    python3-pip \
    python3-venv \
    # LibreOffice and dependencies
    libreoffice \
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-draw \
    libreoffice-java-common \
    libreoffice-base \
    libreoffice-core \
    libreoffice-common \
    libreoffice-calc \
    # unoconv might pull helpful dependencies, though soffice is used directly
    unoconv \
    # Java runtime (often helps LibreOffice)
    openjdk-11-jre \
    # --- FONT INSTALLATION ---
    # Microsoft Core Fonts (Crucial for Office document compatibility)
    ttf-mscorefonts-installer \
    # Other common fonts for better coverage
    fonts-liberation \
    fonts-dejavu \
    # Utility to manage font cache
    fontconfig \
    # -------------------------
    # PDF and image processing
    poppler-utils \
    # tesseract-ocr (Keep if you might need OCR later, otherwise removable)
    tesseract-ocr \
    # X11 dependencies for headless LibreOffice
    libsm6 \
    libxext6 \
    libxrender1 \
    libgl1 \
    # Needed for some font rendering/management
    libfreetype6 \
    # Clean up apt cache
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    # --- UPDATE FONT CACHE --- (Very Important!)
    && fc-cache -f -v \
    # -------------------------
    # Create necessary directories
    && mkdir -p /app/templates /app/static /app/uploads \
    # Set proper permissions (755 is generally safer than 777)
    && chmod 755 /app/uploads

# Copy requirements first to leverage Docker cache
COPY requirements.txt .

# Install Python dependencies
RUN pip3 install --no-cache-dir --upgrade pip && \
    pip3 install --no-cache-dir -r requirements.txt

# Copy application files
COPY templates/ /app/templates/
COPY static/ /app/static/
COPY app.py /app/

# Set environment variables for potential LibreOffice use (good practice)
ENV HOME=/tmp
# Optional: Isolate user profile for LibreOffice
# ENV LIBREOFFICE_PROFILE=/tmp/libreoffice_profile
# Note: PATH is less critical if app.py finds soffice dynamically, but doesn't hurt
ENV PATH="/usr/lib/libreoffice/program:$PATH"

# Health check
HEALTHCHECK --interval=30s --timeout=30s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:5003/health || exit 1

# Expose the Flask port
EXPOSE 5003

# Run the application
CMD ["python3", "app.py"]