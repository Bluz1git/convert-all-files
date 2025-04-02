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
        # python3-venv # Removed: Not strictly necessary inside the final container image
        # LibreOffice and dependencies (Keep this comprehensive list)
        libreoffice \
        libreoffice-writer \
        libreoffice-impress \
        libreoffice-draw \
        libreoffice-java-common \
        # libreoffice-base # Removed: Likely not needed for conversion
        libreoffice-core \
        libreoffice-common \
        libreoffice-calc \
        # unoconv # Removed: soffice is used directly, reduces minor dependency clutter
        # Java runtime (Keep: often helps LibreOffice)
        openjdk-11-jre-headless \
        # --- FONT INSTALLATION --- (Keep all these)
        ttf-mscorefonts-installer \
        fonts-liberation \
        fonts-dejavu \
        # Utility to manage font cache
        fontconfig \
        # -------------------------
        # PDF and image processing
        poppler-utils \
        # tesseract-ocr # Removed: Keep comment if OCR might be needed later, but remove package for now
        # X11 dependencies for headless LibreOffice (Keep all these)
        libsm6 \
        libxext6 \
        libxrender1 \
        libgl1 \
        # Needed for some font rendering/management (Keep)
        libfreetype6 \
        # Curl (Keep: Used in HEALTHCHECK)
        curl \
    # Clean up apt cache
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    # --- UPDATE FONT CACHE --- (Very Important! Keep)
    && fc-cache -f -v \
    # -------------------------
    # Create necessary directories
    && mkdir -p /app/templates /app/static /app/uploads \
    # Set proper permissions (Keep 755)
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

# Set environment variables for potential LibreOffice use (good practice - Keep)
ENV HOME=/tmp
# Optional: Isolate user profile for LibreOffice (Keep commented unless needed)
# ENV LIBREOFFICE_PROFILE=/tmp/libreoffice_profile
# Note: PATH update (Keep: doesn't hurt, good backup)
ENV PATH="/usr/lib/libreoffice/program:$PATH"

# Health check (Keep - Looks good)
HEALTHCHECK --interval=30s --timeout=30s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:5003/health || exit 1

# Expose the Flask port (Keep)
EXPOSE 5003

# Run the application (Keep)
# Using Gunicorn for production is recommended over Flask's built-in server
# Add gunicorn to requirements.txt if you use this
# CMD ["gunicorn", "--bind", "0.0.0.0:5003", "--workers", "2", "--threads", "4", "--timeout", "120", "app:app"]
# For now, keep the simple python3 command:
CMD ["python3", "app.py"]