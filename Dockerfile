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
        # LibreOffice and dependencies
        libreoffice \
        libreoffice-writer \
        libreoffice-impress \
        libreoffice-draw \
        libreoffice-java-common \
        libreoffice-core \
        libreoffice-common \
        libreoffice-calc \
        # Java runtime
        openjdk-11-jre-headless \
        # --- FONT INSTALLATION ---
        ttf-mscorefonts-installer \
        fonts-liberation \
        fonts-dejavu \
        # Utility to manage font cache
        fontconfig \
        # --- PDF and image processing ---
        poppler-utils \
        ghostscript \
        # --- X11 dependencies for headless LibreOffice ---
        libsm6 \
        libxext6 \
        libxrender1 \
        libgl1 \
        # --- Needed for font rendering/management ---
        libfreetype6 \
        # --- Needed for python-magic (ADDED) ---
        libmagic1 \
        # --- Curl (Used in HEALTHCHECK) ---
        curl \
    # Clean up apt cache
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    # --- UPDATE FONT CACHE ---
    && fc-cache -f -v \
    # --- Create necessary directories ---
    && mkdir -p /app/templates /app/static /app/uploads \
    # Set proper permissions
    && chmod 755 /app/uploads

# Copy requirements first to leverage Docker cache
COPY requirements.txt .

# Install Python dependencies
# Consider adding --require-hashes if you generate a hash file for extra security
RUN pip3 install --no-cache-dir --upgrade pip && \
    pip3 install --no-cache-dir -r requirements.txt

# Copy application files
COPY templates/ /app/templates/
COPY static/ /app/static/
COPY app.py /app/

# Set environment variables for potential LibreOffice use
ENV HOME=/tmp
# Optional: Isolate user profile for LibreOffice
# ENV LIBREOFFICE_PROFILE=/tmp/libreoffice_profile
# Note: PATH update
ENV PATH="/usr/lib/libreoffice/program:$PATH"

# Expose the Flask port
EXPOSE 5003

# Run the application using Waitress
# Ensure waitress is in requirements.txt (it is)
# Using 0.0.0.0 to bind to all interfaces inside the container
CMD ["waitress-serve", "--host=0.0.0.0", "--port=5003", "--threads=4", "app:app"]

# Health check (Adjusted to use the correct port and assume waitress startup time)
# Removed curl dependency from healthcheck to simplify, relying on waitress/app exit code
# HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
#   CMD exit 0 # Simplified - rely on container orchestration health checks if possible

# If you need a curl based healthcheck, ensure curl is installed (it is) and adjust host/port:
HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
   CMD curl -f http://localhost:5003/ || exit 1 # Check index route