FROM python:3.9-slim

# Environment variables
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=5003 \
    UPLOAD_FOLDER=/app/uploads \
    GUNICORN_WORKERS=2 \
    GUNICORN_THREADS=2 \
    GUNICORN_TIMEOUT=120

# Install system dependencies
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    libreoffice \
    fonts-liberation \
    fonts-dejavu \
    libsm6 \
    libxext6 \
    libxrender1 \
    poppler-utils \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    && rm -rf /var/cache/apt/*

# Create and configure application directory
RUN useradd -m appuser && \
    mkdir -p ${UPLOAD_FOLDER} && \
    chown appuser:appuser ${UPLOAD_FOLDER}

WORKDIR /app

# Copy and install requirements
COPY --chown=appuser:appuser requirements.txt .
RUN pip install --no-cache-dir --upgrade pip setuptools wheel && \
    pip install --no-cache-dir -r requirements.txt

# Copy application source
COPY --chown=appuser:appuser . .

# Switch to non-root user
USER appuser

# Health check
HEALTHCHECK --interval=30s --timeout=3s \
    CMD wget --no-verbose --tries=1 --spider http://localhost:$PORT/health || exit 1

# Expose port
EXPOSE $PORT

# Run application with explicit values
CMD gunicorn --bind 0.0.0.0:$PORT \
    --workers 2 \
    --threads 2 \
    --timeout 120 \
    --access-logfile - \
    --error-logfile - \
    app:app