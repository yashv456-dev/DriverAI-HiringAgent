# ── DriverAI Hiring Agent ──────────────────────────────────────────────────
# Containerized Dual-Phase Hiring Intelligence & Recruiter Dashboard
# Supports:
#   1. Web Dashboard (FastAPI + Glassmorphic UI) on port 8000
#   2. Headless CLI batch processing & SharePoint sync via bot.py
# ───────────────────────────────────────────────────────────────────────────

FROM python:3.11-slim

# System dependencies:
# - tesseract-ocr & tesseract-ocr-eng for image/scanned resume parsing fallback
# - curl for Docker HEALTHCHECK probes
RUN apt-get update \
    && apt-get install --no-install-recommends -y \
        tesseract-ocr \
        tesseract-ocr-eng \
        curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Cache layer: install dependencies first
COPY requirements-docker.txt .
RUN pip install --no-cache-dir -r requirements-docker.txt

# Copy all source files
COPY . .

# Set working directory to P2 where the application, static files, and config reside
WORKDIR /app/DriverAI_HiringAgent/HiringAgent_P2

# Configure environment defaults:
# - PYTHONUNBUFFERED: Immediate log flush to stdout/stderr
# - PYTHONPATH: Ensure local modules and packages are discoverable
# - HIRING_OLLAMA_HOST: Allows connecting to host machine Ollama if running locally
ENV PYTHONUNBUFFERED=1 \
    PYTHONPATH=/app/DriverAI_HiringAgent/HiringAgent_P2 \
    PORT=8000 \
    HIRING_OLLAMA_HOST=http://host.docker.internal:11434

# Expose web server port
EXPOSE 8000

# Health check to ensure web API is responsive
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:8000/health || exit 1

# Default command: Launch the Recruiter Web Application & Interactive Dashboard
CMD ["python", "web_server.py"]
