# P2 resume scorer, containerised so macOS and Windows run the same thing.
#
# The Windows .bat launchers and app.py's GUI are not in here. This image is the CLI,
# bot.py, which is what the pipeline actually runs. Everything the scorer needs works
# identically on Linux: the worker lock already branches to fcntl off Windows, and no
# code resolves a platform-specific path.
#
#   docker compose build
#   docker compose run --rm p2 --doctor
#
# See README.md for the full flow.

FROM python:3.11-slim

# Tesseract is the OCR fallback only. RapidOCR is primary and needs no system binary,
# but without this the fallback would fail silently on a scanned PDF the primary
# engine could not read. English data only, which is what the resumes are in.
RUN apt-get update \
    && apt-get install --no-install-recommends -y tesseract-ocr tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Dependencies first, so a source edit does not re-run the install layer.
COPY requirements-docker.txt .
RUN pip install --no-cache-dir -r requirements-docker.txt

COPY . .

# Ollama runs on the HOST, not in this container: the model is multi-gigabyte and on a
# Mac it needs Metal acceleration a Linux container cannot reach. host.docker.internal
# resolves to the host from inside Docker Desktop on both macOS and Windows.
ENV HIRING_OLLAMA_HOST=http://host.docker.internal:11434 \
    PYTHONUNBUFFERED=1

# No CMD that scores anything. An image whose default action writes to live SharePoint
# is a foot-gun; you pass the flag you mean.
ENTRYPOINT ["python", "bot.py"]
CMD ["--help"]
