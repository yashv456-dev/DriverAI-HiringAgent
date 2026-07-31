FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
# Skip azure-functions (only needed for Azure Functions host, not Docker).
RUN grep -v '^azure-functions' requirements.txt > /tmp/reqs.txt \
    && pip install --no-cache-dir -r /tmp/reqs.txt \
    && rm /tmp/reqs.txt

COPY . .

ENV HIRING_SERVERLESS=1
ENV PORT=8080
EXPOSE 8080

# Default: HTTP server (for Fly.io, Railway, Render, GCP Cloud Run, AWS App Runner).
# Override to run as a one-shot or polling worker instead:
#   docker run ... hiring-agent python bot.py --score-sharepoint
#   docker run ... hiring-agent python bot.py --score-sharepoint --watch --interval 900
CMD ["python", "server.py"]
