"""Minimal HTTP scoring server for container platforms (Fly.io, Railway, Render, GCP Cloud Run, AWS App Runner).

Endpoints:
  POST /score         score the SharePoint queue (returns JSON summary)
  POST /score?dry=1   dry run (read + log only, no writes)
  GET  /health        health check (container orchestrators ping this)

Env:  PORT (default 8080), HIRING_SERVERLESS=1 (set by the Dockerfile),
      plus the SharePoint creds (TENANT_ID, CLIENT_ID, CLIENT_SECRET, ...).
"""

import json
import logging
import os
import sys
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, parse_qs

sys.path.insert(0, str(Path(__file__).resolve().parent))
os.environ.setdefault("HIRING_SERVERLESS", "1")

PORT = int(os.getenv("PORT", "8080"))


class _Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if urlparse(self.path).path in ("/health", "/healthz", "/"):
            self._json(200, {"status": "ok"})
        else:
            self._json(404, {"error": "not found"})

    def do_POST(self):
        path = urlparse(self.path)
        if path.path == "/score":
            from hiring_agent.config import SHAREPOINT_CONFIGURED
            if not SHAREPOINT_CONFIGURED:
                self._json(500, {"ok": False, "error": "SharePoint not configured"})
                return
            from hiring_agent.sharepoint_scoring import score_from_sharepoint
            dry = "dry" in parse_qs(path.query) or "dry_run" in parse_qs(path.query)
            try:
                result = score_from_sharepoint(dry_run=dry)
                ok = "error" not in result
                self._json(200 if ok else 500, {"ok": ok, **result})
            except Exception as e:
                logging.exception("score_from_sharepoint failed")
                self._json(500, {"ok": False, "error": str(e)})
        else:
            self._json(404, {"error": "not found"})

    def _json(self, code, data):
        body = json.dumps(data).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


if __name__ == "__main__":
    print(f"HiringAgent scoring server listening on :{PORT}")
    ThreadingHTTPServer(("", PORT), _Handler).serve_forever()
