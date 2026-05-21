#!/usr/bin/env python3

import json
import subprocess
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = ROOT / "data" / "data.json"
BUILD_SCRIPT = ROOT / "scripts" / "build_dashboard_data.py"


class DashboardHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def _send_json(self, payload, status=200):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/api/data":
            if not DATA_PATH.exists():
                self._send_json({"status": "error", "message": "data.json not found"}, status=404)
                return
            self._send_json(json.loads(DATA_PATH.read_text(encoding="utf-8")))
            return
        if path == "/api/refresh":
            try:
                subprocess.run([sys.executable, str(BUILD_SCRIPT)], cwd=str(ROOT), check=True)
                payload = json.loads(DATA_PATH.read_text(encoding="utf-8"))
                self._send_json({"status": "success", "message": "Data refreshed", "data": payload})
            except subprocess.CalledProcessError as exc:
                self._send_json({"status": "error", "message": f"Refresh failed: {exc}"}, status=500)
            return
        super().do_GET()


def main():
    port = 8000
    server = ThreadingHTTPServer(("127.0.0.1", port), DashboardHandler)
    print(f"Dashboard server running at http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
