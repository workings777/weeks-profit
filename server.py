"""
로컬 서버 — weeks_profit 웹앱 + RPA 실행 엔드포인트
실행: python server.py
접속: http://localhost:8080
"""

import os, sys, subprocess, threading, json
from http.server import HTTPServer, SimpleHTTPRequestHandler

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RPA_SCRIPT = os.path.join(BASE_DIR, "rpa_fursys_v2.py")
PORT = 8080

_rpa_status = {"running": False, "log": ""}

class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_POST(self):
        if self.path == "/run-rpa":
            self._run_rpa()
        else:
            self.send_error(404)

    def do_GET(self):
        if self.path == "/rpa-status":
            self._rpa_status()
        else:
            super().do_GET()

    def _run_rpa(self):
        try:
            subprocess.Popen(
                f'start "" cmd /k ""{sys.executable}" "{RPA_SCRIPT}""',
                shell=True, cwd=BASE_DIR
            )
            self._json({"status": "started"})
        except Exception as e:
            self._json({"status": "error", "message": str(e)})

    def _rpa_status(self):
        self._json({
            "running": _rpa_status["running"],
            "log": _rpa_status["log"],
        })

    def _json(self, data):
        body = json.dumps(data, ensure_ascii=False).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass  # 콘솔 로그 억제


if __name__ == "__main__":
    os.chdir(BASE_DIR)
    print(f"서버 시작: http://localhost:{PORT}")
    HTTPServer(("", PORT), Handler).serve_forever()
