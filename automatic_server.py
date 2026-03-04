from __future__ import annotations

import cgi
import html
import os
import shutil
import subprocess
import sys
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse


BASE_DIR = Path(__file__).resolve().parent
HTML_PATH = BASE_DIR / "automatic.html"
MAPPING_SCRIPT = BASE_DIR / "apply_mapping.py"
HOST = os.getenv("HOST", "0.0.0.0")
PORT = int(os.getenv("PORT", "8765"))


def safe_name(name: str) -> str:
    return Path(name).name.replace("\x00", "")


class Handler(BaseHTTPRequestHandler):
    server_version = "BaljuAutoServer/1.0"

    def end_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/" or parsed.path == "/automatic.html":
            self.serve_html()
            return
        self.send_error(404, "Not Found")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/generate":
            self.handle_generate()
            return
        self.send_error(404, "Not Found")

    def serve_html(self) -> None:
        if not HTML_PATH.exists():
            self.send_error(500, "automatic.html not found")
            return
        data = HTML_PATH.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def handle_generate(self) -> None:
        ctype, pdict = cgi.parse_header(self.headers.get("Content-Type", ""))
        if ctype != "multipart/form-data":
            self.respond_text(400, "multipart/form-data 요청만 허용됩니다.")
            return

        pdict["boundary"] = pdict["boundary"].encode("utf-8")
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            },
        )

        target_item = form["targetFile"] if "targetFile" in form else None
        source_items = form["sourceFiles"] if "sourceFiles" in form else None

        if target_item is None or not getattr(target_item, "filename", None):
            self.respond_text(400, "기준 파일(targetFile)이 없습니다.")
            return

        if source_items is None:
            self.respond_text(400, "참고 파일(sourceFiles)이 없습니다.")
            return

        if not isinstance(source_items, list):
            source_items = [source_items]
        source_items = [s for s in source_items if getattr(s, "filename", None)]
        if not source_items:
            self.respond_text(400, "참고 파일(sourceFiles)이 없습니다.")
            return

        with tempfile.TemporaryDirectory(prefix="balju_auto_") as tmp:
            tmp_dir = Path(tmp)
            target_name = safe_name(target_item.filename)
            target_path = tmp_dir / target_name
            with target_path.open("wb") as f:
                shutil.copyfileobj(target_item.file, f)

            for s in source_items:
                src_name = safe_name(s.filename)
                src_path = tmp_dir / src_name
                with src_path.open("wb") as f:
                    shutil.copyfileobj(s.file, f)

            cmd = [
                sys.executable,
                str(MAPPING_SCRIPT),
                "--base-dir",
                str(tmp_dir),
                "--target",
                target_name,
                "--backup-keep",
                "0",
            ]
            proc = subprocess.run(cmd, capture_output=True, text=True)
            if proc.returncode != 0:
                err_text = proc.stderr or proc.stdout or "처리 중 오류가 발생했습니다."
                self.respond_text(500, err_text)
                return

            if not target_path.exists():
                self.respond_text(500, "생성된 엑셀 파일을 찾지 못했습니다.")
                return

            out_name = f"generated_{target_name}"
            data = target_path.read_bytes()
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.send_header("Content-Length", str(len(data)))
            self.send_header(
                "Content-Disposition",
                f'attachment; filename="{out_name}"',
            )
            self.send_header("X-Output-Filename", out_name)
            self.end_headers()
            self.wfile.write(data)

    def respond_text(self, status: int, message: str) -> None:
        msg = html.escape(message)
        data = msg.encode("utf-8", errors="ignore")
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt: str, *args) -> None:
        sys.stdout.write("%s - - [%s] %s\n" % (self.client_address[0], self.log_date_time_string(), fmt % args))


def main() -> None:
    if not HTML_PATH.exists():
        raise FileNotFoundError(f"automatic.html not found: {HTML_PATH}")
    if not MAPPING_SCRIPT.exists():
        raise FileNotFoundError(f"apply_mapping.py not found: {MAPPING_SCRIPT}")

    server = ThreadingHTTPServer((HOST, PORT), Handler)
    listen_host = HOST if HOST != "0.0.0.0" else "127.0.0.1"
    print(f"Server started: http://{listen_host}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        print("Server stopped")


if __name__ == "__main__":
    main()
