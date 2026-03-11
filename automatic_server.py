from __future__ import annotations

import html
import os
import subprocess
import sys
import tempfile
from email.parser import BytesParser
from email.policy import default
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


def parse_multipart(content_type: str, body: bytes) -> dict[str, list[dict[str, object]]]:
    message = BytesParser(policy=default).parsebytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    )
    if not message.is_multipart():
        raise ValueError("multipart/form-data 요청만 허용됩니다.")

    files: dict[str, list[dict[str, object]]] = {}
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        files.setdefault(name, []).append(
            {
                "filename": part.get_filename(),
                "content": part.get_payload(decode=True) or b"",
            }
        )
    return files


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
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            self.respond_text(400, "multipart/form-data 요청만 허용됩니다.")
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self.respond_text(400, "잘못된 Content-Length 입니다.")
            return
        if content_length <= 0:
            self.respond_text(400, "본문이 비어 있습니다.")
            return

        try:
            form = parse_multipart(content_type, self.rfile.read(content_length))
        except ValueError as exc:
            self.respond_text(400, str(exc))
            return

        target_items = form.get("targetFile", [])
        source_items = form.get("sourceFiles", [])

        if not target_items or not target_items[0].get("filename"):
            self.respond_text(400, "기준 파일(targetFile)이 없습니다.")
            return

        if not source_items:
            self.respond_text(400, "참고 파일(sourceFiles)이 없습니다.")
            return

        source_items = [s for s in source_items if s.get("filename")]
        if not source_items:
            self.respond_text(400, "참고 파일(sourceFiles)이 없습니다.")
            return

        with tempfile.TemporaryDirectory(prefix="balju_auto_") as tmp:
            tmp_dir = Path(tmp)
            target_item = target_items[0]
            target_name = safe_name(str(target_item["filename"]))
            target_path = tmp_dir / target_name
            with target_path.open("wb") as f:
                f.write(target_item["content"])  # type: ignore[arg-type]

            for s in source_items:
                src_name = safe_name(str(s["filename"]))
                src_path = tmp_dir / src_name
                with src_path.open("wb") as f:
                    f.write(s["content"])  # type: ignore[arg-type]

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
