from __future__ import annotations

import json
import mimetypes
import re
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from obsazovani import build_project, export_project_workbook
from obsazovani.i18n import t

ROOT = Path(__file__).resolve().parent
WEB_ROOT = ROOT / "web"


def slugify_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", value.strip())
    return cleaned.strip("-") or "obsazeni"


class AppHandler(BaseHTTPRequestHandler):
    server_version = "Obsazovani/0.1"

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path == "/":
            return self.serve_file(WEB_ROOT / "index.html")
        if path.startswith("/api/"):
            return self.send_error(HTTPStatus.NOT_FOUND, "Unknown API endpoint")

        target = (WEB_ROOT / path.lstrip("/")).resolve()
        if WEB_ROOT not in target.parents and target != WEB_ROOT:
            return self.send_error(HTTPStatus.FORBIDDEN)
        if not target.exists() or not target.is_file():
            return self.send_error(HTTPStatus.NOT_FOUND)
        self.serve_file(target)

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path == "/api/analyze":
            return self.handle_analyze()
        if path == "/api/export":
            return self.handle_export()
        self.send_error(HTTPStatus.NOT_FOUND, "Unknown API endpoint")

    def serve_file(self, target: Path) -> None:
        content_type, _ = mimetypes.guess_type(str(target))
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type or "application/octet-stream")
        self.end_headers()
        self.wfile.write(target.read_bytes())

    def read_json_body(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length) if length else b"{}"
        try:
            return json.loads(raw.decode("utf-8"))
        except json.JSONDecodeError as exc:
            raise ValueError(t("server.invalid_json")) from exc

    def send_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        encoded = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def handle_analyze(self) -> None:
        try:
            project = build_project(self.read_json_body())
        except Exception as exc:  # noqa: BLE001
            return self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        self.send_json(project)

    def handle_export(self) -> None:
        try:
            project = build_project(self.read_json_body())
            workbook = export_project_workbook(project)
        except Exception as exc:  # noqa: BLE001
            return self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

        filename = f"{slugify_filename(project['title'])}.xlsx"
        self.send_response(HTTPStatus.OK)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(len(workbook)))
        self.end_headers()
        self.wfile.write(workbook)


def main() -> None:
    host = "127.0.0.1"
    port = 8123
    server = ThreadingHTTPServer((host, port), AppHandler)
    print(f"Obsazování běží na http://{host}:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
