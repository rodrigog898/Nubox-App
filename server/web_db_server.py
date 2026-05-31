from __future__ import annotations

import json
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib import error as urlerror
from urllib import parse as urlparse
from urllib import request as urlrequest

BASE_DIR = Path(__file__).resolve().parent.parent
HOST = "127.0.0.1"
PORT = 8765
NUBOX_BASE_URL = "https://app.nubox.com"
NUBOX_PDF_PATH = "/ServiFactura/paginas/dteDocumentosRecibidos.aspx/VerPDFDocManualesElectronicos"


def safe_parse_json(text: str):
    try:
        return json.loads(text)
    except Exception:
        return None


def extract_cookie_value(cookie_header: str, cookie_name: str) -> str:
    parts = [part.strip() for part in cookie_header.split(";") if part.strip()]
    for part in parts:
        if "=" not in part:
            continue
        raw_name, raw_value = part.split("=", 1)
        decoded_name = urlparse.unquote(raw_name)
        if decoded_name == cookie_name or raw_name == cookie_name:
            return urlparse.unquote(raw_value)
    return ""


def find_temp_pdf_path(value) -> str:
    if not isinstance(value, dict):
        return ""

    raw_d = value.get("d")
    if isinstance(raw_d, str):
        if raw_d.startswith("/"):
            return raw_d
        parsed = safe_parse_json(raw_d)
        if isinstance(parsed, dict):
            found = find_temp_pdf_path(parsed)
            if found:
                return found

    for nested in value.values():
        if isinstance(nested, dict):
            found = find_temp_pdf_path(nested)
            if found:
                return found

    return ""


class NuboxProxyHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def _set_cors(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _send_json(self, status: int, payload: dict) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self._set_cors()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _send_bytes(self, status: int, content_type: str, raw: bytes) -> None:
        self.send_response(status)
        self._set_cors()
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _read_json_body(self) -> dict:
        content_length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(content_length)
        return json.loads(raw.decode("utf-8"))

    def _perform_upstream_request(self, url: str, method: str, headers: dict, body: dict | None = None):
        data = json.dumps(body).encode("utf-8") if (method == "POST" and body is not None) else None
        req = urlrequest.Request(url, data=data, headers=headers, method=method)
        try:
            with urlrequest.urlopen(req, timeout=45) as upstream_resp:
                return (
                    upstream_resp.status,
                    upstream_resp.headers.get("Content-Type", "application/json; charset=utf-8"),
                    upstream_resp.read(),
                )
        except urlerror.HTTPError as exc:
            return (
                exc.code,
                exc.headers.get("Content-Type", "application/json; charset=utf-8"),
                exc.read(),
            )

    def _handle_nubox_proxy(self) -> None:
        try:
            payload = self._read_json_body()
            target = payload.get("target", {})
            headers = payload.get("headers", {})
            body = payload.get("body", {})

            base_url = str(target.get("baseUrl", "")).strip()
            path = str(target.get("path", "")).strip()
            method = str(target.get("method", "POST")).upper().strip()

            if not base_url or not path:
                raise ValueError("missing target.baseUrl or target.path")
            if method not in {"POST", "GET"}:
                raise ValueError("only POST or GET method is allowed")

            parsed_base = urlparse.urlparse(base_url)
            if parsed_base.scheme != "https" or parsed_base.netloc.lower() != "app.nubox.com":
                raise ValueError("target baseUrl must be https://app.nubox.com")

            upstream_url = f"{base_url}{path}"
            upstream_data = json.dumps(body).encode("utf-8") if method == "POST" else None

            allowed_headers = {
                "Content-Type",
                "Cookie",
                "X-Requested-With",
                "Accept",
                "Origin",
                "Referer",
                "User-Agent",
            }
            upstream_headers = {
                key: str(value)
                for key, value in headers.items()
                if isinstance(key, str) and key in allowed_headers
            }
            if method == "POST" and "Content-Type" not in upstream_headers:
                upstream_headers["Content-Type"] = "application/json;charset=UTF-8"

            req = urlrequest.Request(
                upstream_url,
                data=upstream_data,
                headers=upstream_headers,
                method=method,
            )

            try:
                with urlrequest.urlopen(req, timeout=45) as upstream_resp:
                    response_bytes = upstream_resp.read()
                    response_status = upstream_resp.status
                    response_content_type = upstream_resp.headers.get(
                        "Content-Type", "application/json; charset=utf-8"
                    )
            except urlerror.HTTPError as exc:
                response_bytes = exc.read()
                response_status = exc.code
                response_content_type = exc.headers.get(
                    "Content-Type", "application/json; charset=utf-8"
                )

            self.send_response(response_status)
            self._set_cors()
            self.send_header("Content-Type", response_content_type)
            self.send_header("Content-Length", str(len(response_bytes)))
            self.end_headers()
            self.wfile.write(response_bytes)
        except ValueError as exc:
            self._send_json(400, {"ok": False, "error": str(exc)})
        except Exception as exc:
            self._send_json(500, {"ok": False, "error": f"proxy failure: {exc}"})

    def _handle_nubox_ver_pdf(self) -> None:
        try:
            payload = self._read_json_body()
            row_id = str(payload.get("id", "")).strip()
            cookie_header = str(payload.get("cookieHeader", "")).strip()

            if not row_id:
                raise ValueError("missing id")
            if not cookie_header:
                raise ValueError("missing cookieHeader")

            token = extract_cookie_value(cookie_header, "COOKIE_SEGURA_SERVIPYME")
            if not token:
                raise ValueError("COOKIE_SEGURA_SERVIPYME not found in cookieHeader")

            post_headers = {
                "Content-Type": "application/json;charset=UTF-8",
                "Cookie": cookie_header,
                "X-Requested-With": "XMLHttpRequest",
                "Accept": "*/*",
                "Origin": "https://app.nubox.com",
                "Referer": "https://app.nubox.com/ServiFactura/paginas/dteDocumentosRecibidos.aspx",
            }
            post_body = {
                "token": token,
                "id": row_id,
            }

            post_status, post_content_type, post_bytes = self._perform_upstream_request(
                f"{NUBOX_BASE_URL}{NUBOX_PDF_PATH}",
                "POST",
                post_headers,
                post_body,
            )

            if post_status >= 400:
                self._send_bytes(post_status, post_content_type, post_bytes)
                return

            if "application/pdf" in post_content_type.lower() or post_bytes.startswith(b"%PDF"):
                self._send_bytes(200, "application/pdf", post_bytes)
                return

            parsed = safe_parse_json(post_bytes.decode("utf-8", errors="replace"))
            temp_path = find_temp_pdf_path(parsed) if isinstance(parsed, dict) else ""
            if not temp_path:
                self._send_json(502, {"ok": False, "error": "ver-pdf response did not include temp pdf path"})
                return

            normalized_temp = temp_path.strip()
            url_candidates = []
            if normalized_temp.startswith("//"):
                url_candidates.append(f"{NUBOX_BASE_URL}{normalized_temp}")
                url_candidates.append(f"{NUBOX_BASE_URL}/{normalized_temp.lstrip('/')}")
            elif normalized_temp.startswith("/"):
                # Nubox temp PDFs are commonly resolved with a double slash after domain.
                url_candidates.append(f"{NUBOX_BASE_URL}/{normalized_temp}")
                url_candidates.append(f"{NUBOX_BASE_URL}{normalized_temp}")
            else:
                url_candidates.append(f"{NUBOX_BASE_URL}//{normalized_temp}")
                url_candidates.append(f"{NUBOX_BASE_URL}/{normalized_temp}")

            get_headers = {
                "Cookie": cookie_header,
                "Accept": "application/pdf,*/*",
                "Referer": "https://app.nubox.com/ServiFactura/paginas/dteDocumentosRecibidos.aspx",
            }

            last_status = 502
            last_content_type = "application/json; charset=utf-8"
            last_bytes = b""

            for url in url_candidates:
                status, content_type, raw = self._perform_upstream_request(url, "GET", get_headers)
                last_status = status
                last_content_type = content_type
                last_bytes = raw

                if status < 400 and ("application/pdf" in content_type.lower() or raw.startswith(b"%PDF")):
                    self._send_bytes(200, "application/pdf", raw)
                    return

            self._send_bytes(last_status, last_content_type, last_bytes)
        except ValueError as exc:
            self._send_json(400, {"ok": False, "error": str(exc)})
        except Exception as exc:
            self._send_json(500, {"ok": False, "error": f"ver-pdf failure: {exc}"})

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self._set_cors()
        self.end_headers()

    def do_GET(self) -> None:
        # Keep static file serving for the frontend.
        return super().do_GET()

    def do_POST(self) -> None:
        if self.path == "/api/nubox/obtener-por-filtro":
            self._handle_nubox_proxy()
            return

        if self.path == "/api/nubox/ver-pdf":
            self._handle_nubox_ver_pdf()
            return

        self.send_response(404)
        self._set_cors()
        self.end_headers()


if __name__ == "__main__":
    print(f"Nubox proxy server: http://{HOST}:{PORT}")
    print("Abrir app: http://127.0.0.1:8765/app/index.html")

    server = ThreadingHTTPServer((HOST, PORT), NuboxProxyHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
