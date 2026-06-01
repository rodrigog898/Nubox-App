from __future__ import annotations

import io
import json
import re
import zipfile
from datetime import datetime
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib import error as urlerror
from urllib import parse as urlparse
from urllib import request as urlrequest

try:
    from openpyxl import Workbook
except Exception:
    Workbook = None

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

BASE_DIR = Path(__file__).resolve().parent.parent
HOST = "127.0.0.1"
PORT = 8765
NUBOX_BASE_URL = "https://app.nubox.com"
NUBOX_PDF_PATH = "/ServiFactura/paginas/dteDocumentosRecibidos.aspx/VerPDFDocManualesElectronicos"

RUT_REGEX = re.compile(r"\b(?:\d{1,2}\.\d{3}\.\d{3}|\d{7,8})-[\dkK]\b")
FOLIO_REGEX = re.compile(r"N\s*[°ºo]?\s*[:#]?\s*(\d{1,12})", re.IGNORECASE)
FECHA_REGEX = re.compile(r"Fecha\s*:\s*([0-3]?\d[-/][01]?\d[-/]\d{4})", re.IGNORECASE)
NUMERIC_REGEX = re.compile(r"\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+,\d+|\d+")
ORDER_REGEX = re.compile(r"Orden\s+de\s+Compra\s+(\S+)(?:\s+([0-3]?\d[-/][01]?\d[-/]\d{4}))?", re.IGNORECASE)
AF_ITEM_REGEX = re.compile(
    r"(?P<cantidad>\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s+(?P<precio>\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s+AF\s+(?P<total>\d{1,3}(?:\.\d{3})*(?:,\d+)?)",
    re.IGNORECASE,
)
AF_ITEM_TAIL_REGEX = re.compile(
    r"^(?P<descripcion>.*?)\s+(?P<cantidad>\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s+(?P<precio>\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s+AF\s+(?P<total>\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s*$",
    re.IGNORECASE,
)
GENERIC_ITEM_TAIL_REGEX = re.compile(
    r"^(?P<descripcion>.*?)\s+(?P<precio>\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s+(?P<total>\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s*$",
    re.IGNORECASE,
)
LEGAL_SUFFIX_REGEX = re.compile(r"\b(LTDA\.?|LIMITADA|S\.A\.?|SPA|EIRL|S A)\b", re.IGNORECASE)


def _normalize_lines(text: str) -> list[str]:
    lines = []
    for raw in text.replace("\r", "\n").split("\n"):
        normalized = re.sub(r"\s+", " ", raw).strip()
        if normalized:
            lines.append(normalized)
    return lines


def _extract_invoice_items(text: str) -> list[dict]:
    lines = _normalize_lines(text)
    upper_lines = [line.upper() for line in lines]

    start_idx = -1
    for idx, line in enumerate(upper_lines):
        if "ITEM" in line and "PRECIO" in line and "TOTAL" in line:
            start_idx = idx
            break

    if start_idx < 0:
        return []

    end_idx = len(lines)
    for idx in range(start_idx + 1, len(upper_lines)):
        if any(marker in upper_lines[idx] for marker in ("IMPUESTOS", "REFERENCIAS", "COMISIONES", "TOTALES")):
            end_idx = idx
            break

    item_lines = lines[start_idx + 1 : end_idx]
    if not item_lines:
        return []

    chunks: list[list[str]] = []
    current: list[str] = []
    for line in item_lines:
        if re.match(r"^\d+(?:\s+|[A-Za-z])", line):
            if current:
                chunks.append(current)
            current = [line]
        elif current:
            current.append(line)
    if current:
        chunks.append(current)

    if not chunks:
        chunks = [item_lines]

    parsed_items = []
    for chunk in chunks:
        first_line = chunk[0]
        first_match = re.match(r"^(\d{1,3})\s+(.*)$", first_line)
        if not first_match:
            first_match = re.match(r"^(\d{1,3})([A-Za-z].*)$", first_line)
        item_number = first_match.group(1) if first_match else ""
        desc_start = first_match.group(2) if first_match else first_line

        all_desc_lines = [desc_start] + chunk[1:]
        all_text = " ".join(all_desc_lines)
        all_text = re.sub(r"\s+", " ", all_text).strip()

        description = all_text
        unit_price = ""
        total_price = ""

        af_tail_match = AF_ITEM_TAIL_REGEX.match(all_text)
        if af_tail_match:
            description = af_tail_match.group("descripcion").strip(" -")
            unit_price = af_tail_match.group("precio")
            total_price = af_tail_match.group("total")
        else:
            af_match = AF_ITEM_REGEX.search(all_text)
            if af_match:
                unit_price = af_match.group("precio")
                total_price = af_match.group("total")
                description = all_text[: af_match.start()].strip(" -")
            else:
                generic_tail_match = GENERIC_ITEM_TAIL_REGEX.match(all_text)
                if generic_tail_match:
                    description = generic_tail_match.group("descripcion").strip(" -")
                    unit_price = generic_tail_match.group("precio")
                    total_price = generic_tail_match.group("total")

        description = re.sub(r"\s+", " ", description).strip(" -")

        parsed_items.append(
            {
                "item_numero": item_number,
                "item_descripcion": description,
                "precio_unitario": unit_price,
                "total_item": total_price,
            }
        )

    return parsed_items


def _extract_razon_social(lines: list[str]) -> str:
    if not lines:
        return ""

    folio_idx = -1
    for idx, line in enumerate(lines):
        if FOLIO_REGEX.search(line):
            folio_idx = idx
            break

    if folio_idx >= 0:
        start_idx = folio_idx + 1
    else:
        factura_idx = -1
        for idx, line in enumerate(lines):
            if "FACTURA" in line.upper():
                factura_idx = idx
                break
        start_idx = factura_idx + 1 if factura_idx >= 0 else 0

    stop_markers = (
        "SEÑOR(ES)",
        "SENOR(ES)",
        "DATOS DE PAGO",
        "DATOS DE DESPACHO",
        "ITEM",
        "RUT:",
        "DIRECCIÓN",
        "DIRECCION",
    )
    address_markers = ("AVDA", "AVENIDA", "CALLE", "CAMINO", "PASAJE", "NRO", "N°")

    candidates: list[str] = []
    for line in lines[start_idx : start_idx + 12]:
        upper_line = line.upper()
        if any(marker in upper_line for marker in stop_markers):
            break

        if any(marker in upper_line for marker in address_markers) and re.search(r"\d", upper_line):
            break

        if re.search(r"\bGIRO\b", upper_line):
            break

        if re.search(r"\b(CONSTRUCCION|SERVICIO|DISTRIBUIDORA|ASESORIA EN)\b", upper_line) and candidates:
            # When we already captured legal name, this usually starts business activity lines.
            break

        clean = line.strip(" -")
        if not clean:
            continue

        suffix_match = LEGAL_SUFFIX_REGEX.search(clean)
        if suffix_match:
            # Keep only the legal name fragment when OCR joins giro/address text on same line.
            clean = clean[: suffix_match.end()].strip(" -")

        candidates.append(clean)
        if suffix_match:
            break

    if not candidates:
        return ""

    if len(candidates) > 2:
        candidates = candidates[:2]

    return " ".join(candidates)


def _extract_invoice_data(pdf_bytes: bytes, row_id: str, pdf_filename: str) -> list[dict]:
    if PdfReader is None:
        raise RuntimeError("Falta dependencia 'pypdf'. Instala con: pip install pypdf")

    reader = PdfReader(io.BytesIO(pdf_bytes))
    text = "\n".join((page.extract_text() or "") for page in reader.pages)
    lines = _normalize_lines(text)
    joined_text = "\n".join(lines)

    ruts = RUT_REGEX.findall(joined_text)
    folio_match = FOLIO_REGEX.search(joined_text)
    fecha_match = FECHA_REGEX.search(joined_text)
    order_match = ORDER_REGEX.search(joined_text)

    rut_emisor = ruts[0] if len(ruts) >= 1 else ""
    razon_social = _extract_razon_social(lines)
    folio = folio_match.group(1) if folio_match else ""
    fecha = fecha_match.group(1) if fecha_match else ""
    orden_compra = order_match.group(1) if order_match else ""
    fecha_orden_compra = order_match.group(2) if order_match and order_match.group(2) else ""

    items = _extract_invoice_items(joined_text)
    if not items:
        items = [
            {
                "item_numero": "",
                "item_descripcion": "",
                "precio_unitario": "",
                "total_item": "",
            }
        ]

    rows = []
    for item in items:
        rows.append(
            {
                "rut_emisor": rut_emisor,
                "razon_social": razon_social,
                "folio": folio,
                "fecha": fecha,
                "orden_compra": orden_compra,
                "fecha_orden_compra": fecha_orden_compra,
                "item_numero": item["item_numero"],
                "item_descripcion": item["item_descripcion"],
                "precio_unitario": item["precio_unitario"],
                "total_item": item["total_item"],
            }
        )

    return rows


def _build_excel_bytes(rows: list[dict]) -> bytes:
    if Workbook is None:
        raise RuntimeError("Falta dependencia 'openpyxl'. Instala con: pip install openpyxl")

    headers = [
        "rut_emisor",
        "razon_social",
        "folio",
        "fecha",
        "orden_compra",
        "fecha_orden_compra",
        "item_numero",
        "item_descripcion",
        "precio_unitario",
        "total_item",
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Facturas"
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header, "") for header in headers])

    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            value = str(row.get(header, "") or "")
            if len(value) > max_len:
                max_len = len(value)
        ws.column_dimensions[chr(64 + idx)].width = min(max_len + 2, 70)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


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

    def _send_bytes(self, status: int, content_type: str, raw: bytes, extra_headers: dict | None = None) -> None:
        self.send_response(status)
        self._set_cors()
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(raw)))
        if extra_headers:
            for key, value in extra_headers.items():
                self.send_header(str(key), str(value))
        self.end_headers()
        self.wfile.write(raw)

    def _read_json_body(self) -> dict:
        content_length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(content_length)
        return json.loads(raw.decode("utf-8"))

    def _read_binary_body(self) -> bytes:
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            return b""
        return self.rfile.read(content_length)

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

    def _fetch_pdf_bytes_for_id(self, row_id: str, cookie_header: str) -> bytes:
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
            raise ValueError(f"upstream POST error {post_status}: {post_bytes[:300].decode('utf-8', errors='replace')}")

        if "application/pdf" in post_content_type.lower() or post_bytes.startswith(b"%PDF"):
            return post_bytes

        parsed = safe_parse_json(post_bytes.decode("utf-8", errors="replace"))
        temp_path = find_temp_pdf_path(parsed) if isinstance(parsed, dict) else ""
        if not temp_path:
            raise ValueError("ver-pdf response did not include temp pdf path")

        normalized_temp = temp_path.strip()
        url_candidates = []
        if normalized_temp.startswith("//"):
            url_candidates.append(f"{NUBOX_BASE_URL}{normalized_temp}")
            url_candidates.append(f"{NUBOX_BASE_URL}/{normalized_temp.lstrip('/')}")
        elif normalized_temp.startswith("/"):
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
                return raw

        raise ValueError(
            f"temp pdf fetch failed {last_status} ({last_content_type}): {last_bytes[:300].decode('utf-8', errors='replace')}"
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

            pdf_bytes = self._fetch_pdf_bytes_for_id(row_id, cookie_header)
            self._send_bytes(200, "application/pdf", pdf_bytes)
        except ValueError as exc:
            self._send_json(400, {"ok": False, "error": str(exc)})
        except Exception as exc:
            self._send_json(500, {"ok": False, "error": f"ver-pdf failure: {exc}"})

    def _handle_download_zip_analyzed(self) -> None:
        try:
            payload = self._read_json_body()
            cookie_header = str(payload.get("cookieHeader", "")).strip()
            ids_raw = payload.get("ids", [])

            if not cookie_header:
                raise ValueError("missing cookieHeader")
            if not isinstance(ids_raw, list):
                raise ValueError("ids must be an array")

            ids: list[str] = []
            for value in ids_raw:
                current = str(value).strip()
                if current:
                    ids.append(current)

            if not ids:
                raise ValueError("ids is empty")

            unique_ids = list(dict.fromkeys(ids))
            all_rows: list[dict] = []
            failed_messages: list[str] = []

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                for row_id in unique_ids:
                    pdf_filename = f"pdf/documento_{row_id}.pdf"
                    try:
                        pdf_bytes = self._fetch_pdf_bytes_for_id(row_id, cookie_header)
                        zf.writestr(pdf_filename, pdf_bytes)
                        all_rows.extend(_extract_invoice_data(pdf_bytes, row_id, pdf_filename))
                    except Exception as exc:
                        error_text = str(exc)
                        failed_messages.append(f"ID {row_id}: {error_text}")

                excel_bytes = _build_excel_bytes(all_rows)
                zf.writestr("analisis_facturas.xlsx", excel_bytes)

                if failed_messages:
                    zf.writestr("errores_descarga.txt", "\n".join(failed_messages).encode("utf-8"))

            zip_raw = zip_buffer.getvalue()
            timestamp = datetime.now().strftime("%Y-%m-%d")
            filename = f"pdf_nubox_analizado_{timestamp}.zip"
            self._send_bytes(
                200,
                "application/zip",
                zip_raw,
                {
                    "Content-Disposition": f'attachment; filename="{filename}"',
                    "X-Total-Ids": str(len(unique_ids)),
                    "X-Failed-Ids": str(len(failed_messages)),
                },
            )
        except ValueError as exc:
            self._send_json(400, {"ok": False, "error": str(exc)})
        except Exception as exc:
            self._send_json(500, {"ok": False, "error": f"zip-analyzed failure: {exc}"})

    def _handle_analyze_uploaded_zip(self) -> None:
        try:
            zip_raw = self._read_binary_body()
            if not zip_raw:
                raise ValueError("missing zip body")

            rows: list[dict] = []
            failed_files: list[str] = []

            with zipfile.ZipFile(io.BytesIO(zip_raw), mode="r") as zf:
                pdf_names = [name for name in zf.namelist() if name.lower().endswith(".pdf")]
                if not pdf_names:
                    raise ValueError("zip does not contain PDF files")

                for pdf_name in pdf_names:
                    try:
                        pdf_bytes = zf.read(pdf_name)
                        row_id = Path(pdf_name).stem
                        rows.extend(_extract_invoice_data(pdf_bytes, row_id, pdf_name))
                    except Exception as exc:
                        failed_files.append(f"{pdf_name}: {exc}")

            self._send_json(
                200,
                {
                    "ok": True,
                    "rows": rows,
                    "totalRows": len(rows),
                    "failedFiles": failed_files,
                    "failedCount": len(failed_files),
                },
            )
        except ValueError as exc:
            self._send_json(400, {"ok": False, "error": str(exc)})
        except zipfile.BadZipFile:
            self._send_json(400, {"ok": False, "error": "invalid zip file"})
        except Exception as exc:
            self._send_json(500, {"ok": False, "error": f"zip-upload analyze failure: {exc}"})

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

        if self.path == "/api/nubox/descargar-zip-analizado":
            self._handle_download_zip_analyzed()
            return

        if self.path == "/api/zip/analizar":
            self._handle_analyze_uploaded_zip()
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
