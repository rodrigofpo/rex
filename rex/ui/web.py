"""
Interface Web Local (Zero-Dependency) para o REX.
Funciona em qualquer navegador moderno sem necessidade de dependências do sistema operacional.
"""

import http.server
import socketserver
import json
import logging
import tempfile
import urllib.parse
from pathlib import Path
import threading
import webbrowser
import time
import os

from rex.workbook import WORKBOOK_MODE_CONSOLIDATED, WORKBOOK_MODES
from rex.errors import REXError
from rex.security import DocumentLimits

MAX_UPLOAD_BYTES = DocumentLimits().max_file_bytes
ALLOWED_HOSTS = {"127.0.0.1", "localhost"}
logger = logging.getLogger(__name__)


def validate_upload_filename(filename: object) -> str:
    """Aceita apenas um nome DOCX, sem componentes de caminho."""
    if (
        not isinstance(filename, str)
        or not filename
        or filename.startswith("~$")
        or "/" in filename
        or "\\" in filename
        or "\x00" in filename
        or Path(filename).suffix.lower() != ".docx"
    ):
        raise ValueError("Selecione um arquivo .docx válido.")
    return filename


def dataframe_records_for_json(dataframe, limit: int = 150) -> list[dict]:
    """Converte uma prévia do DataFrame para registros JSON sem NaN ou NaT."""
    preview = dataframe.head(limit).astype(object)
    preview = preview.where(preview.notna(), None)
    return preview.to_dict(orient="records")


HTML_PAGE = Path(__file__).with_name("web.html").read_text(encoding="utf-8")


class ExtractorWebHandler(http.server.SimpleHTTPRequestHandler):
    target_dir = Path.cwd()
    latest_excel_path: Path | None = None
    latest_csv_path: Path | None = None

    def log_message(self, format, *args):
        pass

    def end_headers(self):
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("X-Frame-Options", "DENY")
        super().end_headers()

    def _host_is_allowed(self) -> bool:
        host = self.headers.get("Host", "").partition(":")[0].lower()
        return host in ALLOWED_HOSTS

    def _send_json(self, status: int, payload: dict | list):
        body = json.dumps(payload, allow_nan=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if not self._host_is_allowed():
            self.send_error(403, "Host não permitido")
            return

        parsed = urllib.parse.urlparse(self.path)

        if parsed.path in ("/", "/index.html"):
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode("utf-8"))

        elif parsed.path == "/api/download/excel":
            excel_path = self.latest_excel_path
            if excel_path is not None and excel_path.is_file():
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{excel_path.name}"')
                self.end_headers()
                self.wfile.write(excel_path.read_bytes())
            else:
                self.send_error(404, "Arquivo Excel não encontrado.")

        elif parsed.path == "/api/download/csv":
            csv_path = self.latest_csv_path
            if csv_path is not None and csv_path.is_file():
                self.send_response(200)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header("Content-Disposition", f'attachment; filename="{csv_path.name}"')
                self.end_headers()
                self.wfile.write(csv_path.read_bytes())
            else:
                self.send_error(404, "Arquivo CSV não encontrado.")

        else:
            self.send_error(404, "Não encontrado")

    def do_POST(self):
        if not self._host_is_allowed():
            self.send_error(403, "Host não permitido")
            return

        if self.path == "/api/extract-upload":
            self._handle_upload()
            return
        self.send_error(404, "Não encontrado")

    def _handle_upload(self):
        try:
            if self.headers.get_content_type() != "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                raise ValueError("Content-Type deve ser um documento DOCX.")
            content_len = int(self.headers.get("Content-Length", 0))
            if content_len <= 0 or content_len > MAX_UPLOAD_BYTES:
                raise ValueError("O laudo deve ter até 250 MB e não pode estar vazio.")
            encoded_name = self.headers.get("X-REX-Filename", "")
            if len(encoded_name) > 1024:
                raise ValueError("Nome de arquivo muito longo.")
            filename = validate_upload_filename(urllib.parse.unquote(encoded_name))
            workbook_mode = self.headers.get("X-REX-Workbook-Mode", WORKBOOK_MODE_CONSOLIDATED)
            if workbook_mode not in WORKBOOK_MODES:
                raise ValueError("Modo de pasta de trabalho inválido.")

            with tempfile.TemporaryDirectory(prefix="rex-upload-") as temp_dir:
                document_path = Path(temp_dir) / filename
                with document_path.open("wb") as destination:
                    remaining = content_len
                    while remaining:
                        chunk = self.rfile.read(min(1024 * 1024, remaining))
                        if not chunk:
                            raise ValueError("O upload foi interrompido antes de terminar.")
                        destination.write(chunk)
                        remaining -= len(chunk)
                self._process_document(document_path, workbook_mode)
        except ValueError as exc:
            self._send_json(400, {"error": str(exc)})
        except OSError:
            logger.exception("Falha ao receber o laudo pela interface web")
            self._send_json(500, {"error": "Não foi possível receber o laudo."})

    def _process_document(self, docx_file: Path, workbook_mode: str):
        t0 = time.time()
        from rex.core.engine import EDSExtractorEngine

        engine = EDSExtractorEngine()
        try:
            df_dados, df_amostras = engine.process(
                docx_file,
                output_dir=self.target_dir,
                workbook_mode=workbook_mode,
            )
            type(self).latest_excel_path = engine.last_export_paths.excel
            type(self).latest_csv_path = engine.last_export_paths.csv
            elapsed = round(time.time() - t0, 1)

            preview_records = dataframe_records_for_json(df_dados)
            elements = sorted(df_dados["Elemento"].unique().tolist()) if not df_dados.empty else []
            points_count = int(df_dados["Ponto"].nunique()) if not df_dados.empty else 0

            response = {
                "samples_count": len(df_amostras),
                "rows_count": len(df_dados),
                "points_count": points_count,
                "elapsed": elapsed,
                "excel_filename": engine.last_export_paths.excel.name,
                "csv_filename": engine.last_export_paths.csv.name,
                "elements": elements,
                "preview": preview_records
            }
            self._send_json(200, response)

        except REXError as exc:
            self._send_json(422, {"error": str(exc)})
        except Exception:
            logger.exception("Falha inesperada durante a extração web")
            self._send_json(500, {"error": "Falha inesperada. Consulte o log do servidor."})


def launch_web_app(port: int = 8085, open_browser: bool = True, target_dir: Path | str = None):
    """Inicia o servidor web local do REX."""
    if target_dir:
        ExtractorWebHandler.target_dir = Path(target_dir).resolve()
    else:
        ExtractorWebHandler.target_dir = Path.cwd().resolve()
    ExtractorWebHandler.latest_excel_path = None
    ExtractorWebHandler.latest_csv_path = None

    print("=" * 65)
    print(f"  REX — INTERFACE WEB LOCAL")
    print("=" * 65)
    print(f"Servidor disponível em: http://localhost:{port}")
    print(f"Diretório de trabalho:  {ExtractorWebHandler.target_dir}")
    print("Pressione Ctrl+C para encerrar o servidor.")
    print("-" * 65)

    if open_browser:
        threading.Thread(target=lambda: (time.sleep(0.8), webbrowser.open(f"http://localhost:{port}")), daemon=True).start()

    with socketserver.TCPServer(("127.0.0.1", port), ExtractorWebHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor web do REX encerrado.")
