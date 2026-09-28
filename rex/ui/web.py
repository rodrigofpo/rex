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


HTML_PAGE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <title>REX — MEV-EDS Report Extractor</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg: #0f0f17;
      --card-bg: #181824;
      --card-border: #28283c;
      --accent: #7c6af7;
      --accent-hover: #9686ff;
      --text: #f0f0f8;
      --text-dim: #9292a8;
      --success: #4caf90;
      --warn: #e5c07b;
      --danger: #e06c75;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      background: var(--bg);
      color: var(--text);
      font-family: 'Plus Jakarta Sans', sans-serif;
      padding: 32px 20px;
      min-height: 100vh;
    }
    .container {
      max-width: 1060px;
      margin: 0 auto;
    }
    header {
      margin-bottom: 28px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    h1 {
      font-size: 26px;
      font-weight: 700;
      letter-spacing: -0.5px;
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .badge {
      background: rgba(124, 106, 247, 0.18);
      color: var(--accent);
      border: 1px solid rgba(124, 106, 247, 0.3);
      padding: 4px 10px;
      border-radius: 999px;
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
    }
    p.subtitle {
      color: var(--text-dim);
      font-size: 14px;
      margin-top: 4px;
    }
    .card {
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 14px;
      padding: 24px;
      margin-bottom: 24px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.35);
    }
    .field-group {
      margin-bottom: 18px;
    }
    label {
      display: block;
      font-size: 12px;
      font-weight: 600;
      color: var(--text-dim);
      margin-bottom: 8px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    select, input[type="text"], input[type="file"] {
      width: 100%;
      background: #11111a;
      border: 1px solid var(--card-border);
      border-radius: 8px;
      padding: 12px 14px;
      color: var(--text);
      font-family: inherit;
      font-size: 14px;
      outline: none;
      transition: border-color 0.2s;
    }
    select:focus, input[type="text"]:focus, input[type="file"]:focus {
      border-color: var(--accent);
    }
    input[type="file"]::file-selector-button {
      background: var(--accent);
      color: #fff;
      border: 0;
      border-radius: 6px;
      padding: 7px 12px;
      margin-right: 12px;
      cursor: pointer;
    }
    .btn-run {
      background: var(--accent);
      color: white;
      border: none;
      border-radius: 10px;
      padding: 14px 28px;
      font-size: 15px;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      width: 100%;
      justify-content: center;
      transition: all 0.2s;
      box-shadow: 0 4px 14px rgba(124, 106, 247, 0.4);
    }
    .btn-run:hover {
      background: var(--accent-hover);
      transform: translateY(-1px);
    }
    .btn-run:disabled {
      opacity: 0.5;
      cursor: not-allowed;
      transform: none;
      box-shadow: none;
    }
    .grid-stats {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
      gap: 16px;
      margin-bottom: 24px;
    }
    .stat-card {
      background: #12121d;
      border: 1px solid var(--card-border);
      border-radius: 12px;
      padding: 18px;
    }
    .stat-title {
      font-size: 11px;
      color: var(--text-dim);
      text-transform: uppercase;
      font-weight: 600;
    }
    .stat-value {
      font-size: 28px;
      font-weight: 700;
      color: var(--accent);
      margin: 6px 0 2px 0;
    }
    .stat-desc {
      font-size: 12px;
      color: #6a6a82;
    }
    .tabs {
      display: flex;
      gap: 8px;
      border-bottom: 1px solid var(--card-border);
      margin-bottom: 18px;
    }
    .tab-btn {
      background: none;
      border: none;
      color: var(--text-dim);
      font-size: 13px;
      font-weight: 600;
      padding: 10px 18px;
      cursor: pointer;
      border-bottom: 2px solid transparent;
      transition: all 0.2s;
    }
    .tab-btn.active {
      color: var(--accent);
      border-bottom-color: var(--accent);
    }
    .tab-content { display: none; }
    .tab-content.active { display: block; }
    #logBox {
      background: #0d0d14;
      border: 1px solid var(--card-border);
      border-radius: 10px;
      padding: 16px;
      font-family: 'JetBrains Mono', monospace;
      font-size: 12px;
      line-height: 1.6;
      height: 320px;
      overflow-y: auto;
      color: #b5b5cb;
      white-space: pre-wrap;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    th, td {
      padding: 10px 12px;
      text-align: left;
      border-bottom: 1px solid #222232;
    }
    th {
      background: #11111a;
      color: var(--text-dim);
      font-weight: 600;
      text-transform: uppercase;
      font-size: 11px;
      letter-spacing: 0.5px;
    }
    tr:hover td {
      background: rgba(124, 106, 247, 0.05);
    }
    .table-container {
      max-height: 420px;
      overflow-y: auto;
      border-radius: 8px;
      border: 1px solid var(--card-border);
    }
    .actions-bar {
      display: flex;
      gap: 12px;
      margin-top: 18px;
    }
    .btn-action {
      background: #1f1f2e;
      border: 1px solid var(--card-border);
      color: var(--text);
      padding: 10px 20px;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      text-decoration: none;
      transition: background 0.2s;
    }
    .btn-action.success {
      background: rgba(76, 175, 144, 0.15);
      color: var(--success);
      border-color: rgba(76, 175, 144, 0.3);
    }
    .btn-action:hover {
      background: #2a2a3e;
    }
    .btn-action.success:hover {
      background: rgba(76, 175, 144, 0.25);
    }
    .hidden { display: none !important; }
    .feedback {
      margin-top: 14px;
      padding: 12px 14px;
      border-radius: 8px;
      border: 1px solid var(--card-border);
    }
    .feedback.error {
      color: var(--danger);
      border-color: var(--danger);
    }
    .feedback.success {
      color: var(--success);
      border-color: var(--success);
    }
    .loader {
      display: inline-block;
      width: 16px;
      height: 16px;
      border: 2px solid rgba(255,255,255,0.3);
      border-radius: 50%;
      border-top-color: #fff;
      animation: spin 0.8s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
  </style>
</head>
<body>
  <div class="container">
    <header>
      <div>
        <h1>🔬 REX <span class="badge">v1.0 • MEV-EDS Extractor</span></h1>
        <p class="subtitle">Processamento em memória & parsing determinístico linear de microscopia MEV-EDS</p>
      </div>
    </header>

    <div class="card">
      <div class="field-group">
        <label for="docxFile">Localize o laudo no computador (.docx):</label>
        <input id="docxFile" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document">
      </div>
      <div class="field-group">
        <label>Estrutura da Pasta de Trabalho Excel:</label>
        <select id="workbookMode">
          <option value="consolidated">Consolidada — todos os dados em uma planilha</option>
          <option value="per-sample">Por amostra — uma planilha para cada amostra</option>
        </select>
      </div>
      <button id="btnRun" class="btn-run" onclick="startExtraction()">
        <span>▶</span> <span>Iniciar Extração Completa</span>
      </button>
      <div id="feedback" class="feedback hidden" role="status" aria-live="polite"></div>
    </div>

    <div id="statsSection" class="grid-stats hidden">
      <div class="stat-card">
        <div class="stat-title">Amostras Detectadas</div>
        <div class="stat-value" id="statSamples">0</div>
        <div class="stat-desc">Micrografias com OCR</div>
      </div>
      <div class="stat-card">
        <div class="stat-title">Pontos de EDS</div>
        <div class="stat-value" id="statPoints">0</div>
        <div class="stat-desc">Regiões quantificadas</div>
      </div>
      <div class="stat-card">
        <div class="stat-title">Medições Químicas</div>
        <div class="stat-value" id="statRows">0</div>
        <div class="stat-desc">Linhas elementares extraídas</div>
      </div>
      <div class="stat-card">
        <div class="stat-title">Tempo de Processamento</div>
        <div class="stat-value" id="statTime">0.0s</div>
        <div class="stat-desc">100% executado em CPU</div>
      </div>
    </div>

    <div class="card">
      <div class="tabs">
        <button class="tab-btn active" onclick="switchTab('log', this)">📋 Log em Tempo Real</button>
        <button class="tab-btn" onclick="switchTab('preview', this)">📊 Prévia dos Dados</button>
        <button class="tab-btn" onclick="switchTab('elements', this)">🧪 Elementos Químicos</button>
      </div>

      <div id="tab-log" class="tab-content active">
        <div id="logBox">Aguardando início do processamento...</div>
      </div>

      <div id="tab-preview" class="tab-content">
        <div class="table-container">
          <table id="dataTable">
            <thead>
              <tr>
                <th>Amostra</th>
                <th>Micrografia</th>
                <th>Série</th>
                <th>Ponto</th>
                <th>Elemento</th>
                <th>Linha</th>
                <th>% Peso</th>
                <th>Sigma</th>
                <th>% Atômica</th>
              </tr>
            </thead>
            <tbody id="dataBody">
              <tr><td colspan="9" style="text-align: center; color: var(--text-dim); padding: 30px;">Nenhum dado processado ainda.</td></tr>
            </tbody>
          </table>
        </div>
      </div>

      <div id="tab-elements" class="tab-content">
        <p style="color: var(--text-dim); margin-bottom: 12px; font-size: 14px;">Elementos químicos detectados no documento:</p>
        <div id="elementsList" style="display: flex; flex-wrap: wrap; gap: 8px;">
          <span style="color: var(--text-dim);">Execute a extração para visualizar os elementos.</span>
        </div>
      </div>

      <div id="downloadBar" class="actions-bar hidden">
        <a href="/api/download/excel" class="btn-action success" download>
          <span>📥</span> Baixar Planilha Excel (.xlsx)
        </a>
        <a href="/api/download/csv" class="btn-action" download>
          <span>📄</span> Baixar Dados em CSV
        </a>
      </div>
    </div>
  </div>

  <script>
    function switchTab(tabId, button) {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      button.classList.add('active');
      document.getElementById('tab-' + tabId).classList.add('active');
    }

    function appendLog(text) {
      const box = document.getElementById('logBox');
      const time = new Date().toLocaleTimeString();
      box.textContent += `[${time}] ${text}\\n`;
      box.scrollTop = box.scrollHeight;
    }

    function showFeedback(message, kind) {
      const feedback = document.getElementById('feedback');
      feedback.textContent = message;
      feedback.className = 'feedback ' + kind;
    }

    async function startExtraction() {
      const file = document.getElementById('docxFile').files[0];
      const workbookMode = document.getElementById('workbookMode').value;
      if (!file) {
        showFeedback('Selecione um arquivo .docx primeiro.', 'error');
        return;
      }
      if (!file.name.toLowerCase().endsWith('.docx') || file.size === 0 || file.size > 250 * 1024 * 1024) {
        showFeedback('Escolha um DOCX não vazio com até 250 MB.', 'error');
        return;
      }

      const btn = document.getElementById('btnRun');
      btn.disabled = true;
      document.getElementById('feedback').className = 'feedback hidden';
      document.getElementById('downloadBar').classList.add('hidden');
      btn.innerHTML = '<span class="loader"></span> <span>Extraindo Dados (RapidOCR em execução)...</span>';

      const logBox = document.getElementById('logBox');
      logBox.textContent = '';
      appendLog(`Iniciando extração do arquivo: ${file.name}`);

      try {
        const res = await fetch('/api/extract-upload', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'X-REX-Filename': encodeURIComponent(file.name),
            'X-REX-Workbook-Mode': workbookMode
          },
          body: file
        });
        const data = await res.json();

        if (!res.ok || data.error) {
          const message = data.error || 'Não foi possível concluir a extração.';
          appendLog(`ERRO: ${message}`);
          showFeedback(message, 'error');
        } else {
          appendLog(`Sucesso! ${data.samples_count} amostras e ${data.rows_count} medições extraídas.`);
          showFeedback(`Arquivos gerados: ${data.excel_filename} e ${data.csv_filename}`, 'success');

          document.getElementById('statSamples').textContent = data.samples_count;
          document.getElementById('statPoints').textContent = data.points_count;
          document.getElementById('statRows').textContent = data.rows_count;
          document.getElementById('statTime').textContent = data.elapsed + 's';
          document.getElementById('statsSection').classList.remove('hidden');
          document.getElementById('downloadBar').classList.remove('hidden');

          const tbody = document.getElementById('dataBody');
          tbody.innerHTML = '';
          data.preview.forEach(row => {
            const tr = document.createElement('tr');
            const values = [
              row.Amostra, row.Micrografia, row.Serie, row.Ponto,
              row.Elemento, row.TipoLinha, row.PercentualPeso,
              row.SigmaPeso, row.PercentualAtomico
            ];
            values.forEach((value, index) => {
              const td = document.createElement('td');
              const content = index === 0 ? document.createElement('strong')
                            : index === 4 ? document.createElement('span')
                            : td;
              if (index === 4) content.className = 'badge';
              content.textContent = value ?? '';
              if (content !== td) td.appendChild(content);
              tr.appendChild(td);
            });
            tbody.appendChild(tr);
          });

          const elDiv = document.getElementById('elementsList');
          elDiv.innerHTML = '';
          data.elements.forEach(el => {
            const pill = document.createElement('span');
            pill.className = 'badge';
            pill.style.fontSize = '13px';
            pill.textContent = el;
            elDiv.appendChild(pill);
          });

          document.querySelectorAll('.tab-btn')[1].click();
        }
      } catch (err) {
        appendLog(`Erro de conexão: ${err.message}`);
        showFeedback('Não foi possível conectar ao servidor local. Verifique se ele continua em execução.', 'error');
      } finally {
        btn.disabled = false;
        btn.innerHTML = '<span>▶</span> <span>Iniciar Extração Completa</span>';
      }
    }

  </script>
</body>
</html>
"""


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
