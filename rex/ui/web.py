"""
Interface Web Local (Zero-Dependency) para o REX.
Funciona em qualquer navegador moderno sem necessidade de dependências do sistema operacional.
"""

import http.server
import socketserver
import json
import urllib.parse
from pathlib import Path
import threading
import webbrowser
import time
import os

MAX_REQUEST_BODY_BYTES = 16 * 1024
ALLOWED_HOSTS = {"127.0.0.1", "localhost"}


def resolve_docx_path(target_dir: Path | str, filename: object) -> Path:
    """Resolve um DOCX simples, sem permitir caminhos fora do diretório publicado."""
    if not isinstance(filename, str) or not filename:
        raise ValueError("Nome de arquivo inválido.")

    requested = Path(filename)
    if requested.name != filename or requested.suffix.lower() != ".docx":
        raise ValueError("Selecione um arquivo .docx do diretório de trabalho.")

    base_dir = Path(target_dir).resolve()
    candidate = (base_dir / requested).resolve()
    if candidate.parent != base_dir or not candidate.is_file():
        raise ValueError("Arquivo .docx não encontrado no diretório de trabalho.")

    return candidate


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
    select, input[type="text"] {
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
    select:focus, input[type="text"]:focus {
      border-color: var(--accent);
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
        <label>Selecione o Laudo Word (.docx):</label>
        <select id="docxSelect"></select>
      </div>
      <button id="btnRun" class="btn-run" onclick="startExtraction()">
        <span>▶</span> <span>Iniciar Extração Completa</span>
      </button>
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
        <button class="tab-btn active" onclick="switchTab('log')">📋 Log em Tempo Real</button>
        <button class="tab-btn" onclick="switchTab('preview')">📊 Prévia dos Dados</button>
        <button class="tab-btn" onclick="switchTab('elements')">🧪 Elementos Químicos</button>
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
    async function loadFiles() {
      const res = await fetch('/api/files');
      const files = await res.json();
      const sel = document.getElementById('docxSelect');
      sel.innerHTML = '';
      if (files.length === 0) {
        const opt = document.createElement('option');
        opt.value = '';
        opt.textContent = 'Nenhum arquivo .docx encontrado no diretório atual';
        sel.appendChild(opt);
        return;
      }
      files.forEach(f => {
        const opt = document.createElement('option');
        opt.value = f;
        opt.textContent = f;
        sel.appendChild(opt);
      });
    }

    function switchTab(tabId) {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      event.target.classList.add('active');
      document.getElementById('tab-' + tabId).classList.add('active');
    }

    function appendLog(text) {
      const box = document.getElementById('logBox');
      const time = new Date().toLocaleTimeString();
      box.textContent += `[${time}] ${text}\\n`;
      box.scrollTop = box.scrollHeight;
    }

    async function startExtraction() {
      const docx = document.getElementById('docxSelect').value;
      if (!docx) return alert('Selecione um arquivo .docx primeiro.');

      const btn = document.getElementById('btnRun');
      btn.disabled = true;
      btn.innerHTML = '<span class="loader"></span> <span>Extraindo Dados (RapidOCR em execução)...</span>';

      const logBox = document.getElementById('logBox');
      logBox.textContent = '';
      appendLog(`Iniciando extração do arquivo: ${docx}`);

      try {
        const res = await fetch('/api/extract', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ filename: docx })
        });
        const data = await res.json();

        if (data.error) {
          appendLog(`ERRO: ${data.error}`);
          alert('Erro durante a extração: ' + data.error);
        } else {
          appendLog(`Sucesso! ${data.samples_count} amostras e ${data.rows_count} medições extraídas.`);

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
      } finally {
        btn.disabled = false;
        btn.innerHTML = '<span>▶</span> <span>Iniciar Extração Completa</span>';
      }
    }

    loadFiles();
  </script>
</body>
</html>
"""


class ExtractorWebHandler(http.server.SimpleHTTPRequestHandler):
    target_dir = Path.cwd()

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

        elif parsed.path == "/api/files":
            docx_files = [f.name for f in self.target_dir.glob("*.docx") if not f.name.startswith("~$")]
            self._send_json(200, sorted(docx_files))

        elif parsed.path == "/api/download/excel":
            excel_path = self.target_dir / "amostras_organizadas_completas.xlsx"
            if excel_path.is_file():
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{excel_path.name}"')
                self.end_headers()
                self.wfile.write(excel_path.read_bytes())
            else:
                self.send_error(404, "Arquivo Excel não encontrado.")

        elif parsed.path == "/api/download/csv":
            csv_path = self.target_dir / "dados_extraidos.csv"
            if csv_path.is_file():
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

        if self.path != "/api/extract":
            self.send_error(404, "Não encontrado")
            return

        try:
            content_type = self.headers.get_content_type()
            if content_type != "application/json":
                raise ValueError("Content-Type deve ser application/json.")
            content_len = int(self.headers.get("Content-Length", 0))
            if content_len <= 0 or content_len > MAX_REQUEST_BODY_BYTES:
                raise ValueError("Tamanho de requisição inválido.")
            post_body = self.rfile.read(content_len)
            data = json.loads(post_body.decode("utf-8"))
            if not isinstance(data, dict):
                raise ValueError("Corpo JSON inválido.")
            docx_file = resolve_docx_path(self.target_dir, data.get("filename"))
        except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as exc:
            self._send_json(400, {"error": str(exc)})
            return

        t0 = time.time()
        from rex.core.engine import EDSExtractorEngine

        engine = EDSExtractorEngine()
        try:
            df_dados, df_amostras = engine.process(docx_file, output_dir=self.target_dir)
            elapsed = round(time.time() - t0, 1)

            preview_records = dataframe_records_for_json(df_dados)
            elements = sorted(df_dados["Elemento"].unique().tolist()) if not df_dados.empty else []
            points_count = int(df_dados["Ponto"].nunique()) if not df_dados.empty else 0

            response = {
                "samples_count": len(df_amostras),
                "rows_count": len(df_dados),
                "points_count": points_count,
                "elapsed": elapsed,
                "elements": elements,
                "preview": preview_records
            }
            self._send_json(200, response)

        except Exception as exc:
            self._send_json(500, {"error": str(exc)})


def launch_web_app(port: int = 8085, open_browser: bool = True, target_dir: Path | str = None):
    """Inicia o servidor web local do REX."""
    if target_dir:
        ExtractorWebHandler.target_dir = Path(target_dir).resolve()
    else:
        ExtractorWebHandler.target_dir = Path.cwd().resolve()

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
