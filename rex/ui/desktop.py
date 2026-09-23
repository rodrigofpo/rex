"""
Interface Gráfica Desktop em CustomTkinter para o REX.
"""

import os
import sys
import time
import platform
import subprocess
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

import customtkinter as ctk
import pandas as pd

from rex.core.engine import EDSExtractorEngine
from rex.workbook import WORKBOOK_MODE_CONSOLIDATED, WORKBOOK_MODE_PER_SAMPLE

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

COLOR_ACCENT = "#7c6af7"
COLOR_ACCENT_HOVER = "#9d8fff"
COLOR_BG_CARD = "#212130"

WORKBOOK_MODE_OPTIONS = {
    "Consolidada — todos os dados em uma planilha": WORKBOOK_MODE_CONSOLIDATED,
    "Por amostra — uma planilha para cada amostra": WORKBOOK_MODE_PER_SAMPLE,
}


def open_file_or_folder_in_os(path: Path | str):
    p = Path(path)
    if not p.exists():
        return
    current_os = platform.system()
    try:
        if current_os == "Windows":
            os.startfile(str(p))
        elif current_os == "Darwin":
            subprocess.run(["open", str(p)], check=True)
        else:
            subprocess.run(["xdg-open", str(p)], check=True)
    except Exception as e:
        print(f"Erro ao abrir arquivo/pasta: {e}")


class REXDesktopApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("REX — MEV-EDS Report Extractor")
        self.geometry("980x760")
        self.minsize(860, 680)

        self.docx_path_var = ctk.StringVar(value="")
        self.out_dir_var = ctk.StringVar(value="")
        self.status_var = ctk.StringVar(value="Selecione um laudo .docx para iniciar.")
        self.workbook_mode_var = ctk.StringVar(value=next(iter(WORKBOOK_MODE_OPTIONS)))

        self.last_excel_path: Path | None = None
        self.last_csv_path: Path | None = None
        self.df_dados: pd.DataFrame | None = None
        self.df_amostras: pd.DataFrame | None = None

        self._build_ui()
        self._apply_treeview_style()

    def _build_ui(self):
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=28, pady=(20, 10))

        title_lbl = ctk.CTkLabel(
            header_frame,
            text="REX — MEV-EDS Report Extractor",
            font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
            text_color="#ffffff"
        )
        title_lbl.pack(anchor="w")

        subtitle_lbl = ctk.CTkLabel(
            header_frame,
            text="Pipeline Inteligente Multiplataforma • RapidOCR + OpenXML em Memória",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#a0a0b8"
        )
        subtitle_lbl.pack(anchor="w")

        card_files = ctk.CTkFrame(self, fg_color=COLOR_BG_CARD, corner_radius=12)
        card_files.pack(fill="x", padx=28, pady=10)

        f_lbl = ctk.CTkLabel(card_files, text="Laudo Microscopia (.docx):", font=ctk.CTkFont(size=12, weight="bold"))
        f_lbl.grid(row=0, column=0, padx=(16, 8), pady=(14, 4), sticky="w")

        self.f_entry = ctk.CTkEntry(
            card_files,
            textvariable=self.docx_path_var,
            placeholder_text="Clique em 'Procurar' para selecionar o laudo .docx...",
            font=ctk.CTkFont(size=12),
            height=36
        )
        self.f_entry.grid(row=1, column=0, padx=(16, 8), pady=(0, 10), sticky="ew")

        btn_browse = ctk.CTkButton(
            card_files,
            text="Procurar...",
            width=110,
            height=36,
            command=self._select_file,
            font=ctk.CTkFont(weight="bold")
        )
        btn_browse.grid(row=1, column=1, padx=(0, 16), pady=(0, 10))

        d_lbl = ctk.CTkLabel(card_files, text="Pasta de Saída dos Resultados:", font=ctk.CTkFont(size=12, weight="bold"))
        d_lbl.grid(row=2, column=0, padx=(16, 8), pady=(4, 4), sticky="w")

        self.d_entry = ctk.CTkEntry(
            card_files,
            textvariable=self.out_dir_var,
            placeholder_text="Pasta onde o Excel e CSV serão gerados...",
            font=ctk.CTkFont(size=12),
            height=36
        )
        self.d_entry.grid(row=3, column=0, padx=(16, 8), pady=(0, 10), sticky="ew")

        btn_browse_dir = ctk.CTkButton(
            card_files,
            text="Alterar...",
            width=110,
            height=36,
            command=self._select_output_dir,
            fg_color="#3a3a4d",
            hover_color="#4d4d66"
        )
        btn_browse_dir.grid(row=3, column=1, padx=(0, 16), pady=(0, 10))

        mode_lbl = ctk.CTkLabel(
            card_files,
            text="Estrutura da Pasta de Trabalho Excel:",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        mode_lbl.grid(row=4, column=0, padx=(16, 8), pady=(4, 4), sticky="w")

        self.workbook_mode_menu = ctk.CTkOptionMenu(
            card_files,
            variable=self.workbook_mode_var,
            values=list(WORKBOOK_MODE_OPTIONS),
            height=36,
        )
        self.workbook_mode_menu.grid(
            row=5, column=0, columnspan=2, padx=16, pady=(0, 16), sticky="ew"
        )

        card_files.columnconfigure(0, weight=1)

        action_frame = ctk.CTkFrame(self, fg_color="transparent")
        action_frame.pack(fill="x", padx=28, pady=(4, 10))

        self.btn_run = ctk.CTkButton(
            action_frame,
            text="▶  Iniciar Extração Completa",
            height=44,
            command=self._start_extraction,
            fg_color=COLOR_ACCENT,
            hover_color=COLOR_ACCENT_HOVER,
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=8
        )
        self.btn_run.pack(fill="x")

        self.progress_bar = ctk.CTkProgressBar(action_frame, height=6)
        self.progress_bar.pack(fill="x", pady=(10, 4))
        self.progress_bar.set(0)

        self.status_lbl = ctk.CTkLabel(
            action_frame,
            textvariable=self.status_var,
            font=ctk.CTkFont(size=11),
            text_color="#9e9eb0"
        )
        self.status_lbl.pack(anchor="w")

        self.tabview = ctk.CTkTabview(self, corner_radius=12, fg_color=COLOR_BG_CARD)
        self.tabview.pack(fill="both", expand=True, padx=28, pady=(6, 12))

        self.tab_log = self.tabview.add("📋 Log em Tempo Real")
        self.tab_preview = self.tabview.add("📊 Prévia dos Dados")
        self.tab_metrics = self.tabview.add("🔬 Resumo Científico")

        self.log_box = ctk.CTkTextbox(
            self.tab_log,
            font=ctk.CTkFont(family="Consolas, Menlo, Monaco, monospace", size=12),
            wrap="word",
            activate_scrollbars=True
        )
        self.log_box.pack(fill="both", expand=True, padx=10, pady=10)

        self._build_preview_tab()
        self._build_metrics_tab()

        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.pack(fill="x", padx=28, pady=(0, 16))

        self.btn_open_excel = ctk.CTkButton(
            bottom_frame,
            text="📊 Abrir Planilha Excel",
            command=self._open_excel,
            fg_color="#2e7d32",
            hover_color="#388e3c",
            state="disabled",
            height=36,
            font=ctk.CTkFont(weight="bold")
        )
        self.btn_open_excel.pack(side="left", padx=(0, 10))

        self.btn_open_folder = ctk.CTkButton(
            bottom_frame,
            text="📁 Abrir Pasta de Destino",
            command=self._open_folder,
            fg_color="#3a3a4d",
            hover_color="#4d4d66",
            state="disabled",
            height=36
        )
        self.btn_open_folder.pack(side="left")

    def _build_preview_tab(self):
        container = ctk.CTkFrame(self.tab_preview, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=8, pady=8)

        cols = ("Amostra", "Micrografia", "Serie", "Ponto", "Elemento", "TipoLinha", "% Peso", "Sigma", "% Atomica")
        self.tree = ttk.Treeview(container, columns=cols, show="headings", selectmode="browse")

        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        col_widths = {
            "Amostra": 180, "Micrografia": 120, "Serie": 50, "Ponto": 50,
            "Elemento": 70, "TipoLinha": 90, "% Peso": 70, "Sigma": 60, "% Atomica": 70
        }
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths.get(col, 80), anchor="center")

        self.tree.column("Amostra", anchor="w")
        self.tree.column("Micrografia", anchor="w")

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

    def _build_metrics_tab(self):
        m_frame = ctk.CTkFrame(self.tab_metrics, fg_color="transparent")
        m_frame.pack(fill="both", expand=True, padx=16, pady=16)

        grid_cards = ctk.CTkFrame(m_frame, fg_color="transparent")
        grid_cards.pack(fill="x", pady=(0, 16))
        for i in range(4):
            grid_cards.columnconfigure(i, weight=1)

        self.card_samples = self._make_stat_card(grid_cards, 0, "Amostras", "0", "Áreas identificadas")
        self.card_points = self._make_stat_card(grid_cards, 1, "Pontos Analisados", "0", "Localizações de EDS")
        self.card_rows = self._make_stat_card(grid_cards, 2, "Linhas Químicas", "0", "Medições de elementos")
        self.card_time = self._make_stat_card(grid_cards, 3, "Tempo Gasto", "0.0s", "Duração da extração")

        detail_card = ctk.CTkFrame(m_frame, fg_color="#181824", corner_radius=10)
        detail_card.pack(fill="both", expand=True)

        ctk.CTkLabel(
            detail_card,
            text="Composição de Elementos Identificados:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=16, pady=(14, 6))

        self.lbl_elements = ctk.CTkLabel(
            detail_card,
            text="Nenhum dado processado ainda.",
            font=ctk.CTkFont(family="Consolas, monospace", size=13),
            text_color="#c0c0d8",
            wraplength=700,
            justify="left"
        )
        self.lbl_elements.pack(anchor="w", padx=16, pady=(0, 14))

    def _make_stat_card(self, parent, col, title, value, subtitle):
        card = ctk.CTkFrame(parent, fg_color="#181824", corner_radius=10)
        card.grid(row=0, column=col, padx=6, pady=4, sticky="ew")

        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=11), text_color="#a0a0b8").pack(anchor="w", padx=12, pady=(10, 0))
        val_lbl = ctk.CTkLabel(card, text=value, font=ctk.CTkFont(size=20, weight="bold"), text_color=COLOR_ACCENT)
        val_lbl.pack(anchor="w", padx=12, pady=(2, 0))
        ctk.CTkLabel(card, text=subtitle, font=ctk.CTkFont(size=10), text_color="#707085").pack(anchor="w", padx=12, pady=(0, 10))
        return val_lbl

    def _apply_treeview_style(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background="#1e1e2c",
            foreground="#e0e0ee",
            fieldbackground="#1e1e2c",
            rowheight=26,
            font=("Segoe UI", 10),
            borderwidth=0
        )
        style.configure(
            "Treeview.Heading",
            background="#2a2a3e",
            foreground="#ffffff",
            font=("Segoe UI", 10, "bold"),
            borderwidth=0,
            relief="flat"
        )
        style.map("Treeview", background=[("selected", COLOR_ACCENT)], foreground=[("selected", "#ffffff")])
        style.map("Treeview.Heading", background=[("active", "#3a3a55")])

    def _log(self, text: str):
        def _append():
            self.log_box.insert("end", f"[{time.strftime('%H:%M:%S')}] {text}\n")
            self.log_box.see("end")
        self.after(0, _append)

    def _select_file(self):
        path = filedialog.askopenfilename(
            title="Selecione o laudo .docx",
            filetypes=[("Documentos Word", "*.docx"), ("Todos os arquivos", "*.*")]
        )
        if path:
            self.docx_path_var.set(path)
            p = Path(path)
            self.out_dir_var.set(str(p.parent))
            self.status_var.set(f"Arquivo selecionado: {p.name}")
            self._log(f"Arquivo selecionado: {p.name} ({p.stat().st_size / (1024*1024):.1f} MB)")

    def _select_output_dir(self):
        path = filedialog.askdirectory(title="Selecione a pasta de destino")
        if path:
            self.out_dir_var.set(path)
            self._log(f"Pasta de saída alterada para: {path}")

    def _start_extraction(self):
        docx = self.docx_path_var.get().strip()
        if not docx or not os.path.isfile(docx):
            messagebox.showwarning("Atenção", "Por favor, selecione um arquivo .docx válido.")
            return

        out_dir = self.out_dir_var.get().strip() or str(Path(docx).parent)
        workbook_mode = WORKBOOK_MODE_OPTIONS[self.workbook_mode_var.get()]

        self.btn_run.configure(state="disabled", text="⏳  Extraindo Dados (Aguarde)...")
        self.btn_open_excel.configure(state="disabled")
        self.btn_open_folder.configure(state="disabled")
        self.progress_bar.set(0)
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()

        self.log_box.delete("1.0", "end")
        self.status_var.set("Executando extração com RapidOCR em memória...")

        threading.Thread(
            target=self._worker,
            args=(docx, out_dir, workbook_mode),
            daemon=True,
        ).start()

    def _worker(self, docx_path: str, out_dir: str, workbook_mode: str):
        t0 = time.time()
        engine = EDSExtractorEngine()

        try:
            self._log("Iniciando varredura determinística linear e OCR...")
            df_dados, df_amostras = engine.process(
                docx_path=docx_path,
                output_dir=out_dir,
                workbook_mode=workbook_mode,
                status_callback=lambda msg: self._log(msg)
            )

            elapsed = time.time() - t0
            self.df_dados = df_dados
            self.df_amostras = df_amostras
            self.last_excel_path = Path(out_dir) / "amostras_organizadas_completas.xlsx"
            self.last_csv_path = Path(out_dir) / "dados_extraidos.csv"

            self.after(0, self._on_extraction_success, elapsed)

        except Exception as e:
            self.after(0, self._on_extraction_error, str(e))

    def _on_extraction_success(self, elapsed: float):
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(1.0)
        self.btn_run.configure(state="normal", text="▶  Iniciar Extração Completa")
        self.btn_open_excel.configure(state="normal")
        self.btn_open_folder.configure(state="normal")

        n_samples = len(self.df_amostras) if self.df_amostras is not None else 0
        n_rows = len(self.df_dados) if self.df_dados is not None else 0
        n_points = self.df_dados['Ponto'].nunique() if self.df_dados is not None and not self.df_dados.empty else 0

        self.status_var.set(f"Concluído com sucesso em {elapsed:.1f}s! ({n_samples} amostras, {n_rows} medições)")
        self._log(f"✔ Sucesso! {n_samples} amostras e {n_rows} linhas químicas.")

        self.card_samples.configure(text=str(n_samples))
        self.card_points.configure(text=str(n_points))
        self.card_rows.configure(text=str(n_rows))
        self.card_time.configure(text=f"{elapsed:.1f}s")

        if self.df_dados is not None and not self.df_dados.empty:
            elems = sorted(self.df_dados['Elemento'].unique().tolist())
            self.lbl_elements.configure(text=" • ".join(elems))

            for item in self.tree.get_children():
                self.tree.delete(item)

            for _, row in self.df_dados.head(300).iterrows():
                self.tree.insert("", "end", values=(
                    row["Amostra"],
                    row["Micrografia"],
                    row["Serie"],
                    row["Ponto"],
                    row["Elemento"],
                    row["TipoLinha"],
                    f"{row['PercentualPeso']:.2f}" if pd.notnull(row['PercentualPeso']) else "",
                    f"{row['SigmaPeso']:.2f}" if pd.notnull(row['SigmaPeso']) else "",
                    f"{row['PercentualAtomico']:.2f}" if pd.notnull(row['PercentualAtomico']) else ""
                ))

        self.tabview.set("📊 Prévia dos Dados")
        messagebox.showinfo("Extração Concluída", f"Foram extraídas com sucesso {n_samples} amostras ({n_rows} medições)!\n\nArquivos salvos em:\n{self.last_excel_path}")

    def _on_extraction_error(self, err_msg: str):
        self.progress_bar.stop()
        self.progress_bar.set(0)
        self.btn_run.configure(state="normal", text="▶  Iniciar Extração Completa")
        self.status_var.set("Erro durante o processamento. Veja o log.")
        self._log(f"✘ ERRO: {err_msg}")
        messagebox.showerror("Erro na Extração", f"Ocorreu uma falha ao processar o arquivo:\n\n{err_msg}")

    def _open_excel(self):
        if self.last_excel_path and self.last_excel_path.exists():
            open_file_or_folder_in_os(self.last_excel_path)

    def _open_folder(self):
        folder = self.out_dir_var.get().strip()
        if folder and Path(folder).exists():
            open_file_or_folder_in_os(folder)


def launch_desktop_app():
    app = REXDesktopApp()
    app.mainloop()


if __name__ == "__main__":
    launch_desktop_app()
