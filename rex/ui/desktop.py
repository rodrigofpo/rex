"""Interface desktop do REX integrada aos módulos de extração e exportação."""

from __future__ import annotations

import os
import platform
import queue
import subprocess
import threading
import time
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Literal

import customtkinter as ctk
import pandas as pd

from rex.core.engine import EDSExtractorEngine
from rex.errors import ExportError, InvalidDocumentError, OCRInitializationError, REXError
from rex.exporter import export_result
from rex.models import ExportPaths, ExtractionResult
from rex.workbook import WORKBOOK_MODE_CONSOLIDATED, WORKBOOK_MODE_PER_SAMPLE


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

MAX_DOCX_BYTES = 250 * 1024 * 1024

COLORS = {
    "page": "#F4F7F8",
    "surface": "#FFFFFF",
    "surface_muted": "#EDF2F3",
    "border": "#C9D5D8",
    "text": "#14272B",
    "muted": "#52676C",
    "primary": "#0B666A",
    "primary_hover": "#084F52",
    "primary_soft": "#DCEEEE",
    "success": "#176B45",
    "success_soft": "#E3F3EA",
    "danger": "#A83232",
    "danger_soft": "#FBE8E8",
    "disabled": "#AAB7BA",
}

WORKBOOK_MODE_OPTIONS = {
    "Consolidado": WORKBOOK_MODE_CONSOLIDATED,
    "Uma aba por amostra": WORKBOOK_MODE_PER_SAMPLE,
}

PREVIEW_COLUMNS = (
    "Amostra",
    "Micrografia",
    "Serie",
    "Ponto",
    "Elemento",
    "TipoLinha",
    "PercentualPeso",
    "SigmaPeso",
    "PercentualAtomico",
)


def count_eds_points(data: pd.DataFrame) -> int:
    """Conta pontos por amostra, micrografia e série, sem fundir números repetidos."""
    keys = ("Amostra", "Micrografia", "Serie", "Ponto")
    if data.empty or any(key not in data.columns for key in keys):
        return 0
    return int(data.loc[:, keys].drop_duplicates().shape[0])


@dataclass(frozen=True)
class ExtractionSuccess:
    result: ExtractionResult
    paths: ExportPaths
    elapsed: float


UiEventKind = Literal["status", "success", "error"]
UiEvent = tuple[UiEventKind, object]


def calculate_window_geometry(
    screen_width: int,
    screen_height: int,
    scale_factor: float = 1.0,
) -> tuple[int, int, int, int]:
    """Calcula tamanho lógico CTk e posição física centralizada na tela."""
    logical_width = max(1, int(screen_width / scale_factor))
    logical_height = max(1, int(screen_height / scale_factor))
    usable_width = min(logical_width, max(640, int(logical_width * 0.90)))
    usable_height = min(logical_height, max(600, int(logical_height * 0.88)))
    window_width = min(1200, usable_width)
    window_height = min(820, usable_height)
    offset_x = max(0, (screen_width - round(window_width * scale_factor)) // 2)
    offset_y = max(0, (screen_height - round(window_height * scale_factor)) // 2)
    return window_width, window_height, offset_x, offset_y


def linux_scale_from_tk(tk_scaling: float) -> float:
    """Converte pixels por ponto do Tk para o fator relativo a 96 DPI."""
    return max(0.4, tk_scaling / (96 / 72))


def open_file_or_folder_in_os(path: Path | str) -> None:
    """Abre um resultado no aplicativo padrão do sistema operacional."""
    target = Path(path)
    if not target.exists():
        raise FileNotFoundError(f"O caminho não existe: {target}")

    current_os = platform.system()
    if current_os == "Windows":
        os.startfile(str(target))  # type: ignore[attr-defined]
    elif current_os == "Darwin":
        subprocess.run(["open", str(target)], check=True)
    else:
        subprocess.run(["xdg-open", str(target)], check=True)


class REXDesktopApp(ctk.CTk):
    """Janela principal do fluxo local do REX."""

    def __init__(self) -> None:
        super().__init__()

        if platform.system() == "Linux":
            linux_scale = linux_scale_from_tk(float(self.tk.call("tk", "scaling")))
            ctk.set_widget_scaling(linux_scale)
            ctk.set_window_scaling(linux_scale)
        self.ui_scale = self._get_window_scaling()

        self.title("REX — MEV-EDS Report Extractor")
        width, height, offset_x, offset_y = calculate_window_geometry(
            self.winfo_screenwidth(),
            self.winfo_screenheight(),
            self.ui_scale,
        )
        self.geometry(f"{width}x{height}+{offset_x}+{offset_y}")
        self.minsize(min(820, width), min(620, height))
        self.configure(fg_color=COLORS["page"])

        self.docx_path: Path | None = None
        self.output_dir: Path | None = None
        self.last_paths: ExportPaths | None = None
        self.current_result: ExtractionResult | None = None
        self.is_processing = False
        self.events: queue.Queue[UiEvent] = queue.Queue()

        self.workbook_mode_var = ctk.StringVar(value="Consolidado")
        self.file_name_var = ctk.StringVar(value="Nenhum laudo selecionado")
        self.file_meta_var = ctk.StringVar(
            value="Localize um arquivo .docx não vazio de até 250 MB."
        )
        self.status_var = ctk.StringVar(value="Aguardando seleção do laudo")
        self.elements_var = ctk.StringVar(value="Nenhum resultado disponível.")

        self._build_ui()
        self._apply_treeview_style()
        self._set_step_two_enabled(False)
        self.after(100, self._drain_events)
        self.after(0, self._maximize_on_desktop)

    def _maximize_on_desktop(self) -> None:
        try:
            if platform.system() == "Windows":
                self.state("zoomed")
            elif platform.system() == "Linux":
                self.attributes("-zoomed", True)
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        shell = ctk.CTkScrollableFrame(
            self,
            fg_color="transparent",
            corner_radius=0,
            scrollbar_button_color=COLORS["border"],
            scrollbar_button_hover_color=COLORS["muted"],
        )
        shell.pack(fill="both", expand=True)
        self.scroll_shell = shell
        if platform.system() == "Linux":
            # CTk escuta apenas Button-4/5 no Linux. Alguns mouses e touchpads
            # sob Wayland/XWayland enviam MouseWheel em vez desses eventos.
            self.bind_all("<MouseWheel>", self._scroll_linux_mousewheel, add=True)

        content = ctk.CTkFrame(shell, fg_color="transparent")
        content.pack(fill="x", expand=True, padx=24, pady=(18, 24))
        content.grid_columnconfigure(0, weight=1)

        self._build_header(content)
        self._build_file_step(content)
        self._build_configuration_step(content)
        self._build_processing_status(content)
        self._build_results(content)

    def _scroll_linux_mousewheel(self, event: tk.Event) -> None:
        if not event.delta or not self.scroll_shell._check_if_valid_scroll(event.widget):
            return
        if self.scroll_shell._parent_canvas.yview() == (0.0, 1.0):
            return
        steps = max(1, round(abs(event.delta) / 120))
        self.scroll_shell._parent_canvas.yview_scroll(
            -steps if event.delta > 0 else steps, "units"
        )

    def _build_header(self, parent: ctk.CTkFrame) -> None:
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="REX",
            text_color=COLORS["primary"],
            font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"),
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            header,
            text="Extraia dados de laudos MEV-EDS e gere XLSX e CSV localmente.",
            text_color=COLORS["muted"],
            font=ctk.CTkFont(family="Segoe UI", size=14),
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(2, 0))

        local_badge = ctk.CTkLabel(
            header,
            text="PROCESSAMENTO LOCAL",
            fg_color=COLORS["primary_soft"],
            text_color=COLORS["primary"],
            corner_radius=6,
            padx=10,
            pady=5,
            font=ctk.CTkFont(size=11, weight="bold"),
        )
        local_badge.grid(row=0, column=1, rowspan=2, sticky="e", padx=(16, 0))

    def _section_card(self, parent: ctk.CTkFrame, row: int) -> ctk.CTkFrame:
        card = ctk.CTkFrame(
            parent,
            fg_color=COLORS["surface"],
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=10,
        )
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.grid_columnconfigure(1, weight=1)
        return card

    def _step_heading(
        self,
        card: ctk.CTkFrame,
        number: str,
        title: str,
        helper: str,
    ) -> None:
        badge = ctk.CTkLabel(
            card,
            text=number,
            width=32,
            height=32,
            corner_radius=16,
            fg_color=COLORS["primary"],
            text_color="#FFFFFF",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        badge.grid(row=0, column=0, rowspan=2, padx=(16, 12), pady=(14, 8), sticky="n")
        ctk.CTkLabel(
            card,
            text=title,
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=17, weight="bold"),
            anchor="w",
        ).grid(row=0, column=1, sticky="ew", padx=(0, 16), pady=(13, 0))
        ctk.CTkLabel(
            card,
            text=helper,
            text_color=COLORS["muted"],
            font=ctk.CTkFont(size=12),
            anchor="w",
            justify="left",
        ).grid(row=1, column=1, sticky="ew", padx=(0, 16), pady=(0, 8))

    def _build_file_step(self, parent: ctk.CTkFrame) -> None:
        card = self._section_card(parent, 1)
        self._step_heading(
            card,
            "1",
            "Escolha o laudo",
            "O arquivo permanece no computador e será processado localmente.",
        )

        selected = ctk.CTkFrame(card, fg_color=COLORS["surface_muted"], corner_radius=8)
        selected.grid(row=2, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 10))
        selected.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            selected,
            textvariable=self.file_name_var,
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=12, pady=(9, 0))
        ctk.CTkLabel(
            selected,
            textvariable=self.file_meta_var,
            text_color=COLORS["muted"],
            font=ctk.CTkFont(size=12),
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 9))

        self.btn_select_file = ctk.CTkButton(
            card,
            text="Localizar arquivo DOCX",
            command=self._select_file,
            height=42,
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            font=ctk.CTkFont(size=13, weight="bold"),
            corner_radius=8,
        )
        self.btn_select_file.grid(
            row=3,
            column=0,
            columnspan=2,
            sticky="ew",
            padx=16,
            pady=(0, 14),
        )

    def _build_configuration_step(self, parent: ctk.CTkFrame) -> None:
        self.step_two_card = self._section_card(parent, 2)
        self._step_heading(
            self.step_two_card,
            "2",
            "Organize o XLSX",
            "A opção escolhida afeta somente o Excel. O CSV será sempre consolidado.",
        )

        self.lock_label = ctk.CTkLabel(
            self.step_two_card,
            text="Selecione um DOCX válido na etapa 1 para liberar esta configuração.",
            fg_color=COLORS["surface_muted"],
            text_color=COLORS["muted"],
            corner_radius=8,
            anchor="w",
            padx=12,
            pady=9,
            font=ctk.CTkFont(size=12),
        )
        self.lock_label.grid(
            row=2,
            column=0,
            columnspan=2,
            sticky="ew",
            padx=16,
            pady=(0, 8),
        )

        choices = ctk.CTkFrame(self.step_two_card, fg_color="transparent")
        choices.grid(row=3, column=0, columnspan=2, sticky="ew", padx=16)
        choices.grid_columnconfigure((0, 1), weight=1)

        self.mode_consolidated = ctk.CTkRadioButton(
            choices,
            text="Consolidado\nTodos os dados em uma aba",
            variable=self.workbook_mode_var,
            value="Consolidado",
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            border_color=COLORS["muted"],
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=12),
        )
        self.mode_consolidated.grid(row=0, column=0, sticky="w", pady=6)

        self.mode_per_sample = ctk.CTkRadioButton(
            choices,
            text="Uma aba por amostra\nResumo e abas separadas",
            variable=self.workbook_mode_var,
            value="Uma aba por amostra",
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            border_color=COLORS["muted"],
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=12),
        )
        self.mode_per_sample.grid(row=0, column=1, sticky="w", padx=(12, 0), pady=6)

        self.btn_extract = ctk.CTkButton(
            self.step_two_card,
            text="Extrair dados",
            command=self._start_extraction,
            height=44,
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=8,
        )
        self.btn_extract.grid(
            row=4,
            column=0,
            columnspan=2,
            sticky="ew",
            padx=16,
            pady=(8, 14),
        )

    def _build_processing_status(self, parent: ctk.CTkFrame) -> None:
        self.status_card = ctk.CTkFrame(
            parent,
            fg_color=COLORS["surface"],
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=10,
        )
        self.status_card.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        self.status_card.grid_columnconfigure(0, weight=1)

        self.status_title = ctk.CTkLabel(
            self.status_card,
            text="Estado da extração",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        )
        self.status_title.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 0))
        self.status_label = ctk.CTkLabel(
            self.status_card,
            textvariable=self.status_var,
            text_color=COLORS["muted"],
            font=ctk.CTkFont(size=12),
            anchor="w",
            justify="left",
            wraplength=900,
        )
        self.status_label.grid(row=1, column=0, sticky="ew", padx=16, pady=(2, 8))

        self.progress = ctk.CTkProgressBar(
            self.status_card,
            height=6,
            fg_color=COLORS["surface_muted"],
            progress_color=COLORS["primary"],
        )
        self.progress.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 12))
        self.progress.set(0)

    def _build_results(self, parent: ctk.CTkFrame) -> None:
        self.results_card = ctk.CTkFrame(
            parent,
            fg_color=COLORS["surface"],
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=10,
        )
        self.results_card.grid(row=4, column=0, sticky="nsew")
        self.results_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.results_card,
            text="Resultados",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=17, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=16, pady=(13, 8))

        metrics = ctk.CTkFrame(self.results_card, fg_color="transparent")
        metrics.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))
        metrics.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.metric_samples = self._metric(metrics, 0, "Amostras")
        self.metric_points = self._metric(metrics, 1, "Pontos EDS")
        self.metric_rows = self._metric(metrics, 2, "Medições")
        self.metric_time = self._metric(metrics, 3, "Tempo")

        ctk.CTkLabel(
            self.results_card,
            text="Elementos encontrados",
            text_color=COLORS["muted"],
            font=ctk.CTkFont(size=11, weight="bold"),
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", padx=16)
        ctk.CTkLabel(
            self.results_card,
            textvariable=self.elements_var,
            text_color=COLORS["text"],
            font=ctk.CTkFont(family="Consolas", size=12),
            anchor="w",
            justify="left",
            wraplength=900,
        ).grid(row=3, column=0, sticky="ew", padx=16, pady=(2, 8))

        tabs = ctk.CTkTabview(
            self.results_card,
            height=286,
            fg_color=COLORS["surface_muted"],
            segmented_button_selected_color=COLORS["primary"],
            segmented_button_selected_hover_color=COLORS["primary_hover"],
        )
        tabs.grid(row=4, column=0, sticky="nsew", padx=16, pady=(0, 10))
        preview_tab = tabs.add("Prévia tabular")
        history_tab = tabs.add("Detalhes do processamento")

        table_frame = ctk.CTkFrame(preview_tab, fg_color="transparent")
        table_frame.pack(fill="both", expand=True, padx=6, pady=6)
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            table_frame,
            columns=PREVIEW_COLUMNS,
            show="headings",
            selectmode="browse",
            height=9,
        )
        vertical = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)

        column_widths = {
            "Amostra": 170,
            "Micrografia": 110,
            "Serie": 55,
            "Ponto": 55,
            "Elemento": 75,
            "TipoLinha": 90,
            "PercentualPeso": 105,
            "SigmaPeso": 85,
            "PercentualAtomico": 120,
        }
        headings = {
            "PercentualPeso": "% peso",
            "SigmaPeso": "Sigma",
            "PercentualAtomico": "% atômica",
        }
        for column in PREVIEW_COLUMNS:
            self.tree.heading(column, text=headings.get(column, column))
            self.tree.column(
                column,
                width=round(column_widths[column] * self.ui_scale),
                minwidth=round(48 * self.ui_scale),
                anchor="w" if column in {"Amostra", "Micrografia"} else "center",
            )

        self.tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")

        self.history_box = ctk.CTkTextbox(
            history_tab,
            fg_color=COLORS["surface"],
            text_color=COLORS["text"],
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
        )
        self.history_box.pack(fill="both", expand=True, padx=6, pady=6)
        self.history_box.insert("end", "Aguardando o início da extração.\n")
        self.history_box.configure(state="disabled")

        downloads = ctk.CTkFrame(self.results_card, fg_color="transparent")
        downloads.grid(row=5, column=0, sticky="ew", padx=16, pady=(0, 14))
        downloads.grid_columnconfigure((0, 1, 2), weight=1)

        self.btn_open_xlsx = self._result_button(
            downloads,
            0,
            "Abrir XLSX",
            self._open_xlsx,
            COLORS["success"],
            "#105437",
        )
        self.btn_open_csv = self._result_button(
            downloads,
            1,
            "Abrir CSV",
            self._open_csv,
            COLORS["primary"],
            COLORS["primary_hover"],
        )
        self.btn_open_folder = self._result_button(
            downloads,
            2,
            "Abrir pasta",
            self._open_folder,
            "#435B61",
            "#31464B",
        )
        self._set_result_actions_enabled(False)

    def _metric(self, parent: ctk.CTkFrame, column: int, caption: str) -> ctk.CTkLabel:
        frame = ctk.CTkFrame(parent, fg_color=COLORS["surface_muted"], corner_radius=8)
        frame.grid(row=0, column=column, sticky="ew", padx=4)
        value = ctk.CTkLabel(
            frame,
            text="—",
            text_color=COLORS["primary"],
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        value.pack(anchor="w", padx=10, pady=(8, 0))
        ctk.CTkLabel(
            frame,
            text=caption,
            text_color=COLORS["muted"],
            font=ctk.CTkFont(size=11),
        ).pack(anchor="w", padx=10, pady=(0, 8))
        return value

    def _result_button(
        self,
        parent: ctk.CTkFrame,
        column: int,
        label: str,
        command,
        color: str,
        hover: str,
    ) -> ctk.CTkButton:
        button = ctk.CTkButton(
            parent,
            text=label,
            command=command,
            height=42,
            fg_color=color,
            hover_color=hover,
            font=ctk.CTkFont(size=12, weight="bold"),
            corner_radius=8,
        )
        button.grid(row=0, column=column, sticky="ew", padx=4)
        return button

    def _apply_treeview_style(self) -> None:
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background="#FFFFFF",
            foreground=COLORS["text"],
            fieldbackground="#FFFFFF",
            rowheight=round(27 * self.ui_scale),
            font=("Segoe UI", 10),
            borderwidth=0,
        )
        style.configure(
            "Treeview.Heading",
            background="#DCE6E8",
            foreground=COLORS["text"],
            font=("Segoe UI", 10, "bold"),
            borderwidth=0,
            relief="flat",
        )
        style.map(
            "Treeview",
            background=[("selected", COLORS["primary"])],
            foreground=[("selected", "#FFFFFF")],
        )

    def _select_file(self) -> None:
        selected_path = filedialog.askopenfilename(
            title="Localizar laudo DOCX",
            filetypes=[("Documento Word", "*.docx")],
        )
        if not selected_path:
            return

        path = Path(selected_path)
        validation_error = self._validate_file_selection(path)
        if validation_error:
            self._show_file_error(validation_error)
            return

        self.docx_path = path
        self.output_dir = path.parent
        self.last_paths = None
        self.current_result = None
        self.file_name_var.set(path.name)
        self.file_meta_var.set(
            f"{path.stat().st_size / (1024 * 1024):.1f} MB · {path.parent}"
        )
        self.status_var.set("Laudo selecionado. Configure o XLSX na etapa 2.")
        self.status_card.configure(border_color=COLORS["border"])
        self.status_label.configure(text_color=COLORS["muted"])
        self.btn_select_file.configure(text="Trocar arquivo DOCX")
        self._set_step_two_enabled(True)
        self._reset_results()

    @staticmethod
    def _validate_file_selection(path: Path) -> str | None:
        if path.suffix.lower() != ".docx":
            return "Selecione um arquivo com extensão .docx."
        if not path.is_file():
            return "O arquivo selecionado não está mais disponível."
        try:
            size = path.stat().st_size
        except OSError as exc:
            return f"Não foi possível ler o arquivo: {exc}"
        if size == 0:
            return "O DOCX está vazio. Escolha um laudo que contenha dados."
        if size > MAX_DOCX_BYTES:
            return "O DOCX ultrapassa o limite de 250 MB."
        return None

    def _show_file_error(self, message: str) -> None:
        self.docx_path = None
        self.output_dir = None
        self.last_paths = None
        self.current_result = None
        self._reset_results()
        self.file_name_var.set("Arquivo não aceito")
        self.file_meta_var.set(message)
        self.status_var.set(f"Erro de arquivo: {message}")
        self.status_card.configure(border_color=COLORS["danger"])
        self.status_label.configure(text_color=COLORS["danger"])
        self._set_step_two_enabled(False)
        messagebox.showerror("Arquivo não aceito", message)

    def _set_step_two_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled and not self.is_processing else "disabled"
        self.mode_consolidated.configure(state=state)
        self.mode_per_sample.configure(state=state)
        self.btn_extract.configure(state=state)
        if enabled:
            self.lock_label.configure(
                text="DOCX válido selecionado. Escolha a organização do XLSX.",
                fg_color=COLORS["success_soft"],
                text_color=COLORS["success"],
            )
        else:
            self.lock_label.configure(
                text="Selecione um DOCX válido na etapa 1 para liberar esta configuração.",
                fg_color=COLORS["surface_muted"],
                text_color=COLORS["muted"],
            )

    def _start_extraction(self) -> None:
        if self.is_processing:
            return
        if self.docx_path is None:
            self._show_file_error("Selecione um DOCX válido antes de continuar.")
            return

        validation_error = self._validate_file_selection(self.docx_path)
        if validation_error:
            self._show_file_error(validation_error)
            return

        self.output_dir = self.docx_path.parent
        self.is_processing = True
        self.last_paths = None
        self.current_result = None
        self._reset_results()
        self._clear_history()
        self._append_history("Extração iniciada. Aguarde a resposta do motor do REX.")

        self.btn_select_file.configure(state="disabled")
        self._set_step_two_enabled(True)
        self.btn_extract.configure(text="Processando…")
        self._set_result_actions_enabled(False)
        self.status_var.set("Processando o laudo. O REX não informa percentual de progresso.")
        self.status_card.configure(border_color=COLORS["primary"])
        self.status_label.configure(text_color=COLORS["primary"])
        self.progress.configure(mode="indeterminate")
        self.progress.start()

        workbook_mode = WORKBOOK_MODE_OPTIONS[self.workbook_mode_var.get()]
        worker = threading.Thread(
            target=self._extract_in_background,
            args=(self.docx_path, self.output_dir, workbook_mode),
            daemon=True,
            name="rex-extraction",
        )
        worker.start()

    def _extract_in_background(
        self,
        source_path: Path,
        output_dir: Path,
        workbook_mode: str,
    ) -> None:
        started_at = time.perf_counter()
        try:
            engine = EDSExtractorEngine()
            result = engine.extract(
                source_path,
                status_callback=lambda text: self.events.put(("status", str(text))),
            )
            self.events.put(("status", "Gerando arquivos XLSX e CSV…"))
            paths = export_result(
                result,
                output_dir=output_dir,
                source_path=source_path,
                workbook_mode=workbook_mode,
            )
            success = ExtractionSuccess(
                result=result,
                paths=paths,
                elapsed=time.perf_counter() - started_at,
            )
            self.events.put(("success", success))
        except (InvalidDocumentError, OCRInitializationError, ExportError, REXError) as exc:
            self.events.put(("error", exc))
        except Exception as exc:  # salvaguarda para falhas inesperadas da thread
            self.events.put(("error", exc))

    def _drain_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "status":
                    text = str(payload)
                    self.status_var.set(text)
                    self._append_history(text)
                elif kind == "success":
                    self._finish_success(payload)
                elif kind == "error":
                    self._finish_error(payload)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._drain_events)

    def _finish_success(self, payload: object) -> None:
        if not isinstance(payload, ExtractionSuccess):
            self._finish_error(RuntimeError("Resposta de extração inválida."))
            return

        self.is_processing = False
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(1)
        self.btn_select_file.configure(state="normal")
        self.btn_extract.configure(text="Extrair novamente")
        self._set_step_two_enabled(True)

        self.current_result = payload.result
        self.last_paths = payload.paths
        data = payload.result.data
        samples = payload.result.samples
        sample_count = len(samples)
        row_count = len(data)
        point_count = count_eds_points(data)

        self.metric_samples.configure(text=str(sample_count))
        self.metric_points.configure(text=str(point_count))
        self.metric_rows.configure(text=str(row_count))
        self.metric_time.configure(text=f"{payload.elapsed:.1f}s")

        elements = []
        if "Elemento" in data.columns:
            elements = sorted(str(value) for value in data["Elemento"].dropna().unique())
        self.elements_var.set(" · ".join(elements) if elements else "Nenhum elemento informado.")
        self._fill_preview(data)
        self._set_result_actions_enabled(True)

        self.status_var.set(
            f"Extração concluída: {sample_count} amostras, "
            f"{point_count} pontos EDS e {row_count} medições."
        )
        self.status_card.configure(border_color=COLORS["success"])
        self.status_label.configure(text_color=COLORS["success"])
        self._append_history(f"XLSX: {payload.paths.excel}")
        self._append_history(f"CSV: {payload.paths.csv}")

    def _finish_error(self, payload: object) -> None:
        self.is_processing = False
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(0)
        self.btn_select_file.configure(state="normal")
        self.btn_extract.configure(text="Tentar novamente")
        self._set_step_two_enabled(self.docx_path is not None)
        self._set_result_actions_enabled(False)

        message = self._friendly_error_message(payload)
        self.status_var.set(message)
        self.status_card.configure(border_color=COLORS["danger"])
        self.status_label.configure(text_color=COLORS["danger"])
        self._append_history(f"Falha: {message}")
        messagebox.showerror("Não foi possível concluir", message)

    @staticmethod
    def _friendly_error_message(error: object) -> str:
        if isinstance(error, OCRInitializationError):
            return f"O mecanismo OCR não pôde ser iniciado. {error}"
        if isinstance(error, InvalidDocumentError):
            return f"O DOCX não pôde ser processado. {error}"
        if isinstance(error, ExportError):
            return f"Os dados foram processados, mas os arquivos não puderam ser gravados. {error}"
        if isinstance(error, REXError):
            return str(error)
        return f"Ocorreu uma falha inesperada durante a extração: {error}"

    def _fill_preview(self, data: pd.DataFrame) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in data.head(300).iterrows():
            values = []
            for column in PREVIEW_COLUMNS:
                value = row[column] if column in data.columns else ""
                if pd.isna(value):
                    values.append("")
                elif column in {"PercentualPeso", "SigmaPeso", "PercentualAtomico"}:
                    try:
                        values.append(f"{float(value):.2f}")
                    except (TypeError, ValueError):
                        values.append(str(value))
                else:
                    values.append(value)
            self.tree.insert("", "end", values=values)

    def _reset_results(self) -> None:
        for label in (
            self.metric_samples,
            self.metric_points,
            self.metric_rows,
            self.metric_time,
        ):
            label.configure(text="—")
        self.elements_var.set("Nenhum resultado disponível.")
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._set_result_actions_enabled(False)

    def _set_result_actions_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.btn_open_xlsx.configure(state=state)
        self.btn_open_csv.configure(state=state)
        self.btn_open_folder.configure(state=state)

    def _clear_history(self) -> None:
        self.history_box.configure(state="normal")
        self.history_box.delete("1.0", "end")
        self.history_box.configure(state="disabled")

    def _append_history(self, text: str) -> None:
        self.history_box.configure(state="normal")
        self.history_box.insert("end", f"[{time.strftime('%H:%M:%S')}] {text}\n")
        self.history_box.see("end")
        self.history_box.configure(state="disabled")

    def _open_xlsx(self) -> None:
        if self.last_paths:
            self._open_path(self.last_paths.excel, "a planilha XLSX")

    def _open_csv(self) -> None:
        if self.last_paths:
            self._open_path(self.last_paths.csv, "o arquivo CSV")

    def _open_folder(self) -> None:
        if self.output_dir:
            self._open_path(self.output_dir, "a pasta de resultados")

    @staticmethod
    def _open_path(path: Path, description: str) -> None:
        try:
            open_file_or_folder_in_os(path)
        except (OSError, subprocess.SubprocessError) as exc:
            messagebox.showerror(
                "Não foi possível abrir",
                f"Não foi possível abrir {description}.\n\n{exc}",
            )


def launch_desktop_app() -> None:
    """Ponto de entrada mantido para ``rex gui``."""
    app = REXDesktopApp()
    app.mainloop()


if __name__ == "__main__":
    launch_desktop_app()
