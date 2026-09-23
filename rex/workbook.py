"""Geração das pastas de trabalho Excel produzidas pelo REX."""

import re
from pathlib import Path

import pandas as pd

WORKBOOK_MODE_CONSOLIDATED = "consolidated"
WORKBOOK_MODE_PER_SAMPLE = "per-sample"
WORKBOOK_MODES = (WORKBOOK_MODE_CONSOLIDATED, WORKBOOK_MODE_PER_SAMPLE)

_INVALID_SHEET_CHARACTERS = re.compile(r"[\\/*?:\[\]]")
_MAX_SHEET_NAME_LENGTH = 31


def validate_workbook_mode(workbook_mode: str) -> str:
    if workbook_mode not in WORKBOOK_MODES:
        available = ", ".join(WORKBOOK_MODES)
        raise ValueError(f"Modo de pasta de trabalho inválido. Opções: {available}.")
    return workbook_mode


def unique_sheet_name(label: object, used_names: set[str]) -> str:
    """Cria um nome de aba válido, curto e único para o Excel."""
    base_name = _INVALID_SHEET_CHARACTERS.sub("_", str(label)).strip().strip("'")
    base_name = base_name or "Amostra"
    base_name = base_name[:_MAX_SHEET_NAME_LENGTH]

    candidate = base_name
    counter = 2
    while candidate.casefold() in used_names:
        suffix = f"_{counter}"
        candidate = f"{base_name[:_MAX_SHEET_NAME_LENGTH - len(suffix)]}{suffix}"
        counter += 1

    used_names.add(candidate.casefold())
    return candidate


def write_excel_workbook(
    df_dados: pd.DataFrame,
    df_resumo: pd.DataFrame,
    xlsx_path: str | Path,
    workbook_mode: str = WORKBOOK_MODE_CONSOLIDATED,
) -> Path:
    """Grava o workbook consolidado ou separado em uma aba por amostra."""
    workbook_mode = validate_workbook_mode(workbook_mode)
    xlsx_path = Path(xlsx_path)
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        if workbook_mode == WORKBOOK_MODE_CONSOLIDATED:
            df_dados.to_excel(writer, index=False, sheet_name="DadosCompletos")
            df_resumo.to_excel(writer, index=False, sheet_name="AmostrasDetectadas")
        else:
            summary_sheet = "AmostrasDetectadas"
            df_resumo.to_excel(writer, index=False, sheet_name=summary_sheet)
            used_names = {summary_sheet.casefold()}

            if "Amostra" in df_dados.columns:
                sample_names = []
                if "SampleName" in df_resumo.columns:
                    sample_names.extend(df_resumo["SampleName"].dropna().tolist())
                sample_names.extend(df_dados["Amostra"].dropna().tolist())

                for sample_name in dict.fromkeys(sample_names):
                    sample_rows = df_dados[df_dados["Amostra"] == sample_name]
                    sheet_name = unique_sheet_name(sample_name, used_names)
                    sample_rows.to_excel(writer, index=False, sheet_name=sheet_name)

    return xlsx_path
