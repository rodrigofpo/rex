"""Testes da camada de exportação desacoplada."""

from pathlib import Path

import pandas as pd
import pytest

import rex.exporter as exporter_module
from rex.errors import ExportError
from rex.exporter import export_result
from rex.models import ExtractionResult


def test_exports_structured_result_without_engine_dependency(tmp_path: Path):
    result = ExtractionResult(
        data=pd.DataFrame(
            [
                {
                    "Amostra": "A01 externo 071815",
                    "Micrografia": "media/image1.png",
                    "Serie": 1,
                    "Ponto": 1,
                    "Elemento": "Si",
                    "TipoLinha": "K series",
                    "PercentualPeso": 59.57,
                    "SigmaPeso": 0.23,
                    "PercentualAtomico": 73.86,
                }
            ]
        ),
        samples=pd.DataFrame(
            [
                {
                    "SampleIndex": 1,
                    "SampleName": "A01 externo 071815",
                    "ImageFile": "media/image1.png",
                }
            ]
        ),
    )

    paths = export_result(result, tmp_path)

    assert paths.excel.is_file()
    assert paths.csv.is_file()


def test_failed_export_preserves_previous_files(tmp_path: Path, monkeypatch):
    previous_excel = tmp_path / "amostras_organizadas_completas.xlsx"
    previous_csv = tmp_path / "dados_extraidos.csv"
    previous_excel.write_bytes(b"planilha-anterior")
    previous_csv.write_text("csv-anterior", encoding="utf-8")
    result = ExtractionResult(data=pd.DataFrame(), samples=pd.DataFrame())

    def fail_workbook(*args, **kwargs):
        raise OSError("disco indisponível")

    monkeypatch.setattr(exporter_module, "write_excel_workbook", fail_workbook)

    with pytest.raises(ExportError, match="Falha ao exportar"):
        export_result(result, tmp_path)

    assert previous_excel.read_bytes() == b"planilha-anterior"
    assert previous_csv.read_text(encoding="utf-8") == "csv-anterior"
    assert not list(tmp_path.glob(".rex-export-*"))


def test_second_replacement_failure_restores_previous_pair(tmp_path: Path, monkeypatch):
    previous_excel = tmp_path / "amostras_organizadas_completas.xlsx"
    previous_csv = tmp_path / "dados_extraidos.csv"
    previous_excel.write_bytes(b"planilha-anterior")
    previous_csv.write_text("csv-anterior", encoding="utf-8")
    result = ExtractionResult(data=pd.DataFrame({"Amostra": ["A01"]}), samples=pd.DataFrame())
    real_replace = exporter_module.os.replace

    def fail_csv_replacement(source, destination):
        if Path(destination) == previous_csv:
            raise OSError("falha na substituição do CSV")
        return real_replace(source, destination)

    monkeypatch.setattr(exporter_module.os, "replace", fail_csv_replacement)

    with pytest.raises(ExportError, match="falha na substituição do CSV"):
        export_result(result, tmp_path)

    assert previous_excel.read_bytes() == b"planilha-anterior"
    assert previous_csv.read_text(encoding="utf-8") == "csv-anterior"
