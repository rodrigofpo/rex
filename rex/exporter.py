"""Persistência dos resultados extraídos pelo REX."""

import logging
import os
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from rex.errors import ExportError
from rex.models import ExportPaths, ExtractionResult
from rex.workbook import WORKBOOK_MODE_CONSOLIDATED, write_excel_workbook

logger = logging.getLogger(__name__)


def export_result(
        result: ExtractionResult,
        output_dir: str | Path,
        source_path: str | Path,
        workbook_mode: str = WORKBOOK_MODE_CONSOLIDATED,
) -> ExportPaths:
    """Grava Excel e CSV sem acoplar persistência ao parser OpenXML/OCR."""
    output_path = Path(output_dir)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    excel_path = output_path / f"{Path(source_path).stem}_REX-{timestamp}.xlsx"
    csv_path = output_path / "dados_extraidos.csv"

    try:
        output_path.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix=".rex-export-", dir=output_path) as temp_dir:
            temporary_path = Path(temp_dir)
            temporary_excel = temporary_path / excel_path.name
            temporary_csv = temporary_path / csv_path.name

            write_excel_workbook(
                result.data,
                result.samples,
                temporary_excel,
                workbook_mode,
            )
            result.data.to_csv(temporary_csv, index=False, encoding="utf-8-sig")

            with zipfile.ZipFile(temporary_excel) as workbook_archive:
                damaged_member = workbook_archive.testzip()
                if damaged_member is not None:
                    raise ExportError(
                        f"A planilha gerada está corrompida no item '{damaged_member}'."
                    )

            replacements = ((temporary_excel, excel_path), (temporary_csv, csv_path))
            backups: dict[Path, Path | None] = {}
            for _, destination in replacements:
                if destination.exists():
                    backup = temporary_path / f"backup-{destination.name}"
                    shutil.copy2(destination, backup)
                    backups[destination] = backup
                else:
                    backups[destination] = None

            replaced: list[Path] = []
            try:
                for source, destination in replacements:
                    os.replace(source, destination)
                    replaced.append(destination)
            except OSError:
                for destination in reversed(replaced):
                    backup = backups[destination]
                    if backup is None:
                        destination.unlink()
                    else:
                        os.replace(backup, destination)
                raise
    except ExportError:
        raise
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        raise ExportError(f"Falha ao exportar os resultados para '{output_path}': {exc}") from exc

    logger.info("Resultados exportados para %s", output_path.resolve())
    return ExportPaths(excel=excel_path, csv=csv_path)
