"""Motor de extração OpenXML e OCR do REX."""

import io
import logging
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pandas as pd

from rex.errors import (
    DocumentLimitError,
    DocumentNotFoundError,
    InvalidDocumentError,
    OCRInitializationError,
)
from rex.exporter import export_result
from rex.models import ExportPaths, ExtractionResult
from rex.security import DocumentLimits, validate_archive, validate_document_file
from rex.workbook import WORKBOOK_MODE_CONSOLIDATED

# Deve ser definido antes do primeiro import do ONNX Runtime.
os.environ.setdefault("ORT_DISABLE_TELEMETRY", "1")

logger = logging.getLogger(__name__)


class EDSExtractorEngine:
    """Extrai amostras e medições EDS mantendo a ordem do OpenXML."""

    def __init__(
            self,
            sample_regex: str | None = None,
            ocr_factory: Callable[[], Any] | None = None,
            limits: DocumentLimits | None = None,
    ):
        pattern = sample_regex or r"[a-zA-ZÀ-ÿ]\d{2}\s+.+?\s+\d{5,7}"
        self.sample_regex = re.compile(pattern, re.IGNORECASE)
        self._ocr_factory = ocr_factory
        self._ocr_engine: Any | None = None
        self.limits = limits or DocumentLimits()
        self.last_export_paths: ExportPaths | None = None

    @property
    def ocr_engine(self) -> Any:
        """Compatibilidade pública com inicialização somente no primeiro uso."""
        return self._get_ocr_engine()

    def _get_ocr_engine(self) -> Any:
        if self._ocr_engine is None:
            try:
                if self._ocr_factory is None:
                    from rapidocr_onnxruntime import RapidOCR

                    self._ocr_engine = RapidOCR()
                else:
                    self._ocr_engine = self._ocr_factory()
            except Exception as exc:
                raise OCRInitializationError(
                    f"Não foi possível inicializar o mecanismo OCR: {exc}"
                ) from exc
        return self._ocr_engine

    def _extract_sample_from_image(self, img_bytes: bytes) -> str | None:
        """Executa OCR, compondo imagens transparentes sobre fundo branco."""
        try:
            import numpy as np
            from PIL import Image

            with Image.open(io.BytesIO(img_bytes)) as image:
                if image.width * image.height > self.limits.max_image_pixels:
                    raise DocumentLimitError(
                        f"A imagem possui {image.width * image.height:,} pixels e excede o limite seguro."
                    )
                if image.width < 1200 or image.height < 800:
                    return None

                if image.mode in ("RGBA", "LA"):
                    background = Image.new("RGB", image.size, (255, 255, 255))
                    background.paste(image, mask=image.getchannel("A"))
                    image_for_ocr = background
                else:
                    image_for_ocr = image.convert("RGB")

                results, _ = self._get_ocr_engine()(np.array(image_for_ocr))

            if not results:
                return None
            for _, text, _ in results:
                match = self.sample_regex.search(text)
                if match:
                    return match.group(0).strip()
            return None
        except (DocumentLimitError, OCRInitializationError):
            raise
        except Exception as exc:
            logger.warning("Falha ao processar imagem no OCR: %s", exc)
            return None

    @staticmethod
    def _emit_status(message: str, callback: Callable[[str], None] | None) -> None:
        logger.info(message)
        if callback:
            callback(message)

    def extract(
            self,
            docx_path: str | Path,
            status_callback: Callable[[str], None] | None = None,
    ) -> ExtractionResult:
        """Extrai e estrutura os dados sem criar arquivos de saída."""
        document_path = Path(docx_path)
        if not document_path.is_file():
            raise DocumentNotFoundError(f"Arquivo não encontrado: {document_path}")
        validate_document_file(document_path, self.limits)

        self._emit_status(f"Lendo documento em memória: {document_path.name}", status_callback)
        try:
            return self._extract_openxml(document_path, status_callback)
        except OCRInitializationError:
            raise
        except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
            raise InvalidDocumentError(
                f"O arquivo '{document_path.name}' não é um DOCX válido ou está corrompido."
            ) from exc

    def _extract_openxml(
            self,
            document_path: Path,
            status_callback: Callable[[str], None] | None,
    ) -> ExtractionResult:
        word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        relationship_namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        all_rows: list[dict[str, Any]] = []
        samples_meta: list[dict[str, Any]] = []
        current_sample: str | None = None
        current_image: str | None = None
        series_id = 1
        last_point_id = 0
        sample_counter = 0

        with zipfile.ZipFile(document_path, "r") as archive:
            validate_archive(archive, self.limits)
            relationships = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
            relationship_map = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
            id_to_target = {
                relationship.get("Id"): relationship.get("Target")
                for relationship in relationships.findall(".//r:Relationship", relationship_map)
            }
            document = ET.fromstring(archive.read("word/document.xml"))
            archive_members = set(archive.namelist())

            for element_node in document.iter():
                tag = element_node.tag.split("}")[-1]
                if tag == "blip":
                    relationship_id = element_node.get(f"{{{relationship_namespace}}}embed")
                    target = id_to_target.get(relationship_id)
                    if target and target.startswith("media/"):
                        image_path = f"word/{target}"
                        if image_path in archive_members:
                            detected_name = self._extract_sample_from_image(archive.read(image_path))
                            if detected_name:
                                sample_counter += 1
                                current_sample = detected_name
                                current_image = target
                                series_id = 1
                                last_point_id = 0
                                samples_meta.append(
                                    {
                                        "SampleIndex": sample_counter,
                                        "SampleName": current_sample,
                                        "ImageFile": current_image,
                                    }
                                )
                                self._emit_status(
                                    f"  [Amostra {sample_counter}] Detectada: "
                                    f"'{current_sample}' ({target})",
                                    status_callback,
                                )
                elif tag == "tbl" and current_sample:
                    first_cell = element_node.find(f".//{{{word_namespace}}}tc")
                    first_text = "".join(first_cell.itertext()).strip() if first_cell is not None else ""
                    if not first_text.isdigit():
                        continue

                    point_id = int(first_text)
                    if point_id <= last_point_id:
                        series_id += 1
                    last_point_id = point_id

                    for row in element_node.findall(f"{{{word_namespace}}}tr"):
                        cells = row.findall(f"{{{word_namespace}}}tc")
                        texts = ["".join(cell.itertext()).strip() for cell in cells]
                        joined = " ".join(texts).lower()
                        if not joined or "element" in joined or "total" in joined or "elemento" in joined:
                            continue

                        element = texts[0] if texts else ""
                        line_type = texts[1] if len(texts) > 1 else ""
                        numbers = re.findall(r"\d+\.\d+|\d+", " ".join(texts[2:]))
                        weight = numbers[0] if numbers else ""
                        sigma = numbers[1] if len(numbers) > 1 else ""
                        atomic = numbers[2] if len(numbers) > 2 else ""
                        if element and element.isalpha() and len(element) <= 3:
                            all_rows.append(
                                {
                                    "Amostra": current_sample,
                                    "Micrografia": current_image,
                                    "Serie": series_id,
                                    "Ponto": point_id,
                                    "Elemento": element,
                                    "TipoLinha": line_type,
                                    "PercentualPeso": weight,
                                    "SigmaPeso": sigma,
                                    "PercentualAtomico": atomic,
                                }
                            )

        data = pd.DataFrame(all_rows)
        samples = pd.DataFrame(samples_meta)
        if not data.empty:
            numeric_columns = ["PercentualPeso", "SigmaPeso", "PercentualAtomico"]
            data[numeric_columns] = data[numeric_columns].apply(pd.to_numeric, errors="coerce")
        return ExtractionResult(data=data, samples=samples)

    def process(
            self,
            docx_path: str | Path,
            output_dir: str | Path | None = None,
            status_callback: Callable[[str], None] | None = None,
            workbook_mode: str = WORKBOOK_MODE_CONSOLIDATED,
    ) -> tuple[pd.DataFrame, pd.DataFrame]:
        """Mantém a API histórica, orquestrando extração e exportação."""
        document_path = Path(docx_path)
        self.last_export_paths = None
        result = self.extract(document_path, status_callback=status_callback)
        paths = export_result(
            result,
            output_dir=Path(output_dir) if output_dir else document_path.parent,
            source_path=document_path,
            workbook_mode=workbook_mode,
        )
        self.last_export_paths = paths
        self._emit_status("\n[Exportação Concluída]", status_callback)
        self._emit_status(f"  -> Excel: {paths.excel.resolve()}", status_callback)
        self._emit_status(f"  -> CSV:   {paths.csv.resolve()}", status_callback)
        return result.data, result.samples
