import zipfile
import io
import os
import re
from pathlib import Path
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple, Any, Callable

import numpy as np
import pandas as pd
from PIL import Image

# Evita telemetria e arquivos persistentes do ONNX Runtime no diretório do projeto.
os.environ.setdefault("ORT_DISABLE_TELEMETRY", "1")

from rapidocr_onnxruntime import RapidOCR


class EDSExtractorEngine:
    """
    Motor de extração determinístico para laudos de MEV-EDS em formato Word (.docx).
    Processa arquivos 100% em memória, utilizando RapidOCR para reconhecimento
    ótico de micrografias e varredura linear do OpenXML.
    """

    def __init__(self, sample_regex: Optional[str] = None):
        pattern = sample_regex or r'[a-zA-ZÀ-ÿ]\d{2}\s+.+?\s+\d{5,7}'
        self.sample_regex = re.compile(pattern, re.IGNORECASE)
        self.ocr_engine = RapidOCR()

    def _extract_sample_from_image(self, img_bytes: bytes) -> Optional[str]:
        """Aplica OCR sobre a imagem garantindo composição com fundo branco se houver canal alfa."""
        try:
            im = Image.open(io.BytesIO(img_bytes))
            if im.width < 1200 or im.height < 800:
                return None

            if im.mode in ("RGBA", "LA"):
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[-1])
                img_to_ocr = bg
            else:
                img_to_ocr = im.convert("RGB")

            results, _ = self.ocr_engine(np.array(img_to_ocr))
            if not results:
                return None

            for _, text, _ in results:
                match = self.sample_regex.search(text)
                if match:
                    return match.group(0).strip()
            return None
        except Exception as exc:
            print(f"[OCR Aviso] Falha ao processar imagem: {exc}")
            return None

    def process(
        self,
        docx_path: str | Path,
        output_dir: Optional[str | Path] = None,
        status_callback: Optional[Callable[[str], None]] = None
    ) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Executa o pipeline completo de extração do documento .docx.
        Retorna uma tupla (df_dados_completos, df_resumo_amostras).
        """
        def _log(msg: str):
            print(msg)
            if status_callback:
                status_callback(msg)

        docx_path = Path(docx_path)
        if not docx_path.is_file():
            raise FileNotFoundError(f"Arquivo não encontrado: {docx_path}")

        out_path = Path(output_dir) if output_dir else docx_path.parent
        _log(f"Lendo documento em memória: {docx_path.name}")

        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        all_rows = []
        samples_meta = []
        current_sample: Optional[str] = None
        current_image: Optional[str] = None
        series_id = 1
        last_point_id = 0
        sample_counter = 0

        with zipfile.ZipFile(docx_path, "r") as zf:
            rels_xml = zf.read("word/_rels/document.xml.rels")
            rels_root = ET.fromstring(rels_xml)
            rel_ns_map = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
            id_to_target = {
                rel.get("Id"): rel.get("Target")
                for rel in rels_root.findall(".//r:Relationship", rel_ns_map)
            }

            doc_xml = zf.read("word/document.xml")
            doc_root = ET.fromstring(doc_xml)

            for elem in doc_root.iter():
                tag = elem.tag.split("}")[-1]

                if tag == "blip":
                    r_embed = elem.get(f"{{{r_ns}}}embed")
                    target = id_to_target.get(r_embed)
                    if target and target.startswith("media/"):
                        img_path_in_zip = f"word/{target}"
                        if img_path_in_zip in zf.namelist():
                            img_data = zf.read(img_path_in_zip)
                            detected_name = self._extract_sample_from_image(img_data)
                            if detected_name:
                                sample_counter += 1
                                current_sample = detected_name
                                current_image = target
                                series_id = 1
                                last_point_id = 0
                                samples_meta.append({
                                    "SampleIndex": sample_counter,
                                    "SampleName": current_sample,
                                    "ImageFile": current_image
                                })
                                _log(f"  [Amostra {sample_counter}] Detectada: '{current_sample}' ({target})")

                elif tag == "tbl":
                    if not current_sample:
                        continue

                    first_tc = elem.find(f".//{{{w_ns}}}tc")
                    first_text = "".join(first_tc.itertext()).strip() if first_tc is not None else ""

                    if not first_text.isdigit():
                        continue

                    point_id = int(first_text)

                    if point_id <= last_point_id:
                        series_id += 1
                    last_point_id = point_id

                    rows = elem.findall(f"{{{w_ns}}}tr")
                    for row in rows:
                        cells = row.findall(f"{{{w_ns}}}tc")
                        texts = ["".join(c.itertext()).strip() for c in cells]
                        joined = " ".join(texts).lower()

                        if not joined or "element" in joined or "total" in joined or "elemento" in joined:
                            continue

                        element = texts[0] if len(texts) > 0 else ""
                        line_type = texts[1] if len(texts) > 1 else ""

                        nums = re.findall(r"\d+\.\d+|\d+", " ".join(texts[2:]))
                        weight = nums[0] if len(nums) > 0 else ""
                        sigma = nums[1] if len(nums) > 1 else ""
                        atomic = nums[2] if len(nums) > 2 else ""

                        if element and element.isalpha() and len(element) <= 3:
                            all_rows.append({
                                "Amostra": current_sample,
                                "Micrografia": current_image,
                                "Serie": series_id,
                                "Ponto": point_id,
                                "Elemento": element,
                                "TipoLinha": line_type,
                                "PercentualPeso": weight,
                                "SigmaPeso": sigma,
                                "PercentualAtomico": atomic
                            })

        df_dados = pd.DataFrame(all_rows)
        df_resumo = pd.DataFrame(samples_meta)

        if not df_dados.empty:
            num_cols = ["PercentualPeso", "SigmaPeso", "PercentualAtomico"]
            df_dados[num_cols] = df_dados[num_cols].apply(pd.to_numeric, errors="coerce")

        xlsx_file = out_path / "amostras_organizadas_completas.xlsx"
        csv_file = out_path / "dados_extraidos.csv"

        with pd.ExcelWriter(xlsx_file, engine="openpyxl") as writer:
            df_dados.to_excel(writer, index=False, sheet_name="DadosCompletos")
            df_resumo.to_excel(writer, index=False, sheet_name="AmostrasDetectadas")

        df_dados.to_csv(csv_file, index=False, encoding="utf-8-sig")

        _log(f"\n[Exportação Concluída]")
        _log(f"  -> Excel: {xlsx_file.resolve()}")
        _log(f"  -> CSV:   {csv_file.resolve()}")

        return df_dados, df_resumo
