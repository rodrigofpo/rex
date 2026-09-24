"""Testes de fumaça e validação do motor REX."""

from pathlib import Path

import pytest

from rex.core.engine import EDSExtractorEngine
from rex.errors import DocumentLimitError, DocumentNotFoundError, InvalidDocumentError
from rex.security import DocumentLimits


def test_engine_initialization_is_lazy():
    created_engines = []
    fake_ocr = object()
    engine = EDSExtractorEngine(ocr_factory=lambda: created_engines.append(fake_ocr) or fake_ocr)

    assert engine._ocr_engine is None
    assert engine.sample_regex is not None
    assert engine.ocr_engine is fake_ocr
    assert engine.ocr_engine is fake_ocr
    assert created_engines == [fake_ocr]


def test_missing_document_has_domain_error(tmp_path: Path):
    with pytest.raises(DocumentNotFoundError, match="Arquivo não encontrado"):
        EDSExtractorEngine(ocr_factory=lambda: object()).extract(tmp_path / "ausente.docx")


def test_invalid_docx_has_domain_error(tmp_path: Path):
    invalid_docx = tmp_path / "invalido.docx"
    invalid_docx.write_text("não é um arquivo zip", encoding="utf-8")

    with pytest.raises(InvalidDocumentError, match="não é um DOCX válido"):
        EDSExtractorEngine(ocr_factory=lambda: object()).extract(invalid_docx)


def test_rejects_document_above_configured_file_limit(tmp_path: Path):
    oversized = tmp_path / "grande.docx"
    oversized.write_bytes(b"12345")
    limits = DocumentLimits(max_file_bytes=4)

    with pytest.raises(DocumentLimitError, match="o limite é"):
        EDSExtractorEngine(ocr_factory=lambda: object(), limits=limits).extract(oversized)


def test_public_engine_import_remains_compatible():
    from rex import EDSExtractorEngine as PublicEDSExtractorEngine

    assert PublicEDSExtractorEngine is EDSExtractorEngine
