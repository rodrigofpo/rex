"""Testes dos limites preventivos aplicados a contêineres DOCX."""

import zipfile
from pathlib import Path

import pytest

from rex.errors import DocumentLimitError, InvalidDocumentError
from rex.security import DocumentLimits, validate_archive


def test_rejects_archive_without_minimum_word_structure(tmp_path: Path):
    document = tmp_path / "incompleto.docx"
    with zipfile.ZipFile(document, "w") as archive:
        archive.writestr("outro.txt", "conteúdo")

    with zipfile.ZipFile(document) as archive:
        with pytest.raises(InvalidDocumentError, match="estrutura mínima"):
            validate_archive(archive, DocumentLimits())


def test_rejects_suspicious_compression_ratio(tmp_path: Path):
    document = tmp_path / "compactacao.docx"
    with zipfile.ZipFile(document, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", "A" * 20_000)
        archive.writestr("word/_rels/document.xml.rels", "<Relationships />")

    limits = DocumentLimits(max_compression_ratio=2.0)
    with zipfile.ZipFile(document) as archive:
        with pytest.raises(DocumentLimitError, match="taxa de compactação"):
            validate_archive(archive, limits)
