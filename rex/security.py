"""Limites preventivos para documentos OOXML processados localmente."""

from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

from rex.errors import DocumentLimitError, InvalidDocumentError


@dataclass(frozen=True)
class DocumentLimits:
    """Limites configuráveis contra consumo excessivo de disco, memória e CPU."""

    max_file_bytes: int = 250 * 1024 * 1024
    max_members: int = 10_000
    max_uncompressed_bytes: int = 1024 * 1024 * 1024
    max_xml_bytes: int = 32 * 1024 * 1024
    max_image_bytes: int = 64 * 1024 * 1024
    max_image_pixels: int = 100_000_000
    max_compression_ratio: float = 500.0


def validate_document_file(path: Path, limits: DocumentLimits) -> None:
    """Rejeita arquivos grandes antes mesmo de abrir o contêiner ZIP."""
    file_size = path.stat().st_size
    if file_size > limits.max_file_bytes:
        raise DocumentLimitError(
            f"O documento possui {file_size / (1024 * 1024):.1f} MB; "
            f"o limite é {limits.max_file_bytes / (1024 * 1024):.0f} MB."
        )


def validate_archive(archive: ZipFile, limits: DocumentLimits) -> None:
    """Inspeciona metadados ZIP sem descompactar o conteúdo."""
    members = archive.infolist()
    if len(members) > limits.max_members:
        raise DocumentLimitError(
            f"O documento contém {len(members)} itens; o limite é {limits.max_members}."
        )

    total_uncompressed = sum(member.file_size for member in members)
    if total_uncompressed > limits.max_uncompressed_bytes:
        raise DocumentLimitError("O conteúdo descompactado do documento excede o limite seguro.")

    required_members = {"word/document.xml", "word/_rels/document.xml.rels"}
    available_members = {member.filename for member in members}
    if not required_members.issubset(available_members):
        raise InvalidDocumentError("O arquivo não contém a estrutura mínima de um documento Word.")

    for member in members:
        if member.flag_bits & 0x1:
            raise InvalidDocumentError("Documentos Word com conteúdo ZIP criptografado não são suportados.")

        if member.file_size > 0:
            if member.compress_size == 0:
                raise DocumentLimitError("O documento contém um item com compactação inválida.")
            compression_ratio = member.file_size / member.compress_size
            if compression_ratio > limits.max_compression_ratio:
                raise DocumentLimitError(
                    f"O item '{member.filename}' excede a taxa de compactação segura."
                )

        if member.filename.endswith(".xml") and member.file_size > limits.max_xml_bytes:
            raise DocumentLimitError(f"O XML '{member.filename}' excede o limite seguro.")
        if member.filename.startswith("word/media/") and member.file_size > limits.max_image_bytes:
            raise DocumentLimitError(f"A imagem '{member.filename}' excede o limite seguro.")
