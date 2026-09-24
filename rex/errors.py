"""Exceções de domínio expostas pelo REX."""


class REXError(Exception):
    """Erro esperado durante uma operação do REX."""


class DocumentNotFoundError(REXError, FileNotFoundError):
    """O documento solicitado não existe."""


class InvalidDocumentError(REXError):
    """O arquivo não possui uma estrutura DOCX válida ou suportada."""


class DocumentLimitError(InvalidDocumentError):
    """O documento excede um limite seguro de processamento."""


class OCRInitializationError(REXError):
    """O mecanismo OCR não pôde ser inicializado."""


class ExportError(REXError):
    """Os resultados não puderam ser gravados."""
