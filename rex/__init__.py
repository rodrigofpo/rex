"""
REX — MEV-EDS Report Extractor
Pipeline determinístico e inteligente para extração de dados de laudos MEV-EDS.
"""

from typing import TYPE_CHECKING, Any

__version__ = "1.0.0"
__author__ = "Rodrigo Fernando Pinheiro Oliveira"

if TYPE_CHECKING:
    from rex.core.engine import EDSExtractorEngine

__all__ = ["EDSExtractorEngine", "__version__"]


def __getattr__(name: str) -> Any:
    """Mantém a API pública sem carregar o runtime OCR ao importar o pacote."""
    if name == "EDSExtractorEngine":
        from rex.core.engine import EDSExtractorEngine

        return EDSExtractorEngine
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
