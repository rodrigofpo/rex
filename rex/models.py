"""Modelos compartilhados entre extração, validação e exportação."""

from dataclasses import dataclass
from pathlib import Path

import pandas as pd


@dataclass(frozen=True)
class ExtractionResult:
    """Dados estruturados produzidos pelo parser, ainda sem persistência."""

    data: pd.DataFrame
    samples: pd.DataFrame


@dataclass(frozen=True)
class ExportPaths:
    """Arquivos persistidos por uma exportação concluída."""

    excel: Path
    csv: Path
