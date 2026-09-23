"""
Testes de fumaça e validação do motor REX.
"""

from rex.core.engine import EDSExtractorEngine


def test_engine_initialization():
    engine = EDSExtractorEngine()
    assert engine.ocr_engine is not None
    assert engine.sample_regex is not None
    print("[OK] EDSExtractorEngine inicializado com sucesso.")


def test_public_engine_import_remains_compatible():
    from rex import EDSExtractorEngine as PublicEDSExtractorEngine

    assert PublicEDSExtractorEngine is EDSExtractorEngine


if __name__ == "__main__":
    test_engine_initialization()
