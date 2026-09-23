---
title: "Decisões Arquiteturais Centrais: RapidOCR e Parsing Linear"
date: "2026-09-23"
type: "decision"
tags: ["rapidocr", "onnxruntime", "openxml", "in-memory", "architecture"]
summary: "Por que o RapidOCR foi escolhido no lugar do Tesseract e como a varredura linear do OpenXML garantiu 100% de precisão de alinhamento."
---

# 2. Decisões Arquiteturais Fundamentais

## Decisão 1: RapidOCR via ONNX Runtime em vez de Tesseract
* **Motivo**: O Tesseract exigia a instalação de binários de sistema operacional externos (`tesseract.exe`), configuração manual de caminhos no Windows e quebrava em sistemas onde não estava instalado.
* **Solução**: O `rapidocr-onnxruntime` é executado 100% em Python via ONNX Runtime na CPU. Os modelos neurais do PP-OCRv4 rodam com ~0.15s por imagem e alcançam >91% de confiança sem necessidade de recortes arbitrários (*ROI crops* fixos).

## Decisão 2: Varredura Linear Determinística do OpenXML
* **Motivo**: Em vez de ler imagens e tabelas separadamente em loops desacoplados, o REX percorre os nós do arquivo `word/document.xml` na ordem exata de aparição no fluxo do texto.
* **Mecanismo**:
  1. O arquivo `word/_rels/document.xml.rels` mapeia o identificador `r:embed` da tag `<a:blip>` para o caminho real da imagem em `word/media/imageX.png`.
  2. Quando uma micrografia é identificada, seu rótulo é lido via OCR e se torna a `Amostra Ativa`.
  3. Todas as tabelas subsequentes de pontos (`Point 1`, `Point 2`...) são vinculadas diretamente a essa amostra ativa até que uma nova micrografia física apareça na árvore do documento.

## Decisão 3: Operação 100% em Memória
* O arquivo `.docx` é aberto diretamente como stream por `zipfile.ZipFile(docx_path)`. Nenhuma pasta temporária ou cópia `.zip` é criada em disco, prevenindo conflitos de bloqueio de arquivo no Windows e I/O desnecessário.
