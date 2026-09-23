---
title: "Interfaces Multiplataforma e Automação CI/CD"
date: "2026-09-23"
type: "rule"
tags: ["cli", "web-ui", "customtkinter", "github-actions", "cicd"]
summary: "Especificação das 3 modalidades de uso (CLI, Web, Desktop) e matriz de compilação automatizada no GitHub Actions."
---

# 4. Interfaces e Automação Multiplataforma

## Modalidades de Execução

1. **CLI (`rex extract <docx_path>`)**:
   - Ideal para pipelines em lote, automações de laboratório ou scripts de processamento em lote.
2. **Web App Local (`rex web`)**:
   - Servidor HTTP embutido em Python puro (`http.server`) com interface responsiva em HTML/CSS/JS.
   - Não depende de bibliotecas gráficas do sistema operacional (como `tkinter` no Linux), garantindo execução instantânea em qualquer distribuição Linux, macOS ou Windows.
3. **Desktop App (`rex gui`)**:
   - Interface com CustomTkinter para usuários que preferem janelas nativas.

## Automação CI/CD (GitHub Actions)

* **`.github/workflows/ci.yml`**: Roda testes cruzados em matriz paralela (`ubuntu-latest`, `windows-latest`, `macos-latest` com Python 3.11 e 3.12) a cada push na branch `main`.
* **`.github/workflows/release.yml`**: Disparado em tags de versão (`v*`). Utiliza o PyInstaller para empacotar binários portáteis com os modelos do RapidOCR embutidos para Windows (`.exe`), Linux e macOS, publicando-os automaticamente na aba Releases do GitHub.
