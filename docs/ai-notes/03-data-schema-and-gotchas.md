---
title: "Gotchas Críticos Resolvidos e Esquema de Dados"
date: "2026-09-23"
type: "gotcha"
tags: ["gotchas", "transparency", "alpha-channel", "data-schema", "eds"]
summary: "Solução para o canal alfa que gerava texto preto sobre preto e tratamento de múltiplas séries de pontos dentro da mesma amostra."
---

# 3. Gotchas Críticos e Esquema de Dados

## Gotcha 1: O Canal Alfa (Transparência) que Ocultava o Texto
* **Problema**: O software do microscópio exporta imagens PNG em modo `RGBA` onde o fundo é transparente e o texto é preto.
* **Impacto**: Ao converter diretamente com `im.convert('RGB')` na biblioteca Pillow, o canal transparente era preenchido com preto `(0, 0, 0)`. O texto preto ficava sobre fundo preto, tornando o OCR completamente cego.
* **Solução**:
  ```python
  bg = Image.new("RGB", im.size, (255, 255, 255))
  bg.paste(im, mask=im.split()[-1])  # Usa o canal alfa como máscara
  ```

## Gotcha 2: Repetição de `Point 1` sob a Mesma Amostra
* **Problema**: Algumas amostras (como `A01 externo 071817`) possuem duas áreas de análise no mesmo laudo: uma série com pontos `[1, 2, 3, 4, 5]` e outra com `[1, 2, 3, 4, 5, 6]`.
* **Impacto**: Os códigos legados tinham a regra `if point_id == 1: current_group_id += 1`. Isso causava falso incremento de amostras, dividindo uma única amostra em duas ou três e descompassando as seguintes.
* **Solução**: O identificador de amostra só muda na detecção de uma nova imagem. Quando o número do ponto retrocede (`point_id <= last_point_id`), incrementa-se apenas a coluna `Serie`, preservando a mesma amostra.

## Esquema Final dos Dados (715 Linhas / 20 Amostras)
* `Amostra` (string): Código extraído via OCR (ex: `A01 externo 071815`)
* `Micrografia` (string): Caminho físico no pacote OpenXML (ex: `media/image1.png`)
* `Serie` (int): Identificador da sub-série de pontos (1, 2...)
* `Ponto` (int): Número da análise pontual (1 a 6)
* `Elemento` (string): Símbolo químico (`Al`, `Si`, `O`, `Fe`, `Ca`...)
* `TipoLinha` (string): Linha espectral (ex: `K series`)
* `PercentualPeso` (float): Concentração ponderal quantitativa
* `SigmaPeso` (float): Incerteza analítica
* `PercentualAtomico` (float): Concentração estequiométrica atômica
