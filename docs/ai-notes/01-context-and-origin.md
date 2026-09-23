---
title: "Origem e Diagnóstico do Problema MEV-EDS"
date: "2026-09-23"
type: "context"
tags: ["mev", "eds", "history", "diagnostics", "openxml"]
summary: "Histórico da evolução do projeto rex a partir dos scripts experimentais docx_extractor e identificação da causa raiz dos descompassos."
---

# 1. Origem e Contexto do Projeto

O projeto **REX** nasceu para substituir uma coleção de mais de 15 scripts experimentais (`docx_extractor_V1` a `V5_modify_2`, `VGMNI`, `VCLDE`) que tentavam extrair dados de laudos de caracterização cerâmica MEV-EDS em arquivos Microsoft Word (`.docx`).

## O Problema dos Laudos MEV-EDS

1. **Separação não-relacional**:
   - O rótulo da amostra (ex: `A01 externo 071815`) não existe como texto no corpo do Word; ele está impresso graficamente dentro da micrografia (imagem PNG de alta resolução).
   - Os dados quantitativos de composição química elementar (% em peso, sigma, % atômica) residem em tabelas estruturadas do Word associadas a pontos de medição (`Point 1`, `Point 2`...).
2. **A Falha das Abordagens Anteriores**:
   - Os scripts legados descompactavam o `.docx` em disco e percorriam a pasta `word/media/` atribuindo `GroupID = 0, 1, 2...` em ordem alfabética de arquivos.
   - O documento real continha **123 imagens**, mas apenas **20 micrografias eram amostras**; as demais 103 imagens eram gráficos de espectro EDS e logos.
   - Essa iteração cega descompassava todos os nomes de amostras em relação às tabelas.
