# Getting Started

Este guia apresenta o caminho mínimo para configurar e executar o Excel Data Framework DBB em um projeto Excel + Power Query.

O objetivo é criar um primeiro pipeline utilizando:

`SRC → STG → TRN → QA → NRM`

Para entender a arquitetura completa do framework, consulte [Architecture](architecture.md).

---

## 1. Pré-requisitos

O framework utiliza:

- Microsoft Excel com Power Query;
- Excel 365 recomendado;
- conhecimento básico de Power Query e linguagem M;
- uma fonte de dados compatível com Power Query.

O framework foi projetado para trabalhar com diferentes fontes de dados, mantendo a lógica de transformação e qualidade independente da origem.

---

## 2. Estrutura básica

Um projeto utilizando o framework é organizado em camadas:

```text
Fonte de dados
      │
      ▼
    SRC
      │
      ▼
    STG
      │
      ▼
    TRN
      │
      ▼
     QA
      │
      ▼
    NRM
      │
      ▼
 DIM / FATO