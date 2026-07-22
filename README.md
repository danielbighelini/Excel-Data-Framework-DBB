# Excel Power Query Framework

Framework para preparação, validação, normalização e integração de dados utilizando Microsoft Excel e Power Query.

## Visão Geral

O Excel Power Query Framework fornece uma arquitetura padronizada para construção de processos ETL (Extract, Transform and Load) utilizando exclusivamente recursos nativos do Excel e do Power Query.

O framework organiza todo o fluxo de processamento em camadas independentes, permitindo importar dados de diferentes fontes, aplicar tratamentos, validar informações, normalizar entidades e produzir tabelas dimensionais para análise.

## Principais Características

- Arquitetura em camadas (SRC → STG → NRM → DIM/FATO)
- Configuração orientada por parâmetros
- Schema declarativo
- Tratamentos configuráveis por coluna
- Validações configuráveis por coluna
- Registro de ocorrências
- Controle de severidades
- Suporte a múltiplas fontes de dados
- Construção automática de dimensões e fatos
- Estrutura preparada para reutilização entre projetos

---

# Arquitetura

```text
                Excel / Banco / API / Arquivos
                           │
                           ▼
                     SRC (Origem)
                           │
                           ▼
                    STG (Stage)
         • Aplicação do Schema
         • Garantia de colunas
         • Tratamentos
         • Validações
         • Registro de ocorrências
                           │
                           ▼
                 NRM (Normalização)
         • Regras de negócio
         • Consistência entre entidades
         • Descarte de registros bloqueantes
                           │
                 ┌─────────┴─────────┐
                 ▼                   ▼
           DIMENSÕES             TABELAS FATO
```

---

# Estrutura do Projeto

```
par     Parâmetros
src     Fontes de dados
stg     Preparação dos dados
nrm     Normalização
dim     Dimensões
fato    Tabelas fato
cfg     Configurações materializadas
fx      Funções reutilizáveis
diag    Diagnósticos
bench   Benchmarks
tst     Testes
```

---

# Pipeline de Processamento

## 1. Origem (SRC)

Responsável pela conexão com as fontes de dados.

Exemplos:

- Excel
- CSV
- JSON
- XML
- PDF
- REST
- OData
- SOAP
- SQL Server
- Oracle
- PostgreSQL
- MySQL
- SharePoint

---

## 2. Stage (STG)

Responsável pela preparação física dos dados.

Nesta etapa o framework:

- remove linhas vazias
- garante a existência das colunas
- aplica o schema
- converte tipos
- executa tratamentos
- executa validações
- registra ocorrências

---

## 3. Normalização (NRM)

Responsável pelas regras de negócio.

Nesta etapa são executadas operações como:

- descarte de registros bloqueantes
- validações entre entidades
- resolução de relacionamentos
- cálculos derivados
- consolidação de informações

---

## 4. Dimensões

Construção das dimensões do modelo analítico.

Exemplos:

- Cliente
- Produto
- Calendário

---

## 5. Fatos

Construção das tabelas fato utilizadas pelo modelo dimensional.

---

# Configuração

Todo o comportamento do framework é controlado por tabelas de parâmetros no Excel.

Entre elas:

- Parâmetros da Pasta de Trabalho
- Schema de Dados
- Tratamentos
- Validações
- Severidades
- Valores Lógicos
- Tipos de Dados
- Formatos de Arquivo
- Headers REST

---

# Schema Declarativo

Cada coluna pode definir:

- Tipo de dado
- Obrigatoriedade
- Ordem
- Tratamentos
- Validações

Exemplo:

| Coluna | Tipo | Tratamentos | Validações |
|--------|------|-------------|------------|
| CPF | Text | DIGITS | CPF |
| Nome | Text | TRIM;PROPER | REQUIRED |
| Idade | Int64 | | MIN(18) |

As validações podem receber parâmetros utilizando a sintaxe:

```
VALIDACAO(param1,param2,...)
```

Exemplos:

```
MIN(0)
MAX(100)
INTERVAL(1,10)
LIST(SP,RJ,MG)
```

---

# Sistema de Ocorrências

Cada validação pode gerar ocorrências contendo informações como:

- Código
- Severidade
- Etapa
- Tabela
- Coluna
- Valor original
- Valor processado
- Mensagem
- Detalhes

As severidades determinam se uma ocorrência é apenas informativa ou bloqueante.

---

# Performance

O framework foi desenvolvido priorizando processamento vetorizado do Power Query.

As principais estratégias adotadas incluem:

- Record lookup O(1)
- Table.Buffer apenas quando necessário
- List.Buffer para reutilização
- Processamento por coluna
- Materialização de configurações
- Minimização de reconstruções de tabela

---

# Diagnóstico

O framework possui consultas de diagnóstico para facilitar manutenção e evolução:

- Consultas existentes
- Tabelas do Excel
- Estrutura do projeto
- Benchmarks
- Métricas de processamento

---

# Tecnologias

- Microsoft Excel
- Power Query (Linguagem M)

---

# Licença

Defina a licença de utilização conforme a necessidade do projeto.