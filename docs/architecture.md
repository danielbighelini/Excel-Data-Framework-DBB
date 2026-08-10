# Architecture

## 1. Visão Geral

O Excel Data Framework DBB é organizado como uma arquitetura de processamento de dados em camadas, orientada por **Schema** e baseada em **pipelines compilados**.

A arquitetura separa três responsabilidades fundamentais:

1. **Configuração** — define como os dados devem ser processados.
2. **Execução** — interpreta e executa as regras configuradas.
3. **Dados** — percorrem as diferentes etapas do pipeline.

O fluxo principal de processamento é:

```text
SRC → STG → TRN → QA → NRM → DIM/FATO
```

Cada camada possui uma responsabilidade específica e deve evitar assumir responsabilidades pertencentes às demais.

---

## 2. Princípios Arquiteturais

### 2.1 Separação de responsabilidades

Cada etapa do pipeline possui uma finalidade específica:

```text
SRC       → Extração
STG       → Estruturação
TRN       → Tratamento
QA        → Qualidade
NRM       → Normalização
DIM/FATO  → Modelagem analítica
```

Essa separação permite que alterações em uma etapa não exijam necessariamente alterações nas demais.

---

### 2.2 Arquitetura orientada por Schema

O comportamento do processamento é definido por metadados.

O Schema descreve características como:

* tabelas;
* colunas;
* tipos;
* obrigatoriedade;
* tratamentos;
* validações;
* ordem das colunas;
* chaves de negócio.

O código do framework fornece o mecanismo de execução, enquanto o Schema fornece as regras que serão executadas.

```text
                 Schema
                    │
                    ▼
            Pipeline Compilado
                    │
                    ▼
                Execução
                    │
                    ▼
                  Dados
```

Essa abordagem reduz a necessidade de implementar regras específicas para cada tabela.

---

### 2.3 Separação entre configuração e execução

A arquitetura diferencia claramente as definições utilizadas pelo framework das funções responsáveis por executá-las.

```text
┌──────────────────────────────────┐
│          CONFIGURAÇÃO            │
│                                  │
│ Schema                           │
│ Parâmetros                       │
│ Tipos                            │
│ Tratamentos                      │
│ Validações                       │
└────────────────┬─────────────────┘
                 │
                 ▼
┌──────────────────────────────────┐
│            COMPILAÇÃO            │
│                                  │
│ fxPipeline                       │
│ cfgPipeline                      │
└────────────────┬─────────────────┘
                 │
                 ▼
┌──────────────────────────────────┐
│             EXECUÇÃO             │
│                                  │
│ STG → TRN → QA → NRM             │
└──────────────────────────────────┘
```

---

## 3. Arquitetura em Camadas

O pipeline completo é organizado nas seguintes camadas:

```text
┌─────────────────────────────────────────────┐
│ SRC — Source / Extração                     │
└──────────────────────┬──────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────┐
│ STG — Staging / Estruturação                │
└──────────────────────┬──────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────┐
│ TRN — Transformation / Tratamento           │
└──────────────────────┬──────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────┐
│ QA — Quality Assurance / Validação          │
└──────────────────────┬──────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────┐
│ NRM — Normalization / Normalização          │
└──────────────────────┬──────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────┐
│ DIM / FATO — Modelo Dimensional             │
└─────────────────────────────────────────────┘
```

---

# 4. SRC — Source

## Responsabilidade

A camada SRC representa a **origem dos dados**.

Pode representar fontes como:

* tabelas do Excel;
* arquivos;
* bancos de dados;
* APIs;
* outras fontes disponíveis no Power Query.

Exemplos:

```text
srcClientes
srcProdutos
srcVendas
```

A responsabilidade dessa camada é disponibilizar os dados para o pipeline.

## Regra arquitetural

A SRC deve representar a fonte e evitar incorporar regras de transformação, validação ou modelagem que pertençam às etapas seguintes.

```text
Fonte externa
     │
     ▼
    SRC
     │
     ▼
  Dados brutos
```

---

# 5. STG — Staging

## Responsabilidade

A camada STG executa a **preparação estrutural dos dados**.

Entre suas responsabilidades estão:

* remoção de linhas completamente vazias;
* normalização de nomes de colunas;
* validação estrutural;
* aplicação de tipos básicos;
* reordenação de colunas;
* registro de ocorrências relacionadas à tipagem.

O framework utiliza `fxStgAplicar()` para executar essa etapa.

Exemplo:

```text
srcClientes
     │
     ▼
fxStgAplicar()
     │
     ▼
stgClientes
```

## Ocorrências de tipo

Falhas de conversão de tipo são registradas em `_STG_Ocorrencias`.

A arquitetura evita que uma falha de conversão seja simplesmente silenciada e transformada em `null`.

Isso permite preservar o diagnóstico do problema sem perder a informação sobre sua ocorrência.

## Regra arquitetural

A STG é responsável pela **estrutura e tipagem inicial**, não pela aplicação de regras de negócio ou validações de qualidade.

---

# 6. TRN — Transformation

## Responsabilidade

A camada TRN aplica os **tratamentos e padronizações definidos pelo Schema**.

Exemplos de tratamentos:

```text
TRIM
UPPER
LOWER
PROPER
CLEAN
DIGITS
ALPHANUMERIC
NORMALIZEBASIC
NUMBER
REPLACE
PADLEFT
PADRIGHT
```

O processamento é realizado por `fxTrnAplicar()`.

```text
stgClientes
     │
     ▼
fxTrnAplicar()
     │
     ▼
trnClientes
```

Os operadores são compilados de acordo com as configurações da coluna.

## Regra arquitetural

A TRN deve modificar ou padronizar os valores conforme as regras configuradas.

Ela não é responsável por determinar se o registro é válido.

Essa responsabilidade pertence à camada QA.

---

# 7. QA — Quality Assurance

## Responsabilidade

A camada QA avalia a **qualidade dos dados**.

O objetivo é identificar e classificar problemas sem necessariamente remover os registros.

O processamento é realizado por:

```text
fxQaValidar()
```

O resultado contém informações de controle como:

```text
_QA_Status
_QA_Ocorrencias
```

O status pode assumir valores como:

```text
OK
AVISO
ERRO
```

As ocorrências registram informações sobre os problemas encontrados, incluindo dados como:

* código;
* severidade;
* coluna;
* tipo;
* descrição.

---

## 7.1 Separação entre dado e diagnóstico

Uma característica importante da arquitetura é separar o dado processado de seu diagnóstico de qualidade.

```text
                 Registro
                    │
          ┌─────────┴─────────┐
          │                   │
          ▼                   ▼
       Dados              Diagnóstico
                              │
                              ▼
                       _QA_Ocorrencias
                              │
                              ▼
                         _QA_Status
```

Isso permite que um registro com problema permaneça disponível para auditoria, análise ou correção.

---

## 7.2 Filtragem

A decisão de remover ou manter registros ocorre posteriormente, utilizando funções como:

```text
fxQaFiltrarPorStatus()
```

Assim, a validação e a filtragem são responsabilidades distintas.

```text
QA
 │
 ├── Registra problemas
 │
 └── Classifica registros
          │
          ▼
     Filtragem opcional
```

## Regra arquitetural

QA **avalia** os dados.

QA não deve ser utilizado como substituto da camada de transformação.

---

# 8. NRM — Normalization

## Responsabilidade

A camada NRM prepara os dados para consumo pelas estruturas analíticas.

Entre suas responsabilidades está a **deduplicação baseada em chaves de negócio**.

O processamento é realizado por:

```text
fxNrmAplicar()
```

Fluxo:

```text
qaClientes
     │
     ▼
Registros selecionados
     │
     ▼
fxNrmAplicar()
     │
     ▼
nrmClientes
```

## Chaves de negócio

As chaves de negócio são definidas no Schema e utilizadas durante a normalização.

Exemplo:

```text
Chave de negócio:
CPF
```

A NRM utiliza essa definição para manter a consistência dos dados destinados às etapas posteriores.

## Regra arquitetural

A NRM não deve substituir a validação.

A qualidade é avaliada na camada QA; a NRM trabalha sobre os dados selecionados para continuidade do pipeline.

---

# 9. DIM / FATO — Modelo Dimensional

A camada final transforma os dados normalizados em estruturas destinadas à análise.

Exemplos:

```text
dimClientes
dimProdutos
dimCalendario

fatoVendas
```

## Dimensões

As dimensões representam entidades utilizadas para análise.

Exemplos:

```text
dimClientes
dimProdutos
dimCalendario
```

## Fatos

As tabelas fato representam eventos ou medidas analíticas.

Exemplo:

```text
fatoVendas
```

## Operações típicas

Essa camada pode envolver:

* criação de chaves substitutas;
* ordenação;
* relacionamentos;
* `Table.NestedJoin`;
* preparação para o modelo de dados.

A modelagem dimensional não deve conter tratamentos genéricos que poderiam ter sido executados nas etapas anteriores.

---

# 10. Fluxo de Dependências

A dependência padrão entre as camadas é:

```text
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
```

Cada camada deve consumir uma saída semanticamente compatível com sua responsabilidade.

## Regra

O fluxo deve permanecer previsível.

Por exemplo, uma dimensão não deve depender diretamente de uma fonte bruta quando os dados deveriam passar pelo pipeline de preparação, tratamento, qualidade e normalização.

Essa disciplina reduz dependências ocultas e facilita manutenção e diagnóstico.

---

# 11. Arquitetura de Configuração

A configuração do framework é baseada em diferentes componentes de metadados.

Entre os componentes estão:

```text
srcSchema
stgSchema
cfgSchema
cfgPipeline
cfgTiposDados
cfgTiposBooleanos
cfgParametrosSeveridades
```

O Schema contém a definição declarativa.

O pipeline compilado transforma essas definições em estruturas otimizadas para execução.

---

## 11.1 Schema

O Schema descreve o comportamento esperado para as tabelas.

Entre suas informações estão:

```text
Tabela
Coluna
Tipo
Obrigatório
Tratamentos
Validações
Chaves de negócio
Ordem
```

A estrutura detalhada do Schema está documentada em:

`docs/schema.md`

---

## 11.2 Pipeline compilado

O Schema não é utilizado diretamente em todas as etapas de execução.

O framework compila suas definições em uma estrutura de execução.

Conceitualmente:

```text
Schema
   │
   ▼
Compilação
   │
   ▼
cfgPipeline
   │
   ├── Ordem
   ├── TiposPorColuna
   ├── TratamentosPorColuna
   ├── ValidaçõesPorColuna
   └── ChavesNegocio
```

Essa separação permite que as configurações sejam preparadas antes da execução operacional.

O funcionamento interno dessa compilação é detalhado em:

`docs/pipeline.md`

---

# 12. Metadados, Engine e Dados

A arquitetura pode ser entendida em três grandes grupos.

## 12.1 Metadados

Contêm as definições utilizadas pelo framework:

```text
Schema
Configurações
Tipos
Parâmetros
Pipeline
```

---

## 12.2 Engine

Contém as funções responsáveis pelo processamento:

```text
fxStgAplicar
fxTrnAplicar
fxQaValidar
fxNrmAplicar
fxPipeline
```

Também fazem parte desse grupo os mecanismos de compilação e execução dos operadores.

---

## 12.3 Dados

São as estruturas efetivamente processadas:

```text
src*
stg*
trn*
qa*
nrm*
dim*
fato*
```

A relação entre os três grupos pode ser representada como:

```text
┌──────────────────────────────────────┐
│              METADADOS               │
│                                      │
│ Schema / Configurações / Pipeline    │
└───────────────────┬──────────────────┘
                    │
                    ▼
┌──────────────────────────────────────┐
│                ENGINE                │
│                                      │
│ Funções / Operadores / Execução      │
└───────────────────┬──────────────────┘
                    │
                    ▼
┌──────────────────────────────────────┐
│                 DADOS                │
│                                      │
│ SRC / STG / TRN / QA / NRM / DIM     │
└──────────────────────────────────────┘
```

Essa separação é um dos fundamentos da arquitetura.

---

# 13. Pipeline e Operadores

Os operadores definidos no Schema são compilados antes da execução.

Conceitualmente:

```text
Definição
   │
   ▼
Operador
   │
   ▼
Parâmetros
   │
   ▼
Resolução em cfgOperadores
   │
   ▼
Operador executável
```

Os tratamentos e validações podem ser combinados com operadores padrão associados ao tipo da coluna.

A documentação detalhada dos operadores está disponível em:

`docs/operators.md`

---

# 14. Regras Arquiteturais

As seguintes regras devem orientar a evolução do framework.

### SRC

Deve representar a origem dos dados.

### STG

Deve cuidar da estrutura e da tipagem inicial.

### TRN

Deve aplicar tratamentos e padronizações.

### QA

Deve avaliar e registrar a qualidade dos dados.

### NRM

Deve normalizar e deduplicar os dados selecionados.

### DIM/FATO

Deve preparar os dados para consumo analítico.

---

## 14.1 Regras de separação

```text
STG ≠ Tratamento de negócio

TRN ≠ Validação

QA ≠ Transformação

NRM ≠ Validação

DIM/FATO ≠ Limpeza genérica
```

Essas separações existem para preservar a previsibilidade do pipeline e evitar que responsabilidades sejam distribuídas arbitrariamente entre consultas.

---

# 15. Tratamento de Erros e Rastreabilidade

O framework utiliza estruturas específicas para registrar problemas encontrados durante o processamento.

Na etapa STG:

```text
_STG_Ocorrencias
```

Na etapa QA:

```text
_QA_Status
_QA_Ocorrencias
```

Esse mecanismo permite manter a rastreabilidade dos problemas sem depender exclusivamente de erros fatais do Power Query.

A arquitetura favorece, portanto, o seguinte fluxo:

```text
Processamento
      │
      ▼
Identificação do problema
      │
      ▼
Registro da ocorrência
      │
      ▼
Classificação
      │
      ├── OK
      ├── AVISO
      └── ERRO
```

---

# 16. Performance como Princípio Arquitetural

A arquitetura foi projetada para reduzir processamento desnecessário.

Entre os princípios utilizados estão:

* processamento por coluna;
* uso de operações nativas do Power Query;
* compilação prévia dos pipelines;
* avaliação lazy;
* bufferização seletiva;
* redução de reconstruções desnecessárias de tabelas.

Por exemplo, o framework prioriza operações como:

```text
Table.TransformColumns
```

em vez de implementar processamento manual de cada registro sempre que uma operação nativa puder executar a mesma tarefa.

Os detalhes e recomendações de performance estão documentados em:

`docs/performance.md`

---

# 17. Modelo Conceitual Completo

A arquitetura completa pode ser representada da seguinte forma:

```text
                         CONFIGURAÇÃO
                              │
                 ┌────────────┴────────────┐
                 │                         │
              Schema                  Parâmetros
                 │                         │
                 └────────────┬────────────┘
                              │
                              ▼
                       PIPELINE COMPILADO
                              │
                              ▼
┌─────────┐             ┌─────────┐
│   SRC   │ ──────────► │   STG   │
└─────────┘             └────┬────┘
                             │
                             ▼
                        ┌─────────┐
                        │   TRN   │
                        └────┬────┘
                             │
                             ▼
                        ┌─────────┐
                        │   QA    │
                        └────┬────┘
                             │
                    ┌────────┴────────┐
                    │                 │
                    ▼                 ▼
                Ocorrências       Registros
                                    válidos
                                      │
                                      ▼
                                ┌─────────┐
                                │   NRM   │
                                └────┬────┘
                                     │
                                     ▼
                              ┌─────────────┐
                              │ DIM / FATO  │
                              └─────────────┘
```

---

# 18. Evolução da Arquitetura

Novos componentes devem preservar as fronteiras existentes.

Ao adicionar uma nova funcionalidade, deve-se determinar primeiro:

1. Qual camada é responsável?
2. A funcionalidade é configuração ou execução?
3. Deve ser representada no Schema?
4. É um tratamento, validação ou transformação estrutural?
5. Produz dados ou metadados?
6. Deve gerar uma ocorrência?
7. Altera alguma regra de dependência entre camadas?

Essa análise evita incorporar novas funcionalidades diretamente às consultas existentes sem uma responsabilidade arquitetural clara.

---

# 19. Documentação Relacionada

Para aprofundamento:

* [Getting Started](getting-started.md) — configuração e primeiro pipeline.
* [Pipeline](pipeline.md) — compilação e execução do pipeline.
* [Schema](schema.md) — estrutura e configuração do Schema.
* [Operators](operators.md) — tratamentos e validações disponíveis.
* [Performance](performance.md) — estratégias de otimização.
* [Examples](examples.md) — cenários completos de utilização.
* [Troubleshooting](troubleshooting.md) — diagnóstico de problemas.

---

# 20. Resumo

O Excel Data Framework DBB pode ser resumido em quatro conceitos:

```text
1. CAMADAS

SRC → STG → TRN → QA → NRM → DIM/FATO


2. CONFIGURAÇÃO

Schema → define o comportamento


3. COMPILAÇÃO

Schema → Pipeline → Execução


4. RESPONSABILIDADES

Dados → Engine → Metadados
```

A arquitetura busca separar **o que deve ser feito** de **como o processamento é executado**, permitindo que o comportamento do pipeline seja configurado por metadados e reutilizado entre diferentes tabelas e projetos.
