# Pipeline

## 1. Visão Geral

O Excel Data Framework DBB utiliza um pipeline orientado por Schema para transformar definições declarativas em estruturas de execução utilizadas pelas diferentes etapas do framework.

O fluxo conceitual é:

```text
Schema
   │
   ▼
Compilação
   │
   ▼
cfgPipeline
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

O pipeline tem duas responsabilidades principais:

1. **compilar as configurações do Schema**;
2. **disponibilizar essas configurações para execução pelas camadas do framework**.

A separação entre compilação e execução evita que as regras do Schema precisem ser interpretadas repetidamente durante o processamento dos dados.

---

# 2. Conceito de Pipeline

O pipeline representa o conjunto de regras necessárias para processar uma tabela.

Para cada tabela, ele pode conter informações como:

```text
Ordem
TiposPorColuna
TratamentosPorColuna
ValidaçõesPorColuna
ChavesNegocio
```

Exemplo conceitual:

```text
tbClientes
    │
    ▼
Pipeline
    │
    ├── Ordem
    ├── TiposPorColuna
    ├── TratamentosPorColuna
    ├── ValidaçõesPorColuna
    └── ChavesNegocio
```

Essas informações são derivadas do Schema e organizadas em uma estrutura adequada para execução.

---

# 3. Schema → Pipeline

O Schema representa a configuração declarativa.

Por exemplo:

```text
Tabela      | Coluna         | Tipo | Obrigatório | Tratamentos | Validações
tbClientes  | CPF            | TEXT | SIM         | TRIM;DIGITS | CPFVAL
tbClientes  | Nome           | TEXT | SIM         | TRIM;PROPER | SIZE(100)
tbClientes  | Estado         | TEXT | NÃO         | TRIM;UPPER  | LIST(RS,SP,SC)
```

O pipeline transforma essas definições em estruturas específicas por coluna.

Conceitualmente:

```text
Schema
  │
  ├── CPF
  │    ├── Tipo: TEXT
  │    ├── Obrigatório: SIM
  │    ├── Tratamentos: TRIM;DIGITS
  │    └── Validação: CPFVAL
  │
  ├── Nome
  │    ├── Tipo: TEXT
  │    ├── Obrigatório: SIM
  │    ├── Tratamentos: TRIM;PROPER
  │    └── Validação: SIZE(100)
  │
  └── Estado
       ├── Tipo: TEXT
       ├── Tratamentos: TRIM;UPPER
       └── Validação: LIST(...)
```

---

# 4. Compilação do Pipeline

A compilação é responsável por transformar as definições do Schema em estruturas que possam ser consumidas pelas etapas de execução.

O processo pode ser representado como:

```text
Schema
   │
   ▼
fxPipeline
   │
   ▼
Compilação
   │
   ├── Tipos
   ├── Tratamentos
   ├── Validações
   ├── Ordem
   └── Chaves
   │
   ▼
cfgPipeline
```

O `cfgPipeline` representa o resultado compilado do Schema.

O README atual descreve `cfgPipeline` como uma estrutura que aplica a compilação do pipeline às tabelas do Schema e mantém os campos efetivamente utilizados pelo framework.

---

# 5. Estrutura do Pipeline Compilado

Um pipeline compilado possui, conceitualmente:

```powerquery
[
    Ordem = {...},
    TiposPorColuna = [...],
    TratamentosPorColuna = [...],
    ValidaçõesPorColuna = [...],
    ChavesNegocio = {...}
]
```

Exemplo:

```powerquery
cfgPipeline[tbClientes] =
[
    Ordem = {
        "CPF",
        "Nome",
        "DataNascimento",
        "Cidade",
        "Estado"
    },

    TiposPorColuna = [
        CPF = type text,
        Nome = type text,
        DataNascimento = type date,
        Cidade = type text,
        Estado = type text
    ],

    TratamentosPorColuna = [
        CPF = {...},
        Nome = {...}
    ],

    ValidaçõesPorColuna = [
        CPF = {...},
        Nome = {...}
    ],

    ChavesNegocio = {"CPF"}
]
```

A estrutura acima é representativa do modelo documentado no projeto.

---

# 6. Tipos por Coluna

`TiposPorColuna` associa cada coluna do Schema ao tipo M correspondente.

Exemplo:

```text
CPF            → type text
Nome           → type text
DataNascimento → type date
Cidade         → type text
Estado         → type text
```

Essa informação é utilizada principalmente pela camada STG.

```text
Pipeline
   │
   ▼
TiposPorColuna
   │
   ▼
fxStgAplicar
   │
   ▼
Tipos aplicados
```

A aplicação de tipos utiliza processamento por coluna.

O framework utiliza `Table.TransformColumns` para essa etapa.

---

# 7. Tratamentos por Coluna

`TratamentosPorColuna` representa os operadores que devem ser executados para cada coluna.

Exemplo:

```text
CPF
 ├── TRIM
 └── DIGITS

Nome
 ├── TRIM
 └── PROPER

Estado
 ├── TRIM
 └── UPPER
```

O pipeline não precisa armazenar apenas o texto original do operador.

Os operadores são compilados para estruturas de execução.

Conceitualmente:

```text
TRIM;DIGITS
     │
     ▼
Compilação
     │
     ▼
{ operador executável, parâmetros, ... }
```

Essa estrutura é então consumida pela camada TRN.

---

# 8. Validações por Coluna

`ValidaçõesPorColuna` representa os validadores associados a cada coluna.

Exemplo:

```text
CPF
 └── CPFVAL

Nome
 └── SIZE(100)

Estado
 └── LIST(RS,SP,SC)
```

Além das validações explicitamente declaradas no Schema, a compilação pode incluir:

* `REQUIRED` quando `Obrigatório = true`;
* validações padrão associadas ao tipo;
* validações declaradas no Schema.

Essa composição está documentada na implementação atual do pipeline.

Conceitualmente:

```text
Obrigatório
     +
Validações padrão
     +
Validações do Schema
     │
     ▼
ValidaçõesPorColuna
```

---

# 9. Ordem das Colunas

O campo `Ordem` representa a sequência de colunas definida pelo Schema.

Exemplo:

```text
Ordem =
{
    "CPF",
    "Nome",
    "DataNascimento",
    "Cidade",
    "Estado"
}
```

Essa informação é utilizada para manter uma estrutura de saída previsível.

A ordenação das colunas é aplicada durante a preparação da tabela.

---

# 10. Chaves de Negócio

`ChavesNegocio` representa as colunas utilizadas para identificar registros no processo de normalização.

Exemplo:

```text
ChavesNegocio = {"CPF"}
```

Em uma tabela em que a identificação depende de mais de uma coluna:

```text
ChavesNegocio =
{
    "CodigoCliente",
    "CodigoProduto"
}
```

Essas informações são consumidas pela camada NRM.

```text
Pipeline
   │
   ▼
ChavesNegocio
   │
   ▼
fxNrmAplicar
   │
   ▼
Normalização
```

---

# 11. Compilação de Operadores

Os operadores definidos no Schema são convertidos em estruturas de execução.

O processo documentado pelo projeto é:

```text
Código do operador
        │
        ▼
Extração dos parâmetros
        │
        ▼
Lookup em cfgOperadores
        │
        ▼
Registro de execução
```

O registro compilado contém o código do operador e os parâmetros extraídos.

Esse processo é realizado por `fxPipelineCompilarOperadores`.

---

# 12. Compilação por Coluna

Depois da resolução dos operadores, o pipeline monta as definições específicas de cada coluna.

Para tratamentos:

```text
Tipo da coluna
      │
      ├── Tratamentos padrão
      │
      └── Tratamentos do Schema
                │
                ▼
       TratamentosPorColuna
```

Para validações:

```text
Tipo da coluna
      │
      ├── REQUIRED implícito
      ├── Validações padrão
      └── Validações do Schema
                │
                ▼
       ValidaçõesPorColuna
```

O README atual descreve essa combinação dentro de `fxPipelineCompilarColuna`.

---

# 13. Execução do Pipeline

Depois de compilado, o pipeline é consumido pelas etapas de processamento.

O fluxo padrão é:

```text
1. STG
   │
   ▼
2. TRN
   │
   ▼
3. QA
   │
   ▼
4. NRM
```

Cada etapa utiliza apenas as partes do pipeline necessárias à sua responsabilidade.

---

## 13.1 STG

O STG utiliza principalmente:

```text
TiposPorColuna
Ordem
```

Fluxo:

```text
Fonte
  │
  ▼
fxStgAplicar
  │
  ├── Preparação estrutural
  ├── Tipagem
  └── Ordenação
  │
  ▼
STG
```

O README descreve `fxStgAplicar` como responsável por preparar a tabela, compilar/obter o pipeline, aplicar tipos e reordenar as colunas.

---

## 13.2 TRN

O TRN utiliza:

```text
TratamentosPorColuna
```

Fluxo:

```text
STG
 │
 ▼
fxTrnAplicar
 │
 ▼
Tratamentos por coluna
 │
 ▼
TRN
```

O processamento utiliza os operadores compilados para cada coluna.

---

## 13.3 QA

O QA utiliza:

```text
ValidaçõesPorColuna
```

Fluxo:

```text
TRN
 │
 ▼
fxQaValidar
 │
 ├── Validações
 ├── Ocorrências
 └── Status
 │
 ▼
QA
```

O resultado contém `_QA_Status` e `_QA_Ocorrencias`.

---

## 13.4 NRM

O NRM utiliza:

```text
ChavesNegocio
```

Fluxo:

```text
QA
 │
 ▼
Registros selecionados
 │
 ▼
fxNrmAplicar
 │
 ▼
Chaves de negócio
 │
 ▼
NRM
```

O README descreve `fxNrmAplicar` como responsável pela deduplicação baseada nas chaves de negócio definidas no Schema.

---

# 14. Pipeline Completo

O fluxo completo pode ser representado como:

```text
                         Schema
                            │
                            ▼
                      fxPipeline
                            │
                            ▼
                     cfgPipeline
                            │
             ┌──────────────┼──────────────┐
             │              │              │
             ▼              ▼              ▼
        TiposPorColuna  Tratamentos    Validações
             │              │              │
             │              │              │
             ▼              ▼              ▼
            STG            TRN             QA
             │              │              │
             └──────────────┴──────────────┘
                            │
                            ▼
                           NRM
                            │
                            ▼
                        DIM/FATO
```

A função de cada estrutura é especializada:

| Estrutura              | Consumidor principal |
| ---------------------- | -------------------- |
| `TiposPorColuna`       | STG                  |
| `Ordem`                | STG                  |
| `TratamentosPorColuna` | TRN                  |
| `ValidaçõesPorColuna`  | QA                   |
| `ChavesNegocio`        | NRM                  |

---

# 15. Pipeline por Tabela

O framework mantém o pipeline organizado por tabela.

Conceitualmente:

```text
cfgPipeline
│
├── tbClientes
│   ├── Ordem
│   ├── TiposPorColuna
│   ├── TratamentosPorColuna
│   ├── ValidaçõesPorColuna
│   └── ChavesNegocio
│
├── tbProdutos
│   ├── Ordem
│   ├── TiposPorColuna
│   ├── TratamentosPorColuna
│   ├── ValidaçõesPorColuna
│   └── ChavesNegocio
│
└── tbVendas
    ├── Ordem
    ├── TiposPorColuna
    ├── TratamentosPorColuna
    ├── ValidaçõesPorColuna
    └── ChavesNegocio
```

Isso permite que diferentes tabelas possuam configurações diferentes sem exigir pipelines específicos implementados manualmente.

---

# 16. Pipeline e Reutilização

O principal benefício da arquitetura é separar o mecanismo de execução da configuração da tabela.

Sem essa separação, uma abordagem tradicional poderia exigir:

```text
fxClientes
fxProdutos
fxVendas
fxFornecedores
...
```

Cada função teria que conter sua própria lógica.

Com o pipeline orientado por Schema:

```text
                 Schema
                    │
          ┌─────────┼─────────┐
          ▼         ▼         ▼
      Clientes   Produtos   Vendas
          │         │         │
          └─────────┼─────────┘
                    ▼
              Mesmo Engine
```

A implementação das etapas permanece reutilizável.

---

# 17. Pipeline e Operadores Padrão

O pipeline pode incorporar operadores padrão determinados pelo tipo da coluna.

O processo conceitual é:

```text
Tipo
 │
 ▼
fxOperadoresPadrao(Tipo)
 │
 ├── Tratamentos
 └── Validações
       │
       ▼
Configuração da coluna
       │
       ▼
Pipeline compilado
```

Isso permite estabelecer comportamentos padrão sem exigir que todas as configurações sejam repetidas manualmente no Schema.

---

# 18. Pipeline e Lazy Evaluation

O pipeline foi estruturado para trabalhar com o modelo de avaliação lazy do Power Query.

A arquitetura evita materializar desnecessariamente estruturas de dados durante as etapas de configuração e execução.

O framework prioriza:

* processamento somente das colunas necessárias;
* operações nativas de tabela;
* bufferização seletiva;
* compilação de estruturas de metadados;
* redução de reconstruções desnecessárias.

O README atual destaca especificamente o uso de `Table.TransformColumns`, `List.Buffer`, `Table.Buffer` em estruturas apropriadas e processamento coluna a coluna.
Os detalhes de otimização estão documentados em `performance.md`.

---

# 19. Pipeline e Ocorrências

O pipeline não trata todas as falhas da mesma maneira.

Na camada STG, problemas relacionados à tipagem podem ser registrados em:

```text
_STG_Ocorrencias
```

Na camada QA, problemas de qualidade são registrados em:

```text
_QA_Ocorrencias
```

e classificados através de:

```text
_QA_Status
```

Assim:

```text
Erro de estrutura/tipo
        │
        ▼
_STG_Ocorrencias


Problema de qualidade
        │
        ▼
_QA_Ocorrencias
        │
        ▼
_QA_Status
```

## Essa distinção mantém separados problemas estruturais e problemas de qualidade dos dados.

# 20. Ordem de Processamento

A ordem conceitual das etapas é:

```text
1. Preparação estrutural
2. Aplicação de tipos
3. Tratamentos
4. Validações
5. Seleção/classificação dos registros
6. Normalização
7. Modelagem dimensional
```

Representação:

```text
Fonte
 │
 ▼
STG
 │
 ├── Estrutura
 ├── Tipos
 └── Ordem
 │
 ▼
TRN
 │
 └── Tratamentos
 │
 ▼
QA
 │
 └── Validações
 │
 ▼
NRM
 │
 └── Normalização
 │
 ▼
DIM/FATO
```

O pipeline centraliza as definições necessárias para que cada etapa execute sua responsabilidade.

---

# 21. Exemplo Completo

Considere:

```text
Tabela: tbClientes

CPF
Nome
DataNascimento
Cidade
Estado
```

Schema:

```text
CPF
  Tipo = TEXT
  Obrigatório = SIM
  Tratamentos = TRIM;DIGITS
  Validações = CPFVAL

Nome
  Tipo = TEXT
  Obrigatório = SIM
  Tratamentos = TRIM;PROPER
  Validações = SIZE(100)

Estado
  Tipo = TEXT
  Obrigatório = NÃO
  Tratamentos = TRIM;UPPER
  Validações = LIST(RS,SP,SC)

Chave de negócio:
CPF
```

Após a compilação:

```text
cfgPipeline[tbClientes]
│
├── Ordem
│
├── TiposPorColuna
│
│   ├── CPF → text
│   ├── Nome → text
│   ├── DataNascimento → date
│   ├── Cidade → text
│   └── Estado → text
│
├── TratamentosPorColuna
│
│   ├── CPF → TRIM, DIGITS
│   ├── Nome → TRIM, PROPER
│   └── Estado → TRIM, UPPER
│
├── ValidaçõesPorColuna
│
│   ├── CPF → REQUIRED, CPFVAL
│   ├── Nome → REQUIRED, SIZE
│   └── Estado → LIST
│
└── ChavesNegocio
    └── CPF
```

Esse pipeline pode então ser utilizado por:

```text
fxStgAplicar
      ↓
fxTrnAplicar
      ↓
fxQaValidar
      ↓
fxNrmAplicar
```

---

# 22. Relação com as Camadas

A relação entre pipeline e arquitetura de dados é:

```text
                    cfgPipeline
                         │
        ┌────────────────┼────────────────┐
        │                │                │
        ▼                ▼                ▼
       STG              TRN               QA
        │                │                │
 TiposPorColuna   TratamentosPorColuna   ValidaçõesPorColuna
        │                │                │
        └────────────────┼────────────────┘
                         │
                         ▼
                        NRM
                         │
                    ChavesNegocio
                         │
                         ▼
                      DIM/FATO
```

O pipeline funciona, portanto, como uma ponte entre a configuração declarativa e a execução operacional.

---

# 23. Responsabilidades do Pipeline

O pipeline é responsável por:

* interpretar as definições do Schema;
* compilar operadores;
* organizar configurações por coluna;
* determinar tipos;
* determinar tratamentos;
* determinar validações;
* determinar a ordem das colunas;
* disponibilizar chaves de negócio;
* preparar estruturas para execução.

O pipeline **não é responsável por**:

* extrair dados da fonte;
* executar diretamente a modelagem dimensional;
* substituir as responsabilidades das camadas STG, TRN, QA ou NRM;
* armazenar os dados operacionais.

---

# 24. Evolução do Pipeline

Alterações no pipeline devem preservar a separação entre:

```text
Configuração
      │
      ▼
Compilação
      │
      ▼
Execução
```

Ao adicionar um novo componente, deve-se avaliar se ele pertence a:

1. definição do Schema;
2. compilação;
3. estrutura do pipeline;
4. execução de uma camada;
5. operador;
6. configuração auxiliar.

Essa separação evita incorporar lógica de execução diretamente às estruturas de configuração.

---

# 25. Documentação Relacionada

* [Getting Started](getting-started.md) — configuração do primeiro pipeline.
* [Architecture](architecture.md) — arquitetura e responsabilidades das camadas.
* [Schema](schema.md) — definição declarativa utilizada pelo pipeline.
* [Operators](operators.md) — tratamentos e validações.
* [Performance](performance.md) — otimizações do processamento.
* [Examples](examples.md) — exemplos completos.
* [Troubleshooting](troubleshooting.md) — diagnóstico de problemas.

---

# 26. Resumo

O pipeline pode ser resumido em três etapas conceituais:

```text
        DEFINIR
           │
           ▼
        Schema
           │
           ▼
       COMPILAR
           │
           ▼
      cfgPipeline
           │
           ▼
        EXECUTAR
           │
           ▼
   STG → TRN → QA → NRM
```

O princípio central é:

> **O Schema define o comportamento; o pipeline compila esse comportamento; as camadas executam as regras sobre os dados.**

Essa separação permite reutilizar o mesmo mecanismo de processamento para diferentes tabelas, mantendo as particularidades de cada conjunto de dados na configuração.
