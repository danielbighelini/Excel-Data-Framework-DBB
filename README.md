# Excel Data Framework DBB - Documentação Técnica

## 📋 Sumário Executivo

Este documento descreve o Excel Data Framework DBB, um framework Power Query organizado em camadas que padroniza extração, staging, transformação, validação e modelagem dimensional.

---

## 🏗️ Arquitetura da Solução

### Fluxo Completo de Dados

```
┌─────────────────────────────────────────────────────────┐
│  SRC (Source / Extração)                                │
│  └─ srcClientes, srcProdutos, srcVendas, etc.           │
│  └─ Dados brutos do Excel / Banco / API                 │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  STG (Staging / Estruturação)                           │
│  └─ stgClientes, stgProdutos, stgVendas                 │
│  └─ Preparação: remoção de linhas vazias, normalização  │
│  └─ Conversão inicial de tipos e reordenação de colunas │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  TRN (Transformação)                                    │
│  └─ trnClientes, trnProdutos, trnVendas                 │
│  └─ Tratamentos: TRIM, UPPER, LOWER, CLEAN, etc.        │
│  └─ Aplicações de regras de formatação e limpeza        │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  QA (Quality Assurance / Validação)                     │
│  └─ qaClientes, qaProdutos, qaVendas                    │
│  └─ Validações de negócio e formato                     │
│  └─ Marcação de registros via `_QA_Status` e `_QA_Ocorrencias` │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  NRM (Normalização)                                     │
│  └─ nrmClientes, nrmProdutos, nrmVendas                 │
│  └─ Deduplicação por chaves de negócio                  │
│  └─ Preparação para dimensões e fatos                   │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  DIM / FATO (Modelo Dimensional)                        │
│  └─ dimClientes, dimProdutos, dimCalendario             │
│  └─ fatoVendas                                          │
│  └─ Tabelas prontas para análise e BI                   │
└─────────────────────────────────────────────────────────┘
```

---

## 🎯 Responsabilidades por Camada

### 1️⃣ STAGING (STG)

**Objetivo:** executar preparação estrutural e conversões iniciais de tipo.

**Principais funções:**

- `fxStgPreparar(tabela, ignorarColunas)`
  - Remove linhas completamente vazias.
  - Normaliza nomes de colunas (trim).
  - Valida estrutura de colunas (nenhum nome vazio ou duplicado).
  - Preserva a tabela como fonte para maximizar avaliação lazy.

- `fxStgAplicar(Tabela, Schema, ignorarColunas)`
  - Chama `fxStgPreparar()`.
  - Compila pipeline de schema via `fxPipeline(Schema)`.
  - Aplica tipos básicos por coluna usando `Table.TransformColumns`.
  - Reordena colunas conforme `Pipeline[Ordem]`.
  - Registra ocorrências de tipo em `_STG_Ocorrencias` sem transformar valores inválidos em `null`.

**Saída típica:**
- Tabela normalizada e estruturada.
- Conversão de tipos básicos aplicada.
- `_STG_Ocorrencias` com falhas de tipo.
- Colunas ordenadas.

**Exemplo:**
```powerquery
stgClientes = let
    Fonte = srcClientes,
    Preparada = fxStgAplicar(Fonte, "tbClientes")
in
    Preparada;
```

---

### 2️⃣ TRANSFORMAÇÃO (TRN)

**Objetivo:** aplicar tratamentos de dados e ajustes de formatação.

**Principais funções:**

- `fxTrnAplicar(Tabela, Schema)`
  - Compila operadores de tratamento definidos no schema.
  - Executa transformações de texto, numéricas e de conversão.
  - Permite encadear tratamentos em uma única passagem.

**Operadores comuns:**
- Texto: `TRIM`, `UPPER`, `LOWER`, `PROPER`, `CLEAN`, `SINGLESPACE`, `DIGITS`, `ALPHANUMERIC`, `NORMALIZEBASIC`.
- Numérico: `ABS`, `ROUND`.
- Conversão/extração: `NUMBER`, `CPF`, `CNPJ`, `CEP`.

**Saída típica:**
- Valores tratados e normalizados.
- Formatos padronizados conforme o schema.
- Tabela pronta para validação.

**Exemplo:**
```powerquery
trnClientes = let
    Fonte = stgClientes,
    Transformada = fxTrnAplicar(Fonte, "tbClientes")
in
    Transformada;
```

---

### 3️⃣ QUALIDADE (QA)

**Objetivo:** validar dados e marcar problemas sem excluir registros.

**Principais funções:**

- `fxQaValidar(Tabela, Schema)`
  - Compila operadores de validação por coluna.
  - Adiciona `_QA_Ocorrencias` com lista de problemas.
  - Define `_QA_Status` como `OK`, `AVISO` ou `ERRO`.
  - Não remove registros — apenas marca.

- `fxQaFiltrarPorStatus(Tabela, Status)`
  - Filtra registros por status desejado.
  - Remove colunas de controle QA quando necessário.

- `fxQaExtrairProblemas(Tabela)`
  - Retorna somente registros com problemas.
  - Útil para auditoria, correção e reporte.

**Validadores disponíveis:**
- `REQUIRED`
- `CPFVAL`, `CNPJVAL`, `CEPVAL`
- `EMAIL`, `URL`, `DOMAIN`
- `LIST`, `MIN`, `MAX`, `INTERVAL`, `SIZE`

**Saída típica:**
- Colunas `_QA_Status` e `_QA_Ocorrencias`.
- Lista de ocorrências com código, severidade, coluna, tipo e descrição.
- Registro original preservado.

**Exemplo:**
```powerquery
qaClientes = let
    Fonte = trnClientes,
    qa = fxQaValidar(Fonte, "tbClientes")
in
    qa;

ClientesValidos = fxQaFiltrarPorStatus(qa, "OK");
ProblemasCliente = fxQaExtrairProblemas(qa);
```

---

### 4️⃣ NORMALIZAÇÃO (NRM)

**Objetivo:** deduplicar dados e aplicar chaves de negócio.

**Principais funções:**

- `fxNrmAplicar(Tabela, Schema)`
  - Remove duplicatas com base em chaves de negócio definidas no schema.
  - Garante que a tabela permaneça consistente para modelagem.

**Saída típica:**
- Dados deduplicados segundo o schema.
- Tabela preparada para produção de dimensões e fatos.

**Exemplo:**
```powerquery
nrmClientes = let
    Fonte = qaClientes,
    Valida = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Valida, "tbClientes")
in
    Normalizada;
```

---

### 5️⃣ MODELO DIMENSIONAL (DIM / FATO)

**Objetivo:** construir tabelas analíticas finais.

**Componentes principais:**
- Dimensões: `dimClientes`, `dimProdutos`, `dimCalendario`.
- Fato: `fatoVendas`.

**Características:**
- Chaves substitutas adicionadas via `Table.AddIndexColumn`.
- Reordenação de colunas para estabilidade.
- Bufferização para performance em queries dependentes.
- Relacionamentos realizados via `Table.NestedJoin`.

**Exemplo:**
```powerquery
dimClientes = let
    Fonte = nrmClientes,
    Chaves = Table.AddIndexColumn(Fonte, "IDCliente", 1, 1, Int64.Type),
    Reordenada = Table.ReorderColumns(Chaves,
        {"IDCliente", "CPF", "Nome", "DataNascimento", "Cidade", "Estado"})
in
    Table.Buffer(Reordenada);
```

---

## 🔧 Componentes de Configuração

- `srcSchema` / `stgSchema`
  - Schema ativo do projeto.
  - Define tabelas, colunas, tipos, tratamentos, validações e chaves.

- `fxPipeline(Schema)`
  - Compila o schema em pipelines de tipos, tratamentos, validações e ordenação.

- `cfgTiposDados`
  - Mapeia nomes de tipo (`TEXTO`, `DATA`, `NÚMERO DECIMAL`, etc.) para tipos M.

- `cfgTiposBooleanos`
  - Converte valores lógicos como `TRUE/FALSE`, `SIM/NÃO`, `1/0`.

- `cfgParametrosSeveridades`
  - Controla se um problema é `AVISO` ou `ERRO`.

---

## ⚙️ Como o Pipeline Funciona

1. `fxPipeline(Schema)` monta:
   - `TiposPorColuna`
   - `TratamentosPorColuna`
   - `ValidaçõesPorColuna`
   - `Ordem`
   - `ChavesNegocio`

2. `fxStgAplicar()` prepara a tabela e aplica tipos básicos.
3. `fxTrnAplicar()` aplica transformações de tratamento.
4. `fxQaValidar()` valida e marca ocorrências.
5. `fxNrmAplicar()` deduplica antes da modelagem dimensional.

---

## 📌 Observações de Performance

- O STG usa `Table.TransformColumns` para converter tipos por coluna.
- A validação de tipo no STG registra falhas em `_STG_Ocorrencias` sem silenciar valores inválidos.
- O QA agrega problemas em `_QA_Ocorrencias` e define `_QA_Status`.
- A arquitetura evita passagens de dados desnecessárias, mantendo separação clara entre etapas.

---

## 📂 Objetos de Dados de Exemplo

- `srcClientes`, `srcProdutos`, `srcVendas` — fontes de dados.
- `stgClientes`, `stgProdutos`, `stgVendas` — preparação de staging.
- `trnClientes`, `trnProdutos`, `trnVendas` — transformações aplicadas.
- `qaClientes`, `qaProdutos`, `qaVendas` — validação de qualidade.
- `nrmClientes`, `nrmProdutos`, `nrmVendas` — normalização.
- `dimClientes`, `dimProdutos`, `dimCalendario`, `fatoVendas` — modelo analítico final.

---

## 🔧 Compiladores de Operadores

### Compilação de Tratamentos

`fxPipelineCompilarOperadores` resolve cada operador do schema em um registro de execução:
- `Código` do operador
- `Parâmetros` extraídos de `CODIGO(p1,p2)`
- lookup em `cfgOperadores`

`fxPipelineCompilarColuna` combina:
- operadores padrão de tratamento para o tipo da coluna (`fxOperadoresPadrao(Tipo)[Tratamentos]`)
- operadores de tratamento definidos no schema (`Definição[Tratamentos] ?? {}`)

O resultado é `TratamentosPorColuna`, um record onde cada coluna aponta para uma lista de operadores compilados.

### Compilação de Validações

`fxPipelineCompilarColuna` também monta o pipeline de validação por coluna:
- `REQUIRED` implícito quando `Obrigatório = true`
- validações padrão por tipo (`fxOperadoresPadrao(Tipo)[Validações]`)
- validações declaradas no schema (`Definição[Validações] ?? {}`)

O resultado é `ValidaçõesPorColuna`, um record onde cada coluna aponta para uma lista de validadores compilados.

---

## 📊 Schema e Pipeline

### Estrutura do Schema

```powerquery
// tbSchema (Excel)
Tabela      | Coluna           | Tipo    | Obrigatório | Tratamentos        | Validações
tbClientes  | CPF              | TEXT    | SIM         | TRIM;UPPER;DIGITS  | CPFVAL
tbClientes  | Nome             | TEXT    | SIM         | TRIM;PROPER        | SIZE(100)
tbClientes  | DataNascimento   | DATE    | NÃO         |                    |
tbClientes  | Cidade           | TEXT    | NÃO         | TRIM;PROPER        |
tbClientes  | Estado           | TEXT    | NÃO         | TRIM;UPPER         |
```

### Pipeline Compilado

`cfgPipeline` aplica `fxPipelineCompilar` a cada tabela do schema e cria um pipeline simplificado com os campos realmente usados pelo framework.

**Campos do pipeline:**
- `Ordem`
- `TiposPorColuna`
- `TratamentosPorColuna`
- `ValidaçõesPorColuna`
- `ChavesNegocio`

```powerquery
cfgPipeline[tbClientes] =
[
    Ordem = {"CPF","Nome","DataNascimento","Cidade","Estado"},
    TiposPorColuna = [CPF=type text, Nome=type text, DataNascimento=type date, Cidade=type text, Estado=type text],
    TratamentosPorColuna = [CPF={...}, Nome={...}, ...],
    ValidaçõesPorColuna = [CPF={...}, Nome={...}, ...],
    ChavesNegocio = {"CPF"}
]
```

---

## 🧩 Métodos de Tratamento Disponíveis

### Tratamentos básicos
- `TRIM` → `fxTratamentoTrim`
- `UPPER` → `fxTratamentoUpper`
- `LOWER` → `fxTratamentoLower`
- `PROPER` → `fxTratamentoProper`
- `CLEAN` → `fxTratamentoClean`
- `EMPTYTONULL` → `fxTratamentoEmptyToNull`
- `NULLTOEMPTY` → `fxTratamentoNullToEmpty`
- `SINGLESPACE` → `fxTratamentoSingleSpace`
- `DIGITS` → `fxTratamentoDigits`
- `ALPHANUMERIC` → `fxTratamentoAlphaNumeric`
- `ABS` → `fxTratamentoAbs`
- `ROUND` → `fxTratamentoRound`
- `NORMALIZEBASIC` → `fxTratamentoNormalizeBasic`
- `NUMBER` → `fxTratamentoNumber`
- `CPF` → `fxTratamentoCPF`
- `CNPJ` → `fxTratamentoCNPJ`
- `CEP` → `fxTratamentoCEP`

### Tratamentos de texto avançados
- `REPLACE` → `fxTratamentoReplace`
- `LEFT` → `fxTratamentoLeft`
- `RIGHT` → `fxTratamentoRight`
- `MID` → `fxTratamentoMid`
- `BEFORE` → `fxTratamentoBefore`
- `AFTER` → `fxTratamentoAfter`
- `ADDPREFIX` → `fxTratamentoAddPrefix`
- `ADDSUFFIX` → `fxTratamentoAddSuffix`
- `REMOVEPREFIX` → `fxTratamentoRemovePrefix`
- `REMOVESUFFIX` → `fxTratamentoRemoveSuffix`
- `PADLEFT` → `fxTratamentoPadLeft`
- `PADRIGHT` → `fxTratamentoPadRight`
- `REMOVECHARS` → `fxTratamentoRemoveChars`
- `KEEPCHARS` → `fxTratamentoKeepChars`
- `REMOVEACCENTS` → `fxTratamentoRemoveAccents`
- `REMOVEPUNCTUATION` → `fxTratamentoRemovePunctuation`
- `KEEPTEXT` → `fxTratamentoKeepText`

---

## ✅ Métodos de Validação Disponíveis

- `REQUIRED` → `fxValidacaoREQUIRED`
- `LIST` → `fxValidacaoList`
- `DOMAIN` → `fxValidacaoDomain`
- `SIZE` → `fxValidacaoSize`
- `MIN` → `fxValidacaoMin`
- `MAX` → `fxValidacaoMax`
- `INTERVAL` → `fxValidacaoInterval`
- `EMAIL` → `fxValidacaoEmail`
- `URL` → `fxValidacaoURL`
- `CEPVAL` → `fxValidacaoCEP`
- `CPFVAL` → `fxValidacaoCPF`
- `CNPJVAL` → `fxValidacaoCNPJ`

---

## 🚀 Otimizações Implementadas

### 1. Processamento de tabelas otimizado

**Depois (Otimizado):**
```
STG (prepare) → STG (tipos) → TRN (transformações) → 
QA (validações) → NRM (normalização)
```
- `fxStgAplicar` faz preparação estrutural e aplicação de tipos básicos em uma sequência única por tabela.
- `fxTrnAplicar` compila funções de tratamento por coluna e aplica com `Table.TransformColumns`.
- `fxQaValidar` cria um único campo `_QA` com status e ocorrências, expandido apenas no final.
- `fxNrmAplicar` faz deduplicação por chaves de negócio com `Table.Distinct`.

### 2. Compilação de Operadores

- `cfgSchema` deriva de `stgSchema` e agrupa definições por tabela.
- `cfgPipeline` compila o schema em um pipeline enxuto com campos usados pelo framework.
- `fxPipelineCompilarColuna` combina operadores padrão e schema, incluindo `REQUIRED` implícito.
- Pipelines por coluna são pré-compilados para reduzir overhead durante a execução.

### 3. Bufferização Estratégica

```powerquery
Table.Buffer(stgSchema)
List.Buffer(...)
```
- Buffer apenas em estruturas de metadados e listas de operadores.
- Evita `Table.Buffer` em cada tabela de dados operacional, mantendo o fold quando possível.

### 4. QA em uma única passagem

- `fxQaValidar` percorre cada linha apenas uma vez e acumula ocorrências em um registro `_QA`.
- O status final (`OK`, `AVISO`, `ERRO`) é definido na mesma passagem.
- `fxQaFiltrarPorStatus` remove colunas de controle após a validação.

### 5. Lazy Evaluation e coluna a coluna

- Cada transformação é aplicada somente às colunas presentes no schema.
- Colunas com `type any` não recebem transformações de tipo desnecessárias.
- O fluxo prioriza reconstrução mínima de tabelas para manter o desempenho.

---

## ✅ Checklist de Implementação

### Fase 1: Configuração
- [ ] Criar tabelas no Excel: tbParametros, tbSchema, etc.
- [ ] Definir schema (Tratamentos, Validações)
- [ ] Carregar fonte de dados (srcClientes, etc.)

### Fase 2: Testes Unitários
- [ ] Testar fxStgAplicar (STG)
- [ ] Testar fxTrnAplicar (TRN)
- [ ] Testar fxQaValidar (QA)
- [ ] Testar fxNrmAplicar (NRM)

### Fase 3: Pipeline Completo
- [ ] Validar stgClientes → trnClientes → qaClientes → nrmClientes
- [ ] Validar dimClientes → fatoVendas
- [ ] Testar com volume de dados real

### Fase 4: Monitoramento
- [ ] Extrair relatório de problemas (fxQaExtrairProblemas)
- [ ] Monitorar performance
- [ ] Validar qualidade dos dados dimensional

---

## 🔐 Boas Práticas Implementadas

1. **Separação de Responsabilidades**
   - Cada camada faz UMA coisa bem
   - Fácil manutenção e debugging

2. **Tratamento de Erros**
   - Try-catch em conversões de tipo
   - Validação de schema
   - Erros informativos

3. **Rastreabilidade**
   - Colunas de controle QA (_QA_Status, _QA_Ocorrencias)
   - Audible trail de transformações
   - Possibilidade de rejeição seletiva

4. **Reutilização**
   - Funções genéricas e parametrizadas
   - Schema-driven architecture
   - Aplicável a qualquer tabela

5. **Documentação**
   - Comentários claros nas seções
   - Nomes descritivos de funções
   - Schema centralizado

---

## 📚 Referências

### ETL/ELT Best Practices
- Separation of Concerns
- Single Responsibility Principle
- DRY (Don't Repeat Yourself)
- KISS (Keep It Simple, Stupid)

### Power Query Otimizações
- Table.Buffer() com propósito
- Lazy evaluation
- Compilação de pipelines
- Operações de "fold" (accumulate)

### Data Quality Frameworks
- Great Expectations
- DAMA Data Management Framework
- Six Sigma DMAIC (Define-Measure-Analyze-Improve-Control)

---

## 🎓 Exemplos de Uso

### Exemplo 1: Pipeline Completo para Clientes

```powerquery
// Extração
stgClientes = fxStgAplicar(srcClientes, "tbClientes");

// Transformação
trnClientes = fxTrnAplicar(stgClientes, "tbClientes");

// Validação
qaClientes = fxQaValidar(trnClientes, "tbClientes");

// Normalização (apenas dados válidos)
ClientesValidos = fxQaFiltrarPorStatus(qaClientes, "OK");
nrmClientes = fxNrmAplicar(ClientesValidos, "tbClientes");

// Dimensão
dimClientes = Table.AddIndexColumn(
    Table.Sort(Table.Distinct(nrmClientes), {"CPF"}),
    "IDCliente", 1, 1, Int64.Type
);
```

### Exemplo 2: Auditoria de Problemas

```powerquery
// Extrair registros com erros
ProblemasDetectados = fxQaExtrairProblemas(qaClientes);

// Ver detalhes dos problemas
ProblemasComDetalhes = Table.ExpandRecordColumn(
    ProblemasDetectados,
    "_QA_Ocorrencias",
    {"Coluna", "Mensagem", "Detalhes"}
);
```

### Exemplo 3: Modelo Dimensional Completo

```powerquery
// Fato com relacionamentos
fatoVendas = let
    Vendas = nrmVendas,
    ComCliente = Table.NestedJoin(Vendas, {"CPF"}, dimClientes, {"CPF"}, "_C", JoinKind.Inner),
    ComProduto = Table.NestedJoin(ComCliente, {"CódigoProduto"}, dimProdutos, {"Código"}, "_P", JoinKind.Inner),
    Expandida = Table.ExpandTableColumn(
        Table.ExpandTableColumn(ComProduto, "_C", {"IDCliente"}),
        "_P",
        {"IDProduto"}
    )
in
    Table.SelectColumns(Expandida, {"Data", "IDCliente", "IDProduto", "Quantidade", "Valor"})
```

---

## 🐛 Troubleshooting

### Problema: Erro na Compilação do Schema
**Solução:** Verifique se o operador está cadastrado em tbParametrosTratamentos ou tbParametrosValidacoes

### Problema: Muitos registros com status AVISO
**Solução:** Usar fxQaExtrairProblemas para auditar, ajustar validações se necessário

### Problema: Performance ruim
**Solução:** Verificar se há múltiplos Table.Buffer desnecessários, consolidar operações

### Problema: Dados perdidos na normalização
**Solução:** Usar fxQaFiltrarPorStatus para passar apenas "OK", ou ajustar validações para AVISO


