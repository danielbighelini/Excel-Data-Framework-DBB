# Excel Data Framework DBB - Documentação Técnica

## 📋 Sumário Executivo

Este documento descreve o Excel Data Framework DBB, com separação clara de responsabilidades em 5 camadas distintas.

---

## 🏗️ Arquitetura da Solução

### Fluxo Completo de Dados

```
┌─────────────────────────────────────────────────────────┐
│  SRC (Source/Extração)                                  │
│  └─ srcClientes, srcProdutos, srcVendas, etc.           │
│  └─ Dados brutos do Excel/BD/APIs                       │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  STG (Staging/Estruturação)                             │
│  └─ stgClientes, stgProdutos, stgVendas                 │
│  └─ Preparação: remoção linhas vazias, normalização     │
│  └─ Aplicação de tipos básicos                          │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  TRN (Transform/Transformação)                          │
│  └─ trnClientes, trnProdutos, trnVendas                 │
│  └─ Tratamentos: TRIM, UPPER, LOWER, etc.               │
│  └─ Limpeza de dados e formatação                       │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  QA (Quality Assurance/Validação)                       │
│  └─ qaClientes, qaProdutos, qaVendas                    │
│  └─ Validações: REQUIRED, EMAIL, CPFVAL, CNPJVAL, CEPVAL, etc.        │
│  └─ Flagging de registros com erros/avisos              │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  NRM (Normalize/Normalização)                           │
│  └─ nrmClientes, nrmProdutos, nrmVendas                 │
│  └─ Deduplicação, relacionamentos, enriquecimento       │
│  └─ Regras de negócio complexas                         │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  DIM/FATO (Dimensional Model)                           │
│  └─ dimClientes, dimProdutos, dimCalendario             │
│  └─ fatoVendas                                          │
│  └─ Modelo pronto para análise                          │
└─────────────────────────────────────────────────────────┘
```

---

## 🎯 Responsabilidades por Camada

### 1️⃣ STAGING (STG)

**Função:** Preparação estrutural e normalização básica

**Funções Principais:**
- `fxStgPreparar(tabela, ignorarColunas)` - Limpeza inicial
  - Remove linhas completamente vazias
  - Normaliza nomes de colunas (trim)
  - Valida estrutura (sem nomes vazios, sem duplicidade)
  - Mantém a tabela como source para otimizar avaliação lazy
  
- `fxStgAplicar(Tabela, Schema, ignorarColunas)` - Aplicação completa do estágio
  - Chama `fxStgPreparar()` para estrutura
  - Obtém `Pipeline = fxPipeline(Schema)`
  - Aplica tipos básicos via `Table.TransformColumns`
  - Reordena colunas com `Table.ReorderColumns`

**Exemplo de Uso:**
```powerquery
stgClientes = let
    Fonte = srcClientes,
    Preparada = fxStgAplicar(Fonte, "tbClientes"),
    Resultado = Table.Distinct(Preparada)
in
    Resultado;
```

**Dados de Saída:**
- Tabela estruturada e validada
- Tipos básicos aplicados conforme schema
- Colunas ordenadas segundo o pipeline
- Sem linhas completamente vazias

---

### 2️⃣ TRANSFORMAÇÃO (TRN)

**Função:** Aplicação de tratamentos e transformações de dados

**Funções Principais:**
- `fxTrnAplicar()` - Executa pipeline de transformações
  - Compila operadores de tratamento do schema
  - Aplica TRIM, UPPER, LOWER, PROPER, CLEAN, etc.
  - Converte tipos com tratamento de erros
  - Executa em uma única passagem (otimizado)

**Operadores Disponíveis:**
```
Texto:
  - TRIM, UPPER, LOWER, PROPER, CLEAN
  - SINGLESPACE, DIGITS, ALPHANUMERIC
  - NORMALIZEBASIC, NORMALIZETYPE

Numérico:
  - ABS, ROUND

Conversão:
  - Conversão automática entre tipos
  - CPF, CNPJ, CEP (extração de dígitos)
```

**Exemplo de Uso:**
```powerquery
trnClientes = let
    Fonte = stgClientes,
    Resultado = fxTrnAplicar(Fonte, "tbClientes")
in
    Resultado;
```

**Dados de Saída:**
- Dados tratados e transformados
- Colunas com tipos refinados
- Pronto para validação

---

### 3️⃣ QUALIDADE (QA)

**Função:** Validação estrutural, semântica e de negócio

**Funções Principais:**
- `fxQaValidar()` - Executa validações
  - Compila operadores de validação
  - Adiciona coluna `_QA_Ocorrencias` com detalhes de erro
  - Adiciona coluna `_QA_Status` (OK, AVISO, ERRO)
  - Não remove dados, apenas marca

- `fxQaFiltrarPorStatus()` - Filtra registros por status
  - OK, AVISO ou ERRO
  - Remove colunas de controle QA

- `fxQaExtrairProblemas()` - Extrai registros com problemas
  - Útil para auditoria e correção

**Validadores Disponíveis:**
```
Obrigatoriedade:
  - REQUIRED

Documentos (Brasil):
  - CPFVAL, CNPJVAL, CEPVAL

Internet:
  - EMAIL, URL, DOMAIN

Intervalo:
  - MIN, MAX, INTERVAL

Listas:
  - LIST, SIZE

Personalizadas:
  - Criar via schema
```

**Exemplo de Uso:**
```powerquery
qaClientes = let
    Fonte = trnClientes,
    Resultado = fxQaValidar(Fonte, "tbClientes")
in
    Resultado;

// Usar apenas dados válidos
ClientesValidos = fxQaFiltrarPorStatus(qaClientes, "OK");

// Extrair problemas para auditoria
ProblemasCliente = fxQaExtrairProblemas(qaClientes);
```

**Dados de Saída:**
- Tabela com colunas de controle QA
- `_QA_Status`: OK, AVISO ou ERRO
- `_QA_Ocorrencias`: Lista de problemas encontrados
- Todas as colunas originais mantidas

---

### 4️⃣ NORMALIZAÇÃO (NRM)

**Função:** Normalização de dados e deduplicação por chave de negócio

**Funções Principais:**
- `fxNrmAplicar(tabela, Schema)` - Deduplica por chaves de negócio definidas no schema
  - Filtra registros válidos (OK da QA)
  - Mantém a tabela intacta se não houver chaves de negócio definidas
  - Usa `Table.Distinct(tabela, Chaves)` para remover duplicatas de domínio

**Exemplo de Uso:**
```powerquery
nrmClientes = let
    Fonte = qaClientes,
    Valida = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Valida, "tbClientes")
in
    Normalizada;
```

**Dados de Saída:**
- Dados deduplicados segundo o schema
- Sem reconstrução extra além da deduplicação
- Pronto para o modelo dimensional

---

### 5️⃣ MODELO DIMENSIONAL (DIM/FATO)

**Função:** Preparação final para análise

**Construções:**
- **Dimensões** (dimClientes, dimProdutos, dimCalendario)
  - Chaves substitutas (surrogate keys)
  - Deduplicadas
  - Ordenadas para estabilidade
  - Bufferizadas para performance
  
- **Tabelas Fato** (fatoVendas)
  - Relacionamentos com dimensões
  - Métricas e medidas
  - Otimizadas para BI

**Exemplo de Uso:**
```powerquery
dimClientes = let
    Fonte = nrmClientes,
    Chaves = Table.AddIndexColumn(Fonte, "IDCliente", 1, 1, Int64.Type),
    Reordenada = Table.ReorderColumns(Chaves, 
        {"IDCliente", "CPF", "Nome", "DataNascimento", "Cidade", "Estado"})
in
    Table.Buffer(Reordenada);

fatoVendas = let
    // Merge com dimensões
    ComCliente = Table.NestedJoin(nrmVendas, {"CPF"}, dimClientes, {"CPF"}, ...)
    // ... relacionar com outras dimensões ...
in
    // Resultado otimizado para análise
    Resultado;
```

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


