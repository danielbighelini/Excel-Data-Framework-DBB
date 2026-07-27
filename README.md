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
│  └─ Validações: REQUIRED, EMAIL, CPF, CNPJ, etc.        │
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
- `fxStgPreparar()` - Limpeza inicial
  - Remove linhas completamente vazias
  - Normaliza nomes de colunas (trim, remover duplicatas)
  - Valida estrutura (sem nomes vazios, sem duplicidade)
  
- `fxStgAplicar()` - Aplicação completa do estágio
  - Chama `fxStgPreparar()`
  - Aplica tipos básicos conforme schema
  - Reordena colunas conforme pipeline

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
- Tabela estruturada e normalizada
- Colunas no tipo correto (quando aplicável)
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

**Função:** Normalização de dados, deduplicação, enriquecimento

**Funções Principais:**
- `fxNrmAplicar()` - Aplicação completa da normalização
  - Filtra registros válidos (OK da QA)
  - Remove duplicatas por domínio
  - Resolve relacionamentos
  - Aplica enriquecimento (customizável por schema)
  - Aplica regras complexas de negócio

**Exemplo de Uso:**
```powerquery
nrmClientes = let
    Fonte = qaClientes,
    Valida = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Valida, "tbClientes"),
    RegistrosUnicos = Table.Distinct(Normalizada, {"CPF"})
in
    RegistrosUnicos;
```

**Dados de Saída:**
- Dados estruturados e verificados
- Sem duplicatas (por chave de negócio)
- Relacionamentos resolvidos
- Enriquecido com dados externos
- Pronto para modelo dimensional

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

```powerquery
shared fxNrmCompilarTratamentosPorColuna = (Operadores as list) as function =>
    let
        OperadoresBuffer = List.Buffer(Operadores ?? {})
    in
        if List.IsEmpty(OperadoresBuffer) then
            (valor) => valor
        else
            (valor) =>
                List.Accumulate(
                    OperadoresBuffer,
                    valor,
                    (Estado, Operador) =>
                        Operador[Função](
                            Estado,
                            Record.FieldOrDefault(Operador, "Parâmetros", null)
                        )
                )
```

**Resultado:** Pipeline compilado, executado uma única vez

### Compilação de Validações

```powerquery
shared fxQaCompilarValidacoesPorColuna = (Operadores as list, Tipo as type, Coluna as text) =>
    // Retorna função que valida um valor
    // Acumula todas as ocorrências de erro
    // Retorna lista de problemas encontrados
```

---

## 📊 Schema e Pipeline

### Estrutura do Schema

```powerquery
// tbSchema (Excel)
Tabela      | Coluna           | Tipo    | Obrigatório | Tratamentos        | Validações
tbClientes  | CPF              | TEXT    | SIM         | TRIM;UPPER;DIGITS  | REQUIRED;CPFVAL
tbClientes  | Nome             | TEXT    | SIM         | TRIM;PROPER        | REQUIRED;SIZE(100)
tbClientes  | DataNascimento   | DATE    | NÃO         |                    |
tbClientes  | Cidade           | TEXT    | NÃO         | TRIM;PROPER        |
tbClientes  | Estado           | TEXT    | NÃO         | TRIM;UPPER         |
```

### Pipeline Compilado

```powerquery
cfgPipeline = Record.TransformFields(
    cfgSchema,
    List.Transform(
        Record.FieldNames(cfgSchema),
        each { _, fxPipelineCompilar }
    )
)
```

**Resultado:**
```
cfgPipeline[tbClientes][TratamentosPorColuna]
{
    "CPF": [Operador1, Operador2, Operador3],
    "Nome": [Operador1, Operador2],
    ...
}

cfgPipeline[tbClientes][ValidacoesPorColuna]
{
    "CPF": [Validador1, Validador2],
    "Nome": [Validador1],
    ...
}
```

---

## 🚀 Otimizações Implementadas

### 1. Processamento de tabelas otimizado

**Depois (Otimizado):**
```
STG (prepare) → STG (tipos) → TRN (transformações) → 
QA (validações) → NRM (normalização)
// Cada camada é uma passagem, não há reconstrução desnecessária
```

### 2. Compilação de Operadores

- Pipelines compilados uma única vez (`cfgPipeline`)
- Execução sequencial dentro de uma célula (não múltiplos passes)
- Acúmulo eficiente de transformações

### 3. Bufferização Estratégica

```powerquery
// Buffer apenas onde necessário
Table.Buffer(Table.Distinct(...))  // Dedup + Cache
List.Buffer(...)                   // Listas de operadores
```

### 4. Lazy Evaluation

- Operações aplicadas apenas a colunas que precisam
- Sem recálculos desnecessários
- Validação condicional (tipo any)

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


