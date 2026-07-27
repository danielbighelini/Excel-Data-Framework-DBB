# 🎯 Exemplos Práticos - Excel Data Framework DBB

## Exemplo 1: Pipeline Completo para Clientes

### Objetivo
Processar dados de clientes (CPF, Nome, Cidade, Estado) com:
- Limpeza e normalização
- Tratamentos (TRIM, UPPER)
- Validações (REQUIRED, CPF válido)
- Deduplicação

### Implementação

```powerquery
// 1. Definir Schema no Excel
// tbSchema:
// Tabela    | Coluna  | Tipo | Obrigatório | Tratamentos      | Validações
// tbClientes| CPF     | TEXT | SIM         | TRIM,UPPER,DIGITS| REQUIRED,CPFVAL
// tbClientes| Nome    | TEXT | SIM         | TRIM,PROPER      | REQUIRED
// tbClientes| Cidade  | TEXT | NÃO         | TRIM,PROPER      |
// tbClientes| Estado  | TEXT | NÃO         | TRIM,UPPER       |

// 2. Fonte de dados
shared srcClientes = srcWorkbook{[Name="tbClientes"]}[Content];

// 3. Pipeline Completo
shared stgClientes = fxStgAplicar(srcClientes, "tbClientes");
shared trnClientes = fxTrnAplicar(stgClientes, "tbClientes");
shared qaClientes = fxQaValidar(trnClientes, "tbClientes");
shared nrmClientes = let
    Fonte = qaClientes,
    Valida = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Valida, "tbClientes"),
    RegistrosUnicos = Table.Distinct(Normalizada, {"CPF"})
in
    RegistrosUnicos;

// 4. Dimensão para análise
shared dimClientes = let
    Fonte = nrmClientes,
    Chaves = Table.AddIndexColumn(
        Table.Sort(Table.Distinct(Fonte), {{"CPF", Order.Ascending}}),
        "IDCliente", 1, 1, Int64.Type
    ),
    Reordenada = Table.ReorderColumns(Chaves, 
        {"IDCliente", "CPF", "Nome", "Cidade", "Estado"})
in
    Table.Buffer(Reordenada);

// 5. Auditoria (Opcional)
shared auditClientes = fxQaExtrairProblemas(qaClientes);
```

### Resultado Esperado

```
STG: 5 registros, alguns com nomes/CPFs com espaços/lowercase
TRN: 5 registros, TRIM/UPPER aplicados
QA:  5 registros, 3 OK, 2 com ERRO (CPF inválido)
NRM: 3 registros (apenas OK), deduplicados por CPF
DIM: 3 registros com ID + colunas ordenadas
AUD: 2 registros com detalhes dos erros
```

---

## Exemplo 2: Auditoria e Tratamento de Erros

### Objetivo
Identificar e corrigir dados problemáticos

### Implementação

```powerquery
// 1. Extrair problemas
shared relatorioProblemas = let
    Fonte = qaClientes,
    ComProblemas = fxQaExtrairProblemas(Fonte),
    Expandida = Table.ExpandRecordColumn(
        ComProblemas,
        "_QA_Ocorrencias",
        {"Codigo", "Coluna", "Mensagem", "Detalhes"}
    ),
    ComNomeCliente = Table.AddColumn(
        Expandida,
        "NomeCliente",
        each [Nome],  // Nome original
        type text
    ),
    Reordenada = Table.ReorderColumns(
        ComNomeCliente,
        {"CPF", "NomeCliente", "Coluna", "Codigo", "Mensagem", "Detalhes"}
    )
in
    Reordenada;

// 2. Análise de problemas por tipo
shared analiseProblemaspPorCodigo = let
    Fonte = relatorioProblemas,
    Agrupada = Table.Group(
        Fonte,
        {"Codigo"},
        {
            {"Quantidade", Table.RowCount, Int64.Type},
            {"Colunas", each Text.Combine(List.Distinct([Coluna]), ", "), type text},
            {"Detalhes", each Text.Combine(List.Distinct(Text.From([Detalhes])), "; "), type text}
        }
    ),
    Ordenada = Table.Sort(Agrupada, {{"Quantidade", Order.Descending}})
in
    Ordenada;

// 3. Exemplo de correção: Registros com avisos (não erros)
shared clientesComAviso = let
    Fonte = qaClientes,
    ComAviso = Table.SelectRows(
        Fonte,
        each [_QA_Status] = "AVISO"
    ),
    Expandida = Table.ExpandRecordColumn(
        ComAviso,
        "_QA_Ocorrencias",
        {"Codigo", "Mensagem"}
    )
in
    Expandida;

// 4. Estatísticas de qualidade
shared estatisticasQA = let
    Fonte = qaClientes,
    Total = Table.RowCount(Fonte),
    OK = Table.RowCount(Table.SelectRows(Fonte, each [_QA_Status] = "OK")),
    AVISO = Table.RowCount(Table.SelectRows(Fonte, each [_QA_Status] = "AVISO")),
    ERRO = Table.RowCount(Table.SelectRows(Fonte, each [_QA_Status] = "ERRO")),
    Resultado = #table(
        {"Métrica", "Quantidade", "Percentual"},
        {
            {"Total", Total, 1.0},
            {"OK", OK, OK/Total},
            {"AVISO", AVISO, AVISO/Total},
            {"ERRO", ERRO, ERRO/Total}
        }
    )
in
    Resultado;
```

### Uso Prático

```
relatorioProblemas → Excel Report
analiseProblemaspPorCodigo → Dashboard
clientesComAviso → Lista para revisão manual
estatisticasQA → KPI de qualidade
```

---

## Exemplo 3: Pipeline com Múltiplas Entidades

### Objetivo
Processar 3 entidades relacionadas (Clientes, Produtos, Vendas) de forma coordenada

### Implementação

```powerquery
// ======================== CLIENTS ========================
shared stgClientes = fxStgAplicar(srcClientes, "tbClientes");
shared trnClientes = fxTrnAplicar(stgClientes, "tbClientes");
shared qaClientes = fxQaValidar(trnClientes, "tbClientes");
shared nrmClientes = let
    Fonte = fxQaFiltrarPorStatus(qaClientes, "OK"),
    Normalizada = fxNrmAplicar(Fonte, "tbClientes")
in
    Table.Distinct(Normalizada, {"CPF"});

// ======================== PRODUCTS ========================
shared stgProdutos = fxStgAplicar(srcProdutos, "tbProdutos");
shared trnProdutos = fxTrnAplicar(stgProdutos, "tbProdutos");
shared qaProdutos = fxQaValidar(trnProdutos, "tbProdutos");
shared nrmProdutos = let
    Fonte = fxQaFiltrarPorStatus(qaProdutos, "OK"),
    Normalizada = fxNrmAplicar(Fonte, "tbProdutos")
in
    Table.Distinct(Normalizada, {"Código"});

// ======================== SALES ========================
shared stgVendas = fxStgAplicar(srcVendas, "tbVendas");
shared trnVendas = fxTrnAplicar(stgVendas, "tbVendas");
shared qaVendas = fxQaValidar(trnVendas, "tbVendas");
shared nrmVendas = let
    Fonte = fxQaFiltrarPorStatus(qaVendas, "OK"),
    Normalizada = fxNrmAplicar(Fonte, "tbVendas"),
    
    // Validar relacionamentos
    ClientesValidos = List.Buffer(nrmClientes[CPF]),
    ProdutosValidos = List.Buffer(nrmProdutos[Código]),
    
    // Filtrar apenas vendas com cliente e produto válidos
    ValidadasRelacionamento = Table.SelectRows(
        Normalizada,
        each List.Contains(ClientesValidos, [CPF])
        and List.Contains(ProdutosValidos, [CodigoProduto])
    ),
    
    // Adicionar valor total
    ComValorTotal = Table.AddColumn(
        ValidadasRelacionamento,
        "ValorTotal",
        each [Quantidade] * [ValorUnitario],
        type number
    )
in
    ComValorTotal;

// ======================== DIMENSIONS ========================
shared dimClientes = let
    Chaves = Table.AddIndexColumn(
        Table.Sort(nrmClientes, {{"CPF", Order.Ascending}}),
        "IDCliente", 1, 1, Int64.Type
    )
in
    Table.Buffer(Chaves);

shared dimProdutos = let
    Chaves = Table.AddIndexColumn(
        Table.Sort(nrmProdutos, {{"Código", Order.Ascending}}),
        "IDProduto", 1, 1, Int64.Type
    )
in
    Table.Buffer(Chaves);

// ======================== FACT TABLE ========================
shared fatoVendas = let
    // Merge com Cliente
    ComCliente = Table.NestedJoin(
        nrmVendas,
        {"CPF"},
        dimClientes,
        {"CPF"},
        "_Cliente",
        JoinKind.Inner
    ),
    ClienteExpandido = Table.ExpandTableColumn(
        ComCliente,
        "_Cliente",
        {"IDCliente"}
    ),
    
    // Merge com Produto
    ComProduto = Table.NestedJoin(
        ClienteExpandido,
        {"CodigoProduto"},
        dimProdutos,
        {"Código"},
        "_Produto",
        JoinKind.Inner
    ),
    ProdutoExpandido = Table.ExpandTableColumn(
        ComProduto,
        "_Produto",
        {"IDProduto"}
    ),
    
    // Selecionar colunas finais
    Final = Table.SelectColumns(
        ProdutoExpandido,
        {"Data", "IDCliente", "IDProduto", "Quantidade", "ValorUnitario", "ValorTotal"}
    )
in
    Final;

// ======================== REPORTS ========================
shared relKPI_VendasPorCliente = let
    Fonte = fatoVendas,
    Agrupada = Table.Group(
        Fonte,
        {"IDCliente"},
        {
            {"TotalVendas", each List.Sum([ValorTotal]), type number},
            {"QuantidadeVendas", Table.RowCount, Int64.Type},
            {"TicketMedio", each List.Average([ValorTotal]), type number}
        }
    ),
    ComNome = Table.NestedJoin(
        Agrupada,
        {"IDCliente"},
        dimClientes,
        {"IDCliente"},
        "_Cliente",
        JoinKind.Inner
    ),
    Expandida = Table.ExpandTableColumn(
        ComNome,
        "_Cliente",
        {"Nome"}
    ),
    Reordenada = Table.ReorderColumns(
        Expandida,
        {"IDCliente", "Nome", "TotalVendas", "QuantidadeVendas", "TicketMedio"}
    ),
    Ordenada = Table.Sort(
        Reordenada,
        {{"TotalVendas", Order.Descending}}
    )
in
    Ordenada;

shared relQA_ResumoQualidade = let
    ClientesOK = Table.RowCount(Table.SelectRows(qaClientes, each [_QA_Status] = "OK")),
    ProdutosOK = Table.RowCount(Table.SelectRows(qaProdutos, each [_QA_Status] = "OK")),
    VendasOK = Table.RowCount(Table.SelectRows(qaVendas, each [_QA_Status] = "OK")),
    
    ClientesTotal = Table.RowCount(qaClientes),
    ProdutosTotal = Table.RowCount(qaProdutos),
    VendasTotal = Table.RowCount(qaVendas),
    
    Resultado = #table(
        {"Entidade", "OK", "Total", "Percentual"},
        {
            {"Clientes", ClientesOK, ClientesTotal, ClientesOK/ClientesTotal},
            {"Produtos", ProdutosOK, ProdutosTotal, ProdutosOK/ProdutosTotal},
            {"Vendas", VendasOK, VendasTotal, VendasOK/VendasTotal}
        }
    )
in
    Resultado;
```

### Visualização do Pipeline

```
srcClientes ──STG──TRN──QA──NRM──┐
                                  ├─ dimClientes ──┐
srcProdutos ──STG──TRN──QA──NRM──┤                ├─ fatoVendas ─┬─ relKPI_VendasPorCliente
                                  ├─ dimProdutos ──┤              │
srcVendas ───STG──TRN──QA──NRM────┤                └─ relQA_ResumoQualidade
                                  │
                                  └─ relProblemas (auditoria)
```

---

## Exemplo 4: Operadores Customizados

### Objetivo
Adicionar um tratamento customizado (ex: formatar como moeda)

### Passo 1: Definir Função

```powerquery
shared fxTratamentoFormatarMoeda = (valor as any, optional parametros as nullable any) as any =>

let
    Moeda = if parametros = null or List.IsEmpty(parametros) then "R$" else parametros{0},
    
    Resultado =
        if valor = null then null
        else
            let
                Numero = Number.From(valor),
                Formatado = Text.Format("#{0} #,##0.00", {Numero}, "pt-BR")
            in
                Moeda & " " & Formatado

in
    Resultado;
```

### Passo 2: Registrar em Operadores

Adicionar à tabela `srcOperadores` ou `tbParametrosTratamentos`:

```
Código          | Descrição                         | Ativo | Categoria
FORMATARMOEDA   | Formata valor como moeda          | SIM   | Tratamento
```

### Passo 3: Usar no Schema

```
Tabela    | Coluna      | Tipo   | Tratamentos
tbProdutos| PrecoLista  | NUMBER | ROUND,FORMATARMOEDA(R$)
```

### Resultado

```
100.5 → "R$ 100,50"
1234.9 → "R$ 1.234,90"
```

---

## Exemplo 5: Teste Unitário de Uma Função

### Objetivo
Testar fxQaValidar isoladamente

### Implementação

```powerquery
// Dados de teste
shared tst_DadosTeste = #table(
    {"CPF", "Nome", "Cidade"},
    {
        {"12345678901", "João Silva", "São Paulo"},
        {"", "Maria Santos", "Rio de Janeiro"},
        {"98765432101", "", "Belo Horizonte"}
    }
);

// Teste STG
shared tst_STG_Resultado = fxStgAplicar(tst_DadosTeste, "tbClientes");

shared tst_STG_Verificar = let
    Resultado = tst_STG_Resultado,
    Esperado = 3
in
    Table.RowCount(Resultado) = Esperado;

// Teste TRN
shared tst_TRN_Resultado = fxTrnAplicar(tst_STG_Resultado, "tbClientes");

shared tst_TRN_Verificar = let
    Resultado = tst_TRN_Resultado,
    PrimeiroNome = Resultado{0}[Nome],
    Esperado = "João Silva"  // PROPER aplicado
in
    PrimeiroNome = Esperado;

// Teste QA
shared tst_QA_Resultado = fxQaValidar(tst_TRN_Resultado, "tbClientes");

shared tst_QA_Verificar = let
    Resultado = tst_QA_Resultado,
    ComErro = Table.SelectRows(Resultado, each [_QA_Status] = "ERRO"),
    QuantidadeErros = Table.RowCount(ComErro),
    Esperado = 2  // 2 registros devem ter ERRO (REQUIRED)
in
    QuantidadeErros = Esperado;

// Dashboard de Testes
shared dashboard_TesesUnitarios = let
    Testes = #table(
        {"Teste", "Resultado", "Status"},
        {
            {"STG - Contagem Linhas", tst_STG_Verificar, if tst_STG_Verificar then "✅ PASS" else "❌ FAIL"},
            {"TRN - PROPER em Nome", tst_TRN_Verificar, if tst_TRN_Verificar then "✅ PASS" else "❌ FAIL"},
            {"QA - Validação REQUIRED", tst_QA_Verificar, if tst_QA_Verificar then "✅ PASS" else "❌ FAIL"}
        }
    )
in
    Testes;
```

---

## Exemplo 6: Monitoramento e Alertas

### Objetivo
Monitora qualidade diária e gera alertas

### Implementação

```powerquery
// Dia anterior
shared historico_QA_Dia01 = 95;  // 95% OK
shared historico_QA_Dia02 = 92;  // 92% OK (queda)
shared historico_QA_Dia03 = 88;  // 88% OK (queda maior)

// Dia atual
shared monitoramento_QA_Hoje = let
    Fonte = estatisticasQA,
    OK = List.Sum(
        Table.SelectRows(Fonte, each [Métrica] = "OK")[Quantidade]
    ),
    Total = List.Sum(
        Table.SelectRows(Fonte, each [Métrica] = "Total")[Quantidade]
    ),
    Percentual = OK/Total
in
    Percentual;

// Alertas
shared alertas_QA = let
    QualidadeHoje = monitoramento_QA_Hoje,
    
    Alertas = #table(
        {"Alerta", "Valor", "Limiar", "Status"},
        {
            {"Qualidade Mínima", QualidadeHoje, 0.85, if QualidadeHoje < 0.85 then "🔴 ALERTA" else "🟢 OK"},
            {"Queda vs Dia Anterior", QualidadeHoje, historico_QA_Dia03 - 0.05, if QualidadeHoje < (historico_QA_Dia03 - 0.05) then "🟡 ATENÇÃO" else "🟢 OK"}
        }
    )
in
    Alertas;
```

---

## Resumo de Exemplos

| Exemplo | Nível | Aplicação |
|---------|-------|-----------|
| 1 | Básico | Pipeline completo simples |
| 2 | Intermediário | Auditoria e correção |
| 3 | Avançado | Múltiplas entidades integradas |
| 4 | Avançado | Extensão com operadores |
| 5 | Avançado | Testes unitários |
| 6 | Avançado | Monitoramento e alertas |

---

## 🎓 Aprender Mais

Próximas funções a explorar:
- `fxCalendario()` - Criar dimensão tempo
- `fxConectorSQL()` - Conectar a bancos de dados
- `fxRESTRequest()` - Integrar com APIs
- Operadores customizados - Expandir framework

