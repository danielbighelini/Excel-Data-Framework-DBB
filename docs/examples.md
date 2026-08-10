# Examples

## 1. Visão Geral

Este documento apresenta exemplos práticos de utilização do Excel Data Framework DBB.

Os exemplos seguem o fluxo arquitetural do framework:

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

A configuração é definida no Schema e compilada pelo pipeline.

```text
Schema
   │
   ▼
Pipeline
   │
   ▼
Execução
```

Os exemplos abaixo utilizam como referência as estruturas e funções documentadas no projeto, como `fxStgAplicar`, `fxTrnAplicar`, `fxQaValidar`, `fxQaFiltrarPorStatus` e `fxNrmAplicar`.

---

# 2. Exemplo Básico

Considere uma tabela de clientes:

```text
CPF
Nome
Cidade
Estado
```

O Schema pode definir:

```text
Tabela      | Coluna  | Tipo | Obrigatório | Tratamentos      | Validações
tbClientes  | CPF     | TEXT | SIM         | TRIM;DIGITS      | CPFVAL
tbClientes  | Nome    | TEXT | SIM         | TRIM;PROPER      | SIZE(100)
tbClientes  | Cidade  | TEXT | NÃO         | TRIM;PROPER      |
tbClientes  | Estado  | TEXT | NÃO         | TRIM;UPPER       | LIST(RS,SP,SC)
```

O fluxo resultante é:

```text
srcClientes
     │
     ▼
stgClientes
     │
     ▼
trnClientes
     │
     ▼
qaClientes
     │
     ▼
nrmClientes
```

---

# 3. STG — Preparação e Tipagem

A camada STG recebe a fonte e aplica a preparação definida pelo Schema.

Exemplo:

```powerquery
stgClientes =
let
    Fonte = srcClientes,
    Preparada = fxStgAplicar(Fonte, "tbClientes")
in
    Preparada;
```

A função:

```text
fxStgAplicar(Tabela, Schema)
```

prepara a estrutura e aplica os tipos básicos definidos no Schema.

A camada também registra problemas de tipagem em `_STG_Ocorrencias`.

---

# 4. TRN — Tratamentos

Depois do STG, os tratamentos definidos no Schema são aplicados.

Exemplo:

```powerquery
trnClientes =
let
    Fonte = stgClientes,
    Transformada = fxTrnAplicar(Fonte, "tbClientes")
in
    Transformada;
```

Considerando:

```text
CPF:
TRIM;DIGITS

Nome:
TRIM;PROPER

Estado:
TRIM;UPPER
```

o pipeline aplica essas transformações por coluna.

A camada TRN utiliza `Table.TransformColumns` e permite encadear tratamentos.

---

# 5. QA — Validação

Depois dos tratamentos, os dados passam pela camada QA.

Exemplo:

```powerquery
qaClientes =
let
    Fonte = trnClientes,
    Validada = fxQaValidar(Fonte, "tbClientes")
in
    Validada;
```

O resultado contém informações de qualidade:

```text
_QA_Status
_QA_Ocorrencias
```

O status pode ser:

```text
OK
AVISO
ERRO
```

A camada QA não remove registros durante a validação. Ela marca os problemas para que a decisão de filtragem possa ocorrer posteriormente.

---

# 6. Extraindo os Problemas

Quando for necessário analisar somente os registros com problemas:

```powerquery
Problemas =
    fxQaExtrairProblemas(qaClientes)
```

O resultado pode ser utilizado para:

* auditoria;
* correção de dados;
* relatório de qualidade;
* acompanhamento das fontes.

A função `fxQaExtrairProblemas` é documentada como responsável por retornar os registros que possuem problemas.

---

# 7. Filtrando Registros Válidos

Caso somente registros `OK` devam continuar para a normalização:

```powerquery
ClientesValidos =
    fxQaFiltrarPorStatus(qaClientes, "OK")
```

O fluxo passa a ser:

```text
qaClientes
     │
     ▼
fxQaFiltrarPorStatus(..., "OK")
     │
     ▼
Clientes válidos
```

A filtragem ocorre depois da validação, mantendo a responsabilidade de diagnóstico separada da decisão de continuidade.

---

# 8. NRM — Normalização

Depois da validação e filtragem, a camada NRM pode remover duplicidades com base nas chaves de negócio.

Exemplo:

```powerquery
nrmClientes =
let
    Fonte = qaClientes,
    Valida = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Valida, "tbClientes")
in
    Normalizada;
```

Se `CPF` for a chave de negócio:

```text
ChavesNegocio
    ↓
CPF
    ↓
Table.Distinct
```

A função `fxNrmAplicar` utiliza as chaves de negócio definidas no Schema para deduplicação.

---

# 9. Pipeline Completo

Com as etapas anteriores, o fluxo completo pode ser representado como:

```text
srcClientes
     │
     ▼
fxStgAplicar
     │
     ▼
stgClientes
     │
     ▼
fxTrnAplicar
     │
     ▼
trnClientes
     │
     ▼
fxQaValidar
     │
     ▼
qaClientes
     │
     ├──────────────► Problemas / Auditoria
     │
     ▼
fxQaFiltrarPorStatus
     │
     ▼
Registros válidos
     │
     ▼
fxNrmAplicar
     │
     ▼
nrmClientes
```

Esse fluxo corresponde à sequência STG → TRN → QA → NRM documentada no projeto.

---

# 10. Exemplo de CPF

Considere os valores:

```text
123.456.789-09
12345678909
  123.456.789-09
```

Schema:

```text
Tratamentos:
TRIM;DIGITS

Validação:
CPFVAL
```

O tratamento:

```text
TRIM
  ↓
DIGITS
```

remove espaços e caracteres não numéricos.

Depois:

```text
CPF tratado
     ↓
CPFVAL
     ↓
OK / ERRO
```

A vantagem é separar:

```text
Tratamento
    ↓
Padroniza o valor

Validação
    ↓
Verifica o valor
```

---

# 11. Exemplo de Nome

Considere:

```text
"   joão da silva   "
```

Schema:

```text
Tratamentos:
TRIM;PROPER
```

Pipeline:

```text
"   joão da silva   "
          │
          ▼
        TRIM
          │
          ▼
"joão da silva"
          │
          ▼
        PROPER
          │
          ▼
"João Da Silva"
```

O exemplo demonstra a composição sequencial de operadores.

O comportamento exato de cada operador deve ser consultado em `operators.md`.

---

# 12. Exemplo de Estado

Considere:

```text
rs
```

Schema:

```text
Tratamentos:
TRIM;UPPER

Validação:
LIST(RS,SP,SC)
```

Fluxo:

```text
rs
 │
 ▼
TRIM
 │
 ▼
UPPER
 │
 ▼
RS
 │
 ▼
LIST(RS,SP,SC)
 │
 └── OK
```

Nesse caso, o tratamento padroniza o valor antes da validação.

---

# 13. Exemplo de Campo Obrigatório

Considere:

```text
Nome = ""
```

Schema:

```text
Obrigatório = SIM
```

A compilação pode incluir `REQUIRED` implicitamente:

```text
Obrigatório = SIM
       │
       ▼
REQUIRED
       │
       ▼
QA
       │
       ▼
ERRO
```

Não é necessário depender exclusivamente da declaração textual de `REQUIRED`.

O comportamento de `REQUIRED` implícito é parte da compilação das validações por coluna documentada no projeto.

---

# 14. Exemplo de Validação por Lista

Considere uma coluna `Estado`.

Schema:

```text
Validações:
LIST(RS,SP,SC)
```

Valores:

```text
RS
SP
SC
RJ
```

Resultado conceitual:

```text
RS → OK
SP → OK
SC → OK
RJ → ERRO
```

O objetivo da validação é verificar pertencimento ao domínio configurado.

---

# 15. Exemplo de Tamanho

Para um campo limitado a 100 caracteres:

```text
Validações:
SIZE(100)
```

O pipeline pode representar:

```text
Nome
 │
 ▼
SIZE(100)
 │
 ├── dentro do limite → OK
 └── acima do limite  → ocorrência
```

Esse tipo de validação está disponível através de `fxValidacaoSize`.

---

# 16. Exemplo de Chave de Negócio

Considere:

```text
Tabela: tbClientes

Chave:
CPF
```

Dados:

```text
CPF          Nome
12345678909  João
12345678909  João
98765432100  Maria
```

Após a NRM:

```text
CPF          Nome
12345678909  João
98765432100  Maria
```

A deduplicação utiliza a chave de negócio configurada no Schema.

---

# 17. Exemplo de Chave Composta

Quando um único campo não identifica o registro:

```text
CodigoCliente
CodigoProduto
```

podem formar uma chave composta.

Conceitualmente:

```text
ChavesNegocio =
{
    CodigoCliente,
    CodigoProduto
}
```

A identidade do registro passa a ser determinada pela combinação dos campos.

---

# 18. Exemplo de Pipeline com Dimensão

Depois da NRM, uma dimensão pode ser criada.

Exemplo documentado:

```powerquery
dimClientes =
let
    Fonte = nrmClientes,
    Chaves = Table.AddIndexColumn(
        Fonte,
        "IDCliente",
        1,
        1,
        Int64.Type
    ),
    Reordenada = Table.ReorderColumns(
        Chaves,
        {"IDCliente", "CPF", "Nome", "DataNascimento", "Cidade", "Estado"}
    )
in
    Table.Buffer(Reordenada);
```

O exemplo demonstra:

1. utilização da tabela normalizada;
2. criação de chave substituta;
3. ordenação das colunas;
4. materialização da dimensão.

A utilização de `Table.AddIndexColumn`, `Table.ReorderColumns` e bufferização em dimensões está documentada no projeto.

---

# 19. Exemplo de Fato

Uma tabela fato pode utilizar dimensões previamente construídas.

Conceitualmente:

```text
nrmVendas
    │
    ├── Join dimClientes
    │
    ├── Join dimProdutos
    │
    └── Join dimCalendario
             │
             ▼
          fatoVendas
```

A integração entre estruturas pode utilizar `Table.NestedJoin`.

Esse padrão é utilizado na etapa de modelagem dimensional documentada no projeto.

---

# 20. Exemplo de Tratamento Parametrizado

Operadores podem receber parâmetros.

Exemplo:

```text
ADDPREFIX(P)
```

O pipeline identifica:

```text
Código:
ADDPREFIX

Parâmetro:
P
```

Outro exemplo:

```text
PADLEFT(3,0,P)
```

O compilador extrai os parâmetros:

```text
Operador:
PADLEFT

Parâmetros:
3
0
P
```

A compilação dos parâmetros é realizada antes da execução do pipeline.

---

# 21. Exemplo de Múltiplos Tratamentos

Uma coluna pode possuir vários tratamentos:

```text
TRIM;UPPER;REMOVEACCENTS
```

Fluxo:

```text
Valor
  │
  ▼
TRIM
  │
  ▼
UPPER
  │
  ▼
REMOVEACCENTS
  │
  ▼
Resultado
```

A ordem dos operadores faz parte da configuração.

Por isso, a seguinte sequência pode produzir um fluxo diferente:

```text
REMOVEACCENTS;TRIM;UPPER
```

---

# 22. Exemplo de Múltiplas Validações

Uma coluna pode possuir mais de uma validação:

```text
REQUIRED;SIZE(100)
```

Fluxo conceitual:

```text
Valor
  │
  ├── REQUIRED
  │
  └── SIZE(100)
       │
       ▼
   Ocorrências
```

O QA consolida as ocorrências e determina o status final.

---

# 23. Exemplo de Registro com Problemas

Considere:

```text
CPF = inválido
Nome = vazio
Estado = RJ
```

Schema:

```text
CPF
  Obrigatório = SIM
  Validação = CPFVAL

Nome
  Obrigatório = SIM

Estado
  Validação = LIST(RS,SP,SC)
```

O resultado pode conter:

```text
_QA_Status = ERRO

_QA_Ocorrencias =
[
    problema de CPF,
    problema de Nome,
    Estado fora da lista
]
```

O registro original permanece disponível na camada QA.

Isso permite auditoria antes da filtragem.

---

# 24. Exemplo de Auditoria

Para obter apenas registros problemáticos:

```powerquery
Problemas =
    fxQaExtrairProblemas(qaClientes)
```

Fluxo:

```text
qaClientes
     │
     ├── OK
     │
     ├── AVISO
     │
     └── ERRO
          │
          ▼
fxQaExtrairProblemas
          │
          ▼
Relatório de problemas
```

Essa estrutura pode ser usada para identificar problemas de qualidade sem alterar a tabela operacional principal.

---

# 25. Exemplo de Fluxo Completo de Produção

Uma implementação típica pode separar as consultas por etapa:

```text
srcClientes
     │
     ▼
stgClientes
     │
     ▼
trnClientes
     │
     ▼
qaClientes
     │
     ├──────────────► problemasClientes
     │
     ▼
nrmClientes
     │
     ▼
dimClientes
```

Para vendas:

```text
srcVendas
     │
     ▼
stgVendas
     │
     ▼
trnVendas
     │
     ▼
qaVendas
     │
     ▼
nrmVendas
     │
     ▼
fatoVendas
```

Esse padrão mantém as responsabilidades das etapas separadas.

---

# 26. Exemplo de Estrutura de Projeto

Uma organização possível das consultas é:

```text
SRC
├── srcClientes
├── srcProdutos
└── srcVendas

STG
├── stgClientes
├── stgProdutos
└── stgVendas

TRN
├── trnClientes
├── trnProdutos
└── trnVendas

QA
├── qaClientes
├── qaProdutos
└── qaVendas

NRM
├── nrmClientes
├── nrmProdutos
└── nrmVendas

DIM
├── dimClientes
├── dimProdutos
└── dimCalendario

FATO
└── fatoVendas
```

Essa organização torna o fluxo visualmente identificável no ambiente do Power Query.

---

# 27. Exemplo de Configuração para Clientes

Um Schema completo pode ser organizado conceitualmente assim:

```text
Tabela: tbClientes

CPF
  Tipo: TEXT
  Obrigatório: SIM
  Tratamentos: TRIM;DIGITS
  Validações: CPFVAL
  Chave: SIM

Nome
  Tipo: TEXT
  Obrigatório: SIM
  Tratamentos: TRIM;PROPER
  Validações: SIZE(100)

DataNascimento
  Tipo: DATE
  Obrigatório: NÃO

Cidade
  Tipo: TEXT
  Obrigatório: NÃO
  Tratamentos: TRIM;PROPER

Estado
  Tipo: TEXT
  Obrigatório: NÃO
  Tratamentos: TRIM;UPPER
  Validações: LIST(RS,SP,SC)
```

O resultado esperado é:

```text
Schema
  │
  ▼
cfgPipeline[tbClientes]
  │
  ├── STG
  ├── TRN
  ├── QA
  └── NRM
```

---

# 28. Exemplo de Configuração para Produtos

Uma tabela de produtos pode seguir o mesmo mecanismo:

```text
Tabela: tbProdutos

Codigo
  Tipo: TEXT
  Obrigatório: SIM
  Tratamentos: TRIM;UPPER

Descricao
  Tipo: TEXT
  Obrigatório: SIM
  Tratamentos: TRIM;PROPER
  Validações: SIZE(200)

Preco
  Tipo: DECIMAL
  Obrigatório: SIM
  Tratamentos: NUMBER
  Validações: MIN(0)
```

A lógica de processamento continua sendo a mesma.

Somente a configuração muda.

---

# 29. Exemplo de Reutilização

O principal benefício do modelo declarativo pode ser observado quando duas tabelas diferentes utilizam o mesmo operador.

```text
tbClientes
    │
    └── Nome → TRIM;PROPER

tbFornecedores
    │
    └── Nome → TRIM;PROPER
```

O operador não precisa ser implementado duas vezes.

A mesma função é reutilizada pelo pipeline.

```text
              TRIM
               │
        ┌──────┴──────┐
        ▼             ▼
    Clientes     Fornecedores
```

---

# 30. Exemplo de Pipeline Compilado

Conceitualmente, a configuração:

```text
CPF
  Tratamentos = TRIM;DIGITS
  Validações = CPFVAL
```

pode resultar em:

```text
TratamentosPorColuna =
[
    CPF = {
        TRIM,
        DIGITS
    }
]

ValidaçõesPorColuna =
[
    CPF = {
        REQUIRED,
        CPFVAL
    }
]
```

O pipeline compilado mantém as funções e parâmetros necessários para execução.

---

# 31. Exemplo de Performance

Considere uma tabela com:

```text
100.000 registros
20 colunas
```

mas apenas três colunas possuem tratamentos:

```text
CPF
Nome
Estado
```

O Schema permite que o pipeline concentre o processamento nessas colunas:

```text
CPF   → tratamentos
Nome  → tratamentos
Estado → tratamentos

17 colunas
     ↓
sem tratamento específico
```

Essa abordagem evita aplicar transformações genéricas sobre todas as colunas.

O framework documenta explicitamente o processamento somente das colunas presentes no Schema.

---

# 32. Exemplo de QA em Uma Passagem

Considere três validações:

```text
CPFVAL
REQUIRED
SIZE(100)
```

A estratégia do framework é acumular as ocorrências durante a avaliação:

```text
Registro
   │
   ▼
Validação CPF
   │
   ▼
Validação REQUIRED
   │
   ▼
Validação SIZE
   │
   ▼
_QA
```

Em vez de criar uma tabela intermediária para cada validação.

O resultado final é consolidado em:

```text
_QA_Status
_QA_Ocorrencias
```

Essa estratégia é documentada como uma das principais otimizações do framework.

---

# 33. Exemplo de Fluxo para Dados Inválidos

O framework mantém uma separação entre tratamento, validação e filtragem.

```text
Fonte
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
 ├── Registro válido
 │
 └── Registro com problema
          │
          ▼
     _QA_Ocorrencias
          │
          ▼
     _QA_Status
          │
          ▼
    Filtragem opcional
```

Isso permite que os dados problemáticos sejam analisados antes de serem descartados da etapa posterior.

---

# 34. Exemplo de Integração com Modelo Dimensional

Um fluxo completo de vendas pode ser:

```text
srcClientes ──► stgClientes ──► trnClientes ──► qaClientes ──► nrmClientes ──► dimClientes
                                                                                       │
                                                                                       │
srcProdutos ──► stgProdutos ──► trnProdutos ──► qaProdutos ──► nrmProdutos ──► dimProdutos
                                                                                       │
                                                                                       │
srcVendas ────► stgVendas ────► trnVendas ────► qaVendas ────► nrmVendas ─────► fatoVendas
                                                                                       │
                                                                                       ▼
                                                                                Modelo analítico
```

As dimensões são preparadas antes da construção do fato, permitindo que as relações sejam resolvidas durante a modelagem.

---

# 35. Exemplo de Princípio Declarativo

O objetivo do framework é permitir que o usuário descreva:

```text
CPF
→ limpar
→ validar
→ usar como chave
```

em vez de implementar manualmente:

```text
if ...
then ...
else ...
```

para cada tabela.

A configuração representa a intenção:

```text
Tratamentos = TRIM;DIGITS
Validações = CPFVAL
Chave = CPF
```

O framework determina como executar essa configuração.

---

# 36. Quando Criar um Novo Operador

Se uma regra aparece repetidamente em diferentes Schemas, pode ser apropriado transformá-la em um operador reutilizável.

Exemplo:

```text
Regra repetida:

TRIM
UPPER
REMOVEACCENTS
```

Se a combinação for recorrente e possuir semântica própria, pode-se avaliar a criação de um operador específico.

A decisão deve considerar:

* reutilização;
* clareza;
* estabilidade da regra;
* impacto na manutenção;
* necessidade de parâmetros.

Os detalhes de extensão do catálogo pertencem a `operators.md`.

---

# 37. Exemplo de Fluxo Recomendado

Para uma nova tabela:

```text
1. Criar fonte
       ↓
2. Definir Schema
       ↓
3. Configurar tipos
       ↓
4. Configurar tratamentos
       ↓
5. Configurar validações
       ↓
6. Definir chaves de negócio
       ↓
7. Criar STG
       ↓
8. Criar TRN
       ↓
9. Criar QA
       ↓
10. Criar NRM
       ↓
11. Criar DIM/FATO
```

Esse processo evita colocar regras específicas diretamente nas consultas de processamento.

---

# 38. Referência Rápida

| Necessidade           | Configuração / Função  |
| --------------------- | ---------------------- |
| Preparar tabela       | `fxStgAplicar`         |
| Aplicar tipos         | `TiposPorColuna`       |
| Aplicar tratamentos   | `fxTrnAplicar`         |
| Validar dados         | `fxQaValidar`          |
| Extrair problemas     | `fxQaExtrairProblemas` |
| Filtrar status        | `fxQaFiltrarPorStatus` |
| Deduplicar            | `fxNrmAplicar`         |
| Definir comportamento | Schema                 |
| Compilar configuração | `fxPipeline`           |
| Controlar tipos       | `cfgTiposDados`        |
| Controlar operadores  | `cfgOperadores`        |

A relação entre essas estruturas está documentada no README do projeto.

---

# 39. Documentação Relacionada

* [Getting Started](getting-started.md) — primeiro contato com o framework.
* [Architecture](architecture.md) — arquitetura geral.
* [Schema](schema.md) — configuração declarativa.
* [Pipeline](pipeline.md) — compilação e execução.
* [Operators](operators.md) — tratamentos e validações.
* [Performance](performance.md) — estratégias de otimização.
* [Troubleshooting](troubleshooting.md) — diagnóstico.

---

# 40. Resumo

Os exemplos demonstram o princípio central do framework:

```text
              CONFIGURAÇÃO
                   │
                   ▼
                 Schema
                   │
                   ▼
             Pipeline
                   │
                   ▼
        ┌──────────┼──────────┐
        ▼          ▼          ▼
       STG        TRN         QA
        │          │          │
     Tipagem   Tratamentos  Validações
        │          │          │
        └──────────┼──────────┘
                   ▼
                  NRM
                   │
                   ▼
              DIM / FATO
```

A configuração define o comportamento, o pipeline compila as regras e as camadas executam cada responsabilidade sobre os dados.

Os exemplos devem ser utilizados em conjunto com:

* `schema.md` para entender a configuração;
* `operators.md` para consultar os operadores;
* `pipeline.md` para entender a compilação;
* `architecture.md` para compreender as responsabilidades das camadas;
* `performance.md` para compreender as decisões de otimização.
