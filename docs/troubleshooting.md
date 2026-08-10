# Troubleshooting

## 1. Visão Geral

Este documento apresenta procedimentos para diagnosticar problemas comuns no Excel Data Framework DBB.

A estratégia geral de diagnóstico é seguir o fluxo do pipeline:

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

Quando ocorrer um problema, primeiro identifique **em qual camada ele aparece**.

```text
Problema
   │
   ▼
Qual camada?
   │
   ├── STG → estrutura / tipos
   ├── TRN → tratamentos
   ├── QA  → validações / ocorrências
   ├── NRM → normalização / duplicidades
   └── DIM/FATO → modelagem
```

O framework também possui estruturas específicas para rastreamento, como `_STG_Ocorrencias`, `_QA_Status` e `_QA_Ocorrencias`.

---

# 2. Estratégia de Diagnóstico

Antes de alterar o código, siga esta sequência:

```text
1. Identificar a camada
        ↓
2. Identificar a consulta
        ↓
3. Verificar o Schema
        ↓
4. Verificar o pipeline compilado
        ↓
5. Inspecionar ocorrências
        ↓
6. Corrigir a configuração ou implementação
        ↓
7. Executar novamente
```

Evite corrigir diretamente uma consulta de saída quando o problema estiver na configuração que gerou seu comportamento.

---

# 3. Erro na Compilação do Schema

## Sintoma

O pipeline apresenta erro durante a compilação do Schema.

Exemplos de situações:

```text
Operador não encontrado
Validação não encontrada
Tratamento não encontrado
Parâmetro inválido
```

## Causa provável

O operador utilizado no Schema não está cadastrado na estrutura de parâmetros correspondente.

O README identifica `tbParametrosTratamentos` e `tbParametrosValidacoes` como estruturas utilizadas para cadastro dos operadores.

## Diagnóstico

Verifique:

1. o nome do operador no Schema;
2. se o operador está cadastrado;
3. se ele está cadastrado como tratamento ou validação;
4. se a sintaxe utilizada está correta;
5. se os parâmetros foram informados corretamente.

Exemplo:

```text
Schema
    │
    ▼
TRIM;NOVOOPERADOR
             │
             ▼
       operador cadastrado?
```

## Correção

Cadastre o operador na estrutura correspondente ou corrija a referência utilizada no Schema.

Depois, execute novamente a compilação.

---

# 4. Operador Existente, mas Pipeline Não Compila

## Sintoma

O operador parece estar cadastrado, mas a compilação continua falhando.

## Diagnóstico

Verifique a resolução do operador no pipeline.

O framework compila operadores através de estruturas como:

```text
fxPipelineCompilarOperadores
cfgOperadores
```

Durante a compilação são resolvidos:

* código do operador;
* parâmetros;
* função executável.

Verifique principalmente:

```text
Código
Parâmetros
Categoria
Função associada
```

## Correção

Compare a definição do operador no Schema com o cadastro utilizado pelo compilador.

Um erro de nomenclatura pode impedir a resolução mesmo que exista uma função M correspondente.

---

# 5. Erro Relacionado a Parâmetros

## Sintoma

Um operador sem parâmetros funciona, mas uma versão parametrizada apresenta erro.

Exemplo:

```text
PADLEFT
```

funciona, enquanto:

```text
PADLEFT(3,0,P)
```

falha.

## Diagnóstico

O pipeline extrai os parâmetros da definição textual antes de resolver o operador.

Verifique:

1. nome do operador;
2. quantidade de parâmetros;
3. ordem dos parâmetros;
4. sintaxe;
5. tipo esperado dos parâmetros.

## Correção

Compare a chamada no Schema com a assinatura esperada pelo operador documentado em `operators.md`.

---

# 6. Erro na STG

## Sintoma

A consulta `stg*` apresenta erro durante a preparação ou tipagem.

## Diagnóstico

A STG utiliza o Schema para:

* preparar a estrutura;
* aplicar tipos;
* ordenar colunas;
* registrar problemas de tipagem.

A implementação registra falhas em `_STG_Ocorrencias` sem simplesmente transformar valores inválidos em `null`.

Verifique:

1. nome da tabela no Schema;
2. nome das colunas;
3. tipo configurado;
4. existência das colunas na fonte;
5. `_STG_Ocorrencias`.

## Correção

Corrija o Schema ou a estrutura da fonte conforme a causa identificada.

---

# 7. Problema de Conversão de Tipo

## Sintoma

Uma coluna apresenta valores incompatíveis com o tipo configurado.

Exemplo conceitual:

```text
Tipo esperado:
DATE

Valor:
"texto inválido"
```

## Diagnóstico

Verifique:

```text
stgClientes
    │
    ▼
_STG_Ocorrencias
```

A ocorrência deve ser utilizada para identificar quais registros apresentaram problema.

## Correção

Existem três possibilidades:

### Corrigir a origem

Se o valor estiver incorreto na fonte.

### Ajustar o tratamento

Se o valor puder ser convertido através de uma transformação apropriada.

### Ajustar o Schema

Se o tipo configurado não representar corretamente o domínio da coluna.

Não altere o tipo apenas para eliminar o erro sem verificar o significado do dado.

---

# 8. Coluna Não Encontrada

## Sintoma

A STG ou outra etapa apresenta erro porque uma coluna esperada não existe.

## Diagnóstico

Compare:

```text
Fonte
   │
   ├── Colunas reais
   │
   ▼
Schema
   │
   └── Colunas esperadas
```

Verifique diferenças de:

* nome;
* acentuação;
* espaços;
* caracteres especiais;
* alterações na fonte.

## Correção

Corrija a origem ou o Schema conforme o caso.

O problema deve ser resolvido na definição estrutural, e não através de alterações arbitrárias nas etapas posteriores.

---

# 9. Problema na TRN

## Sintoma

A tabela STG funciona, mas a TRN apresenta valores inesperados ou erro de processamento.

## Diagnóstico

Verifique:

```text
stg*
 │
 ▼
fxTrnAplicar
 │
 ▼
TratamentosPorColuna
```

O `fxTrnAplicar` compila os tratamentos definidos no Schema e os aplica por coluna.

Verifique:

1. tratamentos configurados;
2. ordem dos operadores;
3. parâmetros;
4. tipo da coluna;
5. combinação de operadores.

---

# 10. Tratamentos Produzem Resultado Incorreto

## Sintoma

O pipeline executa sem erro, mas o valor resultante não é o esperado.

## Exemplo

Schema:

```text
TRIM;UPPER
```

Resultado esperado:

```text
"JOAO"
```

mas o resultado observado não corresponde à regra esperada.

## Diagnóstico

Teste os operadores individualmente:

```text
TRIM
```

depois:

```text
UPPER
```

e finalmente:

```text
TRIM;UPPER
```

Isso permite identificar em qual operador ocorre a divergência.

## Correção

Verifique a configuração do Schema e a documentação do operador em `operators.md`.

Não altere o pipeline para corrigir uma regra que deveria estar representada no Schema.

---

# 11. Muitos Registros com Status `AVISO`

## Sintoma

A consulta QA retorna uma quantidade inesperadamente alta de:

```text
_QA_Status = AVISO
```

## Diagnóstico

Extraia os problemas:

```powerquery
Problemas =
    fxQaExtrairProblemas(qaClientes)
```

O README recomenda `fxQaExtrairProblemas` para auditar os registros classificados com problemas.

Analise:

* coluna;
* código da validação;
* mensagem;
* detalhes;
* severidade.

## Correção

Depois de identificar a causa:

* corrija os dados, se o problema estiver na fonte;
* ajuste o tratamento, se a normalização estiver insuficiente;
* ajuste a validação, se a regra estiver incorreta;
* ajuste a severidade, se a classificação estiver inadequada.

Não reduza a severidade apenas para esconder problemas de qualidade.

---

# 12. Muitos Registros com Status `ERRO`

## Sintoma

Grande quantidade de registros termina com:

```text
_QA_Status = ERRO
```

## Diagnóstico

Extraia os problemas:

```powerquery
Problemas =
    fxQaExtrairProblemas(qaClientes)
```

Depois agrupe mentalmente os problemas por:

```text
Coluna
Validação
Código
Severidade
```

Se a maioria dos erros estiver concentrada em uma única coluna, verifique primeiro a configuração dessa coluna no Schema.

## Correção

Corrija a causa na origem ou na configuração responsável pela regra.

---

# 13. Como Inspecionar as Ocorrências QA

A estrutura QA mantém informações de diagnóstico em:

```text
_QA_Ocorrencias
```

e:

```text
_QA_Status
```

As ocorrências podem conter informações como:

```text
Código
Severidade
Coluna
Tipo
Descrição
```

O registro original é preservado durante a validação.

Isso permite analisar o problema sem perder o contexto do registro.

---

# 14. Dados Foram Removidos na Normalização

## Sintoma

A quantidade de registros em `nrm*` é menor do que em `qa*`.

## Causa provável

A etapa NRM está recebendo apenas os registros selecionados após QA.

Um fluxo comum é:

```powerquery
ClientesValidos =
    fxQaFiltrarPorStatus(
        qaClientes,
        "OK"
    );

nrmClientes =
    fxNrmAplicar(
        ClientesValidos,
        "tbClientes"
    );
```

A filtragem por status ocorre antes da normalização.

## Diagnóstico

Compare:

```text
qaClientes
    │
    ├── OK
    ├── AVISO
    └── ERRO
```

com:

```text
ClientesValidos
    │
    └── OK
```

Se registros `AVISO` ou `ERRO` não deveriam ser removidos, revise o critério de filtragem.

## Correção

Ajuste:

```text
fxQaFiltrarPorStatus
```

ou a regra de qualidade responsável pelo status.

---

# 15. Dados Duplicados na NRM

## Sintoma

Registros que deveriam ser únicos permanecem duplicados.

## Diagnóstico

Verifique as `ChavesNegocio` configuradas no Schema.

Exemplo:

```text
ChavesNegocio:
CPF
```

A NRM utiliza essas chaves para deduplicação.

Verifique:

1. se a chave está configurada;
2. se a coluna existe;
3. se o valor foi corretamente tratado na TRN;
4. se a chave realmente identifica o registro.

## Correção

Corrija a definição de `ChavesNegocio` ou os tratamentos necessários para produzir uma chave consistente.

---

# 16. Duplicidades Inesperadas

## Sintoma

A NRM remove registros que parecem diferentes.

## Causa provável

Os registros possuem os mesmos valores nas colunas definidas como chave de negócio.

Exemplo:

```text
Chave:
CPF
```

Dois registros com o mesmo CPF são considerados duplicados para essa regra.

## Diagnóstico

Compare os valores das colunas que formam a chave.

## Correção

Se a identidade do registro exigir mais campos, utilize uma chave composta.

Exemplo:

```text
ChavesNegocio:
CodigoCliente
CodigoProduto
```

A definição deve representar a identidade de negócio real do registro.

---

# 17. Problema na DIM

## Sintoma

A dimensão apresenta:

* IDs inesperados;
* colunas fora da ordem;
* registros duplicados;
* desempenho ruim em consultas dependentes.

## Diagnóstico

Verifique primeiro:

```text
nrm*
 │
 ▼
dim*
```

A dimensão deve consumir dados já normalizados.

O README documenta o uso de:

```text
Table.AddIndexColumn
Table.ReorderColumns
Table.Buffer
```

na construção das dimensões.

## Correção

Verifique primeiro a NRM antes de modificar a dimensão.

Um problema de duplicidade ou qualidade na dimensão pode ser consequência de uma etapa anterior.

---

# 18. Problema em `fato*`

## Sintoma

A tabela fato não encontra uma dimensão ou retorna registros incompletos.

## Diagnóstico

Verifique os relacionamentos realizados com `Table.NestedJoin`.

Fluxo:

```text
nrmVendas
   │
   ├── Join dimClientes
   │
   └── Join dimProdutos
```

O README utiliza `Table.NestedJoin` para esses relacionamentos.

Verifique:

1. chave utilizada no fato;
2. chave correspondente na dimensão;
3. tratamento dos códigos;
4. existência da chave na dimensão;
5. tipo dos campos utilizados no relacionamento.

---

# 19. Performance Ruim

## Sintoma

O refresh do Power Query está mais lento que o esperado.

## Primeira verificação

Procure por:

```text
Table.Buffer
```

em tabelas operacionais.

O framework recomenda bufferização de estruturas de metadados e listas de operadores, evitando `Table.Buffer` indiscriminado em tabelas de dados operacionais.

## Segunda verificação

Verifique se o mesmo pipeline está sendo compilado repetidamente.

O framework utiliza pipelines por coluna pré-compilados para reduzir overhead.

## Terceira verificação

Verifique se operações estão sendo aplicadas somente às colunas configuradas.

O processamento foi projetado para trabalhar por coluna e evitar transformações desnecessárias.

## Quarta verificação

Verifique se existem múltiplas passagens sobre os mesmos dados.

O QA, por exemplo, foi estruturado para acumular ocorrências em uma única passagem.

---

# 20. `Table.Buffer` Não Melhorou a Performance

## Sintoma

Depois de adicionar:

```powerquery
Table.Buffer(Tabela)
```

a consulta continua lenta ou ficou mais lenta.

## Diagnóstico

A bufferização não deve ser considerada uma otimização automática.

No framework, a estratégia é:

```text
Metadados
   ↓
Bufferização seletiva

Dados operacionais
   ↓
Evitar bufferização indiscriminada
```

O uso de `Table.Buffer` deve possuir uma justificativa relacionada ao padrão de avaliação e reutilização da tabela.

## Correção

Remova o buffer se ele não apresentar benefício mensurável.

Depois compare o tempo de execução antes e depois da alteração.

---

# 21. Erro em Validação

## Sintoma

Uma validação apresenta erro durante o QA.

## Diagnóstico

Verifique:

```text
Schema
  │
  ▼
Validações
  │
  ▼
ValidaçõesPorColuna
  │
  ▼
fxQaValidar
```

Confirme:

1. nome do validador;
2. parâmetros;
3. tipo da coluna;
4. configuração de `Obrigatório`;
5. cadastro do operador.

Os validadores disponíveis incluem `REQUIRED`, `LIST`, `SIZE`, `MIN`, `MAX`, `INTERVAL`, `EMAIL`, `URL`, `CEPVAL`, `CPFVAL`, `CNPJVAL` e `PHONEVAL`.

---

# 22. `REQUIRED` Está Sendo Executado sem Estar no Schema

## Sintoma

Uma coluna possui:

```text
Obrigatório = SIM
```

mas `REQUIRED` não aparece explicitamente no campo `Validações`.

## Comportamento esperado

O pipeline pode incluir `REQUIRED` implicitamente durante a compilação.

```text
Obrigatório = SIM
       │
       ▼
REQUIRED implícito
```

Esse comportamento é documentado na compilação de `fxPipelineCompilarColuna`.

## Diagnóstico

Verifique o pipeline compilado e as validações associadas à coluna.

---

# 23. Status Não é o Esperado

## Sintoma

O registro aparece como:

```text
AVISO
```

quando deveria ser:

```text
ERRO
```

ou vice-versa.

## Diagnóstico

Verifique:

```text
cfgParametrosSeveridades
```

Essa estrutura controla a classificação de problemas como `AVISO` ou `ERRO`.

Também verifique qual validação gerou a ocorrência.

## Correção

Ajuste a configuração da severidade em vez de alterar a lógica do QA.

---

# 24. Pipeline Funciona para uma Tabela, mas Não para Outra

## Sintoma

Por exemplo:

```text
tbClientes → funciona
tbProdutos → erro
```

## Diagnóstico

Compare os Schemas.

Verifique:

```text
Tabela
Colunas
Tipos
Tratamentos
Validações
ChavesNegocio
```

Como o pipeline é compilado a partir do Schema, uma diferença de configuração pode gerar um pipeline diferente para cada tabela.

## Correção

Compare `cfgPipeline` das duas tabelas e identifique a primeira diferença relevante.

---

# 25. Consulta Funciona Isoladamente, mas Falha no Pipeline Completo

## Sintoma

Uma consulta intermediária funciona quando atualizada individualmente, mas falha quando o projeto inteiro é atualizado.

## Diagnóstico

Verifique dependências:

```text
SRC
 ↓
STG
 ↓
TRN
 ↓
QA
 ↓
NRM
```

O problema pode estar relacionado a:

* configuração compartilhada;
* Schema;
* pipeline;
* parâmetros;
* dados produzidos por uma etapa anterior.

## Estratégia

Execute e valide progressivamente:

```text
stgClientes
    ↓
trnClientes
    ↓
qaClientes
    ↓
nrmClientes
```

Esse fluxo corresponde também ao checklist de validação do projeto.

---

# 26. Como Isolar o Problema

Quando o pipeline completo falhar, não tente diagnosticar todas as etapas simultaneamente.

Utilize:

```text
STG
 │
 ├── funciona?
 │       │
 │       ├── NÃO → investigar STG
 │       └── SIM
 │
 ▼
TRN
 │
 ├── funciona?
 │       │
 │       ├── NÃO → investigar TRN
 │       └── SIM
 │
 ▼
QA
 │
 ├── funciona?
 │       │
 │       ├── NÃO → investigar QA
 │       └── SIM
 │
 ▼
NRM
```

Isso reduz significativamente o espaço de investigação.

---

# 27. Checklist de Diagnóstico

## Configuração

* [ ] A tabela está configurada no Schema?
* [ ] As colunas estão corretas?
* [ ] Os tipos estão corretos?
* [ ] Os tratamentos estão corretos?
* [ ] As validações estão corretas?
* [ ] As chaves de negócio estão corretas?

## Operadores

* [ ] O operador está cadastrado?
* [ ] O operador está na categoria correta?
* [ ] Os parâmetros estão corretos?
* [ ] A função associada está correta?

## STG

* [ ] A estrutura da fonte corresponde ao Schema?
* [ ] Existem `_STG_Ocorrencias`?
* [ ] Existem problemas de tipagem?

## TRN

* [ ] Os tratamentos estão sendo aplicados?
* [ ] A ordem dos operadores está correta?
* [ ] Os valores resultantes estão corretos?

## QA

* [ ] `_QA_Status` está correto?
* [ ] `_QA_Ocorrencias` contém os problemas esperados?
* [ ] As severidades estão corretas?
* [ ] Os registros deveriam ser filtrados?

## NRM

* [ ] As chaves de negócio estão corretas?
* [ ] Existem duplicidades?
* [ ] Registros foram removidos pela filtragem anterior?

## DIM/FATO

* [ ] As dimensões estão corretas?
* [ ] As chaves existem?
* [ ] Os `NestedJoin` encontram correspondências?
* [ ] O modelo possui os registros esperados?

## Performance

* [ ] Existem `Table.Buffer` desnecessários?
* [ ] O pipeline está sendo compilado previamente?
* [ ] Existem múltiplas passagens sobre a mesma tabela?
* [ ] O processamento está restrito às colunas configuradas?

---

# 28. Diagnóstico por Sintoma

| Sintoma                           | Primeira área a verificar       |
| --------------------------------- | ------------------------------- |
| Erro ao compilar Schema           | Cadastro dos operadores         |
| Coluna não encontrada             | Schema / fonte                  |
| Erro de conversão                 | STG / tipo                      |
| Valor transformado incorretamente | TRN / operador                  |
| Muitos `AVISO`                    | QA / validações                 |
| Muitos `ERRO`                     | QA / validações                 |
| Registro desapareceu              | Filtragem QA                    |
| Duplicidade na NRM                | Chaves de negócio               |
| Dimensão incorreta                | NRM / DIM                       |
| Fato sem correspondência          | DIM / `NestedJoin`              |
| Refresh lento                     | Buffer / passagens / compilação |
| `REQUIRED` inesperado             | `Obrigatório` no Schema         |
| Severidade incorreta              | `cfgParametrosSeveridades`      |

---

# 29. Procedimento Recomendado para Novos Erros

Quando um problema ainda não estiver documentado, registre-o seguindo este padrão:

```text
### Problema: [descrição curta]

**Sintoma:**
O que está acontecendo.

**Camada:**
STG / TRN / QA / NRM / DIM / FATO

**Causa provável:**
Hipótese inicial.

**Diagnóstico:**
Como confirmar a causa.

**Correção:**
O que deve ser alterado.

**Prevenção:**
Como evitar que o problema volte a ocorrer.
```

Esse formato mantém o documento consistente à medida que novos casos forem identificados.

---

# 30. Princípio de Diagnóstico

O princípio fundamental para investigar problemas no framework é:

```text
Não corrigir o sintoma antes de identificar a camada responsável.
```

O fluxo recomendado é:

```text
Sintoma
   │
   ▼
Camada
   │
   ▼
Configuração
   │
   ▼
Pipeline
   │
   ▼
Execução
   │
   ▼
Resultado
```

A arquitetura do framework foi projetada justamente para permitir esse isolamento.

---

# 31. Documentação Relacionada

* [Getting Started](getting-started.md) — configuração inicial.
* [Architecture](architecture.md) — responsabilidades das camadas.
* [Schema](schema.md) — configuração do processamento.
* [Pipeline](pipeline.md) — compilação e execução.
* [Operators](operators.md) — tratamentos e validações.
* [Performance](performance.md) — diagnóstico e otimização.
* [Examples](examples.md) — exemplos de utilização.

---

# 32. Resumo

A estratégia de troubleshooting do framework pode ser resumida em:

```text
                 PROBLEMA
                    │
                    ▼
             Identificar camada
                    │
        ┌───────────┼───────────┐
        ▼           ▼           ▼
       STG         TRN          QA
        │           │           │
      Tipos      Tratamentos  Validações
        │           │           │
        └───────────┼───────────┘
                    │
                    ▼
                   NRM
                    │
                    ▼
                DIM / FATO
```

Os principais mecanismos de diagnóstico são:

```text
_STG_Ocorrencias
_QA_Status
_QA_Ocorrencias
fxQaExtrairProblemas
fxQaFiltrarPorStatus
cfgPipeline
cfgParametrosSeveridades
```

A abordagem recomendada é sempre **identificar primeiro a origem do comportamento e somente depois alterar a implementação**.
