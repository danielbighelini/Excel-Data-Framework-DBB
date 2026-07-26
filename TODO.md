Analisei o arquivo completo e considerei exclusivamente a implementação existente como referência. 

A boa notícia é que a arquitetura evoluiu bastante em relação às primeiras versões. Você já eliminou diversos gargalos clássicos (processamento registro a registro, múltiplos `Table.AddColumn`, vários `Record.ToTable`, etc.). Hoje os problemas restantes são mais sutis e estão concentrados em alguns pontos específicos.

Abaixo está o relatório priorizado.

---

# Prioridade 1 (alto impacto / baixo-médio esforço)

## 1. Pipeline reconstruído a cada chamada

### Atual

```text
fxStgAplicar
    ↓
fxStgPipelineGerar(schema)
```

`fxStgPipelineGerar` monta todo o pipeline toda vez que uma tabela é processada.

Ele consulta:

* cfgObjetos
* cfgTabelas
* cfgSchema
* Record.Combine
* Record.FromList
* etc.

Mesmo sendo estruturas em memória, isso acontece para **toda execução**. 

### Melhor

Ao invés de

> gerar o pipeline dinamicamente

utilize

> `cfgPipeline`

ou

> `cfgPipelines`

materializado uma única vez.

Exemplo:

```text
cfgPipeline[tbClientes]

cfgPipeline[tbProdutos]
```

---

### Impacto

Muito alto.

Principalmente quando dezenas de tabelas utilizam o framework.

---

### Esforço

Médio.

---

# Prioridade 2 (alto impacto / baixo esforço)

## 2. Procura linear da posição da coluna

Em várias funções ocorre:

```m
List.PositionOf(...)
```

por exemplo:

* fxStgPadronizar
* fxStgAplicarValidacoes

Cada coluna procura novamente sua posição.

Isso transforma o custo em aproximadamente

```text
O(colunas²)
```

---

### Melhor

Ao invés de

```text
List.PositionOf(...)
```

utilize

```text
Record

Coluna -> índice
```

Exemplo

```text
[
Nome = 0,
CPF = 1,
Cidade = 2
]
```

A busca passa a ser O(1).

---

### Impacto

Muito alto em tabelas largas.

---

### Esforço

Baixo.

---

# Prioridade 3 (alto impacto)

## 3. Table.ColumnNames chamado repetidamente

Diversas funções fazem

```m
Table.ColumnNames(tabela)
```

embora a tabela seja exatamente a mesma.

Exemplo

```
Preparar

↓

GarantirColunas

↓

RemoverColunas

↓

Padronizar

↓

Validar
```

Todas voltam a ler os nomes.

---

### Melhor

Ao invés de

```text
Table.ColumnNames(...)
```

em todas as funções

utilize

```text
Estado

[
Tabela,
Colunas
]
```

ou passe a lista de colunas.

---

### Impacto

Médio/alto.

---

### Esforço

Baixo.

---

# Prioridade 4 (alto impacto)

## 4. Record.FieldNames(pipeline)

Também aparece em praticamente todas as funções.

Exemplo

```
fxStgPadronizar

fxStgAplicarValidacoes

fxStgGarantirColunas

fxStgOrdenarColunas

fxStgAplicarTipos
```

Cada uma reconstrói

```text
Record.FieldNames(...)
```

---

### Melhor

Materializar uma vez:

```text
Pipeline

[
Definicoes,
Colunas
]
```

onde

```
Colunas =
List.Buffer(...)
```

---

### Impacto

Médio.

---

### Esforço

Baixo.

---

# Prioridade 5 (médio impacto)

## 5. Pipeline ainda contém muita informação repetida

Cada coluna guarda

```
Tipo

Obrigatório

Ordem

Tratamentos

Validações
```

Mesmo quando

```
Tratamentos = {}

Validações = {}

Tipo = any
```

Na prática, a maioria das colunas possui apenas valores padrão.

---

### Melhor

Ao invés de

```
Record.Combine
```

para todas as colunas

utilize

apenas os campos existentes.

---

### Impacto

Médio.

Principalmente consumo de memória.

---

### Esforço

Médio.

---

# Prioridade 6 (médio impacto)

## 6. fxStgAplicarTipos faz inferência individual

Quando

```
Tipo = any
```

cada coluna executa

```
fxStgIdentificarTipoColuna
```

que faz

```
Table.Column(...)
```

novamente.

Ou seja:

```
uma leitura da coluna
```

para inferir

e depois

```
nova leitura
```

para converter.

---

### Melhor

Ao invés de

```
inferir durante a conversão
```

utilize

```
pipeline já tipado
```

ou

pré-inferência única.

---

### Impacto

Médio.

---

### Esforço

Médio.

---

# Prioridade 7 (médio impacto)

## 7. Muitas reconstruções de tabela

O fluxo hoje faz aproximadamente

```
Preparar

↓

Garantir

↓

Remover

↓

Padronizar

↓

Validar

↓

Tipos

↓

Ordenar
```

Cada uma gera nova tabela.

Embora inevitável em parte, algumas poderiam compartilhar o mesmo estado.

---

### Melhor

Ao invés de

```
7 reconstruções
```

utilize

```
estado único

↓

materializar apenas no final
```

---

### Impacto

Médio.

---

### Esforço

Alto.

---

# Prioridade 8 (médio impacto)

## 8. Falta de Buffer do Pipeline

Hoje:

```
Pipeline =
fxStgPipelineGerar(...)
```

O Record nunca recebe

```
Record.Buffer
```

(não existe)

nem seus componentes recebem Buffer.

As listas internas

```
Tratamentos

Validações
```

já recebem `List.Buffer` durante a construção do schema, o que é positivo. Porém, a lista de colunas do pipeline e o próprio catálogo de pipelines não estão materializados. 

---

### Melhor

Criar

```
cfgPipeline

↓

listas bufferizadas
```

---

### Impacto

Médio.

---

### Esforço

Baixo.

---

# Prioridade 9 (baixo impacto)

## 9. Buffer desnecessário

Existem alguns

```
Table.Buffer(...)
```

em consultas pequenas

como

```
cfgParametrosTratamentos

cfgParametrosValidacoes

cfgParametrosSeveridades
```

Como essas consultas já terminam em `Record`, o ganho tende a ser pequeno.

---

### Melhor

Reavaliar caso a caso.

---

### Impacto

Baixo.

---

### Esforço

Baixo.

---

# Prioridade 10 (baixo impacto)

## 10. Query Folding

O framework praticamente elimina qualquer possibilidade de Query Folding logo no início.

Os principais motivos são:

* `Table.ToColumns`
* `Table.FromColumns`
* `List.Transform`
* `List.Accumulate`
* `Table.Column`
* tratamento célula a célula

Isso não significa que esteja errado; é consequência natural de um framework de transformação orientado por metadados. Entretanto, quando a origem suporta folding (SQL Server, Oracle, PostgreSQL etc.), é recomendável manter todas as transformações que podem ser traduzidas para SQL antes da chamada de `fxStgAplicar`, deixando o framework responsável apenas pelas regras que realmente exigem processamento em M.

---

# Resumo executivo

| Prioridade | Ao invés de...                                    | Utilize...                               | Impacto                              | Esforço |
| ---------- | ------------------------------------------------- | ---------------------------------------- | ------------------------------------ | ------- |
| ⭐⭐⭐⭐⭐      | Gerar `Pipeline` em `fxStgPipelineGerar` toda vez | `cfgPipeline` materializado              | Muito alto                           | Médio   |
| ⭐⭐⭐⭐⭐      | `List.PositionOf` para localizar colunas          | Mapa `Coluna → Índice`                   | Muito alto                           | Baixo   |
| ⭐⭐⭐⭐       | `Table.ColumnNames` repetidamente                 | Lista de colunas compartilhada           | Alto                                 | Baixo   |
| ⭐⭐⭐⭐       | `Record.FieldNames` repetidamente                 | Lista de colunas bufferizada no pipeline | Alto                                 | Baixo   |
| ⭐⭐⭐        | Pipeline com muitos valores padrão                | Pipeline compacto                        | Médio                                | Médio   |
| ⭐⭐⭐        | Inferir tipo durante `fxStgAplicarTipos`          | Tipo previamente resolvido               | Médio                                | Médio   |
| ⭐⭐⭐        | Reconstruir tabela em cada etapa                  | Estado único até a materialização final  | Médio                                | Alto    |
| ⭐⭐⭐        | Pipeline não materializado                        | `cfgPipeline` pré-construído             | Médio                                | Baixo   |
| ⭐⭐         | Buffers em consultas pequenas                     | Revisar necessidade                      | Baixo                                | Baixo   |
| ⭐⭐         | Aplicar framework antes do folding                | Executar após as etapas com folding      | Baixo a médio (dependente da origem) | Baixo   |

A principal conclusão da inspeção é que o maior gargalo remanescente não está mais nos tratamentos ou validações em si, mas na **infraestrutura do pipeline**. A criação repetida do pipeline e as buscas lineares por metadados (colunas e índices) tendem a consumir mais tempo do que a execução das próprias funções de tratamento em muitas cargas de trabalho. A consolidação do pipeline em um `cfgPipeline` pré-materializado e a eliminação dessas buscas repetidas seriam, na minha avaliação, as duas mudanças com melhor relação entre ganho de desempenho e esforço de implementação.





25/07/2026
Analisei o arquivo completo fornecido, utilizando-o como única base de referência para a avaliação. A inspeção foi focada em `fxStgAplicar`, seu pipeline (`fxStgPipelineGerar`) e todas as estruturas `cfg*` relacionadas ao processamento da Stage. 

A conclusão é que o principal problema não está em uma única função lenta, mas na combinação de diversas decisões arquiteturais que multiplicam o custo do processamento. Há vários pontos em que o framework deixa de aproveitar operações vetorizadas do motor do Power Query e passa a reconstruir listas, registros e tabelas repetidamente.

# Prioridade 1 — Muito Alto Impacto

| Prioridade | Estrutura                             | Método atual                                               | Melhor abordagem                                                | Esforço |
| ---------- | ------------------------------------- | ---------------------------------------------------------- | --------------------------------------------------------------- | ------- |
| 1          | `fxStgAplicar` + `fxStgPipelineGerar` | Gera o pipeline novamente a cada chamada                   | utilizar `cfgPipeline` já materializado                         | Baixo   |
| 1          | `fxStgPadronizar`                     | Reconstrói toda a tabela a cada coluna tratada             | aplicar transformações por `Table.TransformColumns`             | Alto    |
| 1          | `fxStgAplicarValidacoes`              | Converte tabela inteira para listas e reconstrói novamente | validar diretamente sobre colunas usando transformações nativas | Alto    |
| 1          | `fxStgAplicarTipos`                   | Executa função personalizada por célula                    | utilizar `Table.TransformColumnTypes` sempre que possível       | Médio   |

---

## 1) Pipeline é reconstruído em toda execução

Hoje:

```
fxStgAplicar
    ↓
fxStgPipelineGerar
    ↓
cfgSchema
cfgTabelas
Record.Combine
Record.FromList
```

Ou seja, toda Stage recompõe um pipeline completo.

Entretanto o projeto já possui:

```
cfgPipeline
```

que representa exatamente esse objeto.

Hoje existe duplicação de trabalho.

### Melhor

Ao invés de

```
fxStgPipelineGerar(schema)
```

utilizar

```
cfgPipeline[schema]
```

O ganho é imediato porque elimina:

* Record.Combine
* Record.FromList
* Record.FieldOrDefault em cascata
* geração de listas

em toda execução.

**Impacto esperado:** Muito alto.

---

# 2) Padronização reconstrói a tabela inteira diversas vezes

Hoje:

```
Table.ToColumns()

↓

List.Accumulate

↓

List.ReplaceRange

↓

Table.FromColumns()
```

Para cada coluna tratada.

Isso significa que para N colunas:

```
lista completa

↓

lista completa

↓

lista completa

↓

lista completa
```

O custo cresce muito rapidamente.

O Power Query possui um operador próprio exatamente para isso:

```
Table.TransformColumns()
```

que altera apenas a coluna desejada.

Ao invés de

```
ToColumns

↓

ReplaceRange

↓

FromColumns
```

utilizar

```
Table.TransformColumns
```

Impacto extremamente alto.

---

# 3) Validação faz diversas reconstruções de listas

Em

```
fxStgAplicarValidacoes
```

ocorre:

```
Table.ToColumns

↓

List.Accumulate

↓

List.Transform

↓

List.Combine

↓

Table.FromColumns
```

Além disso existe outro loop interno em

```
fxStgAplicarValidacoesColuna
```

ou seja:

```
validação

↓

linha

↓

coluna

↓

reconstrução
```

O volume de alocação de memória é enorme.

É provavelmente o maior gargalo do framework.

---

# 4) Conversão de tipos é feita célula por célula

Hoje:

```
Table.TransformColumns

↓

try

↓

fxConversor

↓

Record.Field

↓

Conversor
```

para absolutamente todas as células.

Quando o tipo já é conhecido, o Power Query possui:

```
Table.TransformColumnTypes()
```

que é muito mais otimizado.

A função atual somente deveria ser usada para tipos especiais.

---

# Prioridade 2 — Alto impacto

| Estrutura                    | Método atual                       | Melhor abordagem                       | Esforço |
| ---------------------------- | ---------------------------------- | -------------------------------------- | ------- |
| `fxStgAplicarTipos`          | infere tipo toda execução          | armazenar tipo inferido no pipeline    | Médio   |
| `fxStgIdentificarTipoColuna` | faz leitura da coluna              | inferir apenas uma vez                 | Médio   |
| `fxStgPreparar`              | Record.FieldValues em todas linhas | filtro mais simples para linhas vazias | Médio   |

---

## 5) Inferência de tipos ocorre em todas execuções

Mesmo usando o mesmo Schema:

```
Table.Column()

↓

List.FirstN()

↓

List.Transform()

↓

List.Distinct()
```

é executado novamente.

Isso deveria acontecer apenas uma vez.

---

## 6) Preparação verifica todas as células

```
Record.FieldValues

↓

List.Transform

↓

List.AnyTrue
```

para todas as linhas.

É caro.

Dependendo do cenário pode representar boa parte do tempo quando existem muitas colunas.

---

# Prioridade 3 — Médio impacto

| Estrutura                      | Problema                         | Esforço |
| ------------------------------ | -------------------------------- | ------- |
| `fxStgAplicarValidacoesColuna` | muitas listas intermediárias     | Alto    |
| `cfgPipeline`                  | existe mas não é utilizado       | Baixo   |
| `fxStgAplicarTratamentos`      | vários List.Transform sucessivos | Médio   |
| `fxStgOrdenarColunas`          | sempre executa                   | Baixo   |

---

## 7) Tratamentos fazem vários passes na coluna

Hoje:

```
TRIM

↓

CLEAN

↓

UPPER

↓

LOWER

↓

...
```

cada tratamento executa um:

```
List.Transform
```

novo.

Exemplo:

5 tratamentos

↓

5 varreduras completas da coluna.

O ideal seria compilar o pipeline em uma função única.

---

## 8) Ordenação sempre acontece

Mesmo quando a tabela já está na ordem correta.

```
Table.ReorderColumns
```

não é gratuito.

Pode ser evitado comparando previamente a ordem.

---

# Buffers

## Bem utilizados

Observei bom uso em:

* `cfgSchema`
* `cfgParametros`
* `cfgParametrosTratamentos`
* `cfgParametrosValidacoes`
* `cfgParametrosTipos`
* `cfgParametrosBooleanos`
* `cfgParametrosRESTHeaders`
* listas de parâmetros do schema

Não identifiquei ausência significativa de `Table.Buffer` nesses objetos.

---

## Onde buffers não resolverão

Os maiores gargalos não decorrem da ausência de `Buffer`.

Eles decorrem principalmente de:

```
Table

↓

List

↓

Table

↓

List

↓

Table
```

repetidamente.

Adicionar mais buffers nesses pontos tende apenas a aumentar o consumo de memória.

---

# Query Folding

As seguintes funções interrompem completamente qualquer possibilidade de folding:

* `Table.ToColumns`
* `Table.FromColumns`
* `Table.Column`
* `Table.ToRecords`
* `Record.FieldValues`
* `List.Accumulate`
* `List.Transform`
* `List.ReplaceRange`

Na prática, a partir de `fxStgPadronizar` e `fxStgAplicarValidacoes`, o processamento passa a ocorrer integralmente no motor M.

Para Excel isso não altera muito, mas para SQL Server, Oracle, PostgreSQL e outras fontes impede que parte do trabalho seja delegada ao servidor.

---

# Fragilidades arquiteturais

### `cfgPipeline` está praticamente inutilizado

Existe um catálogo completo de pipelines, porém `fxStgAplicar` continua chamando `fxStgPipelineGerar`.

É uma redundância clara.

---

### Pipeline recomposto em tempo de execução

Há muitos:

```
Record.Combine

Record.FromList

Record.FieldOrDefault
```

durante a execução.

Esses objetos poderiam ser completamente estáticos.

---

### Excesso de conversões

O framework alterna frequentemente entre:

```
table

↓

list

↓

record

↓

table
```

quando poderia permanecer em formato tabular durante quase todo o processamento.

# Resumo priorizado

| Ao invés de                                                   | Utilize                                       | Estruturas                    |
| ------------------------------------------------------------- | --------------------------------------------- | ----------------------------- |
| gerar pipeline com `fxStgPipelineGerar`                       | `cfgPipeline`                                 | `fxStgAplicar`, `cfgPipeline` |
| reconstruir tabelas via `Table.ToColumns`/`Table.FromColumns` | `Table.TransformColumns`                      | `fxStgPadronizar`             |
| validar convertendo tudo para listas                          | validação tabular por coluna                  | `fxStgAplicarValidacoes`      |
| converter tipos com função por célula                         | `Table.TransformColumnTypes` quando aplicável | `fxStgAplicarTipos`           |
| inferir tipos a cada execução                                 | armazenar no pipeline                         | `fxStgIdentificarTipoColuna`  |
| executar um `List.Transform` por tratamento                   | compilar tratamentos em uma única função      | `fxStgAplicarTratamentos`     |
| ordenar sempre                                                | ordenar apenas quando necessário              | `fxStgOrdenarColunas`         |

## Avaliação do esforço

* **Baixo:** substituir `fxStgPipelineGerar` por `cfgPipeline`; evitar ordenação desnecessária.
* **Médio:** refatorar aplicação de tipos e inferência de tipos.
* **Alto:** reescrever `fxStgPadronizar` e `fxStgAplicarValidacoes` para operar de forma predominantemente tabular, eliminando o ciclo repetitivo de conversões entre tabelas, listas e registros.

Pela análise do arquivo, esses dois últimos componentes (`fxStgPadronizar` e `fxStgAplicarValidacoes`) concentram a maior parte do custo computacional do framework atual e são os candidatos com maior potencial de redução do tempo total de processamento. 
