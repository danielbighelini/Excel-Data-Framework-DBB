# Performance

## 1. Visão Geral

O Excel Data Framework DBB foi projetado para reduzir processamento desnecessário no Power Query por meio de uma combinação de:

* processamento por coluna;
* uso de operações nativas de tabela;
* pré-compilação de pipelines;
* bufferização seletiva de metadados;
* redução de passagens sobre os dados;
* processamento de validações em uma única passagem;
* aproveitamento da avaliação lazy do Power Query.

O princípio geral é:

> **Processar apenas o que é necessário, no menor número possível de passagens e evitando materializações desnecessárias.**

A arquitetura de performance pode ser resumida como:

```text
                    CONFIGURAÇÃO
                         │
                         ▼
                 Pipeline pré-compilado
                         │
                         ▼
              Processamento por coluna
                         │
              ┌──────────┴──────────┐
              ▼                     ▼
             TRN                   QA
              │                     │
       Transformações        Validação única
              │                     │
              └──────────┬──────────┘
                         ▼
                        NRM
```

---

# 2. Princípios de Performance

As principais estratégias utilizadas pelo framework são:

1. **Processar por coluna em vez de reconstruir registros individualmente.**
2. **Pré-compilar regras antes de processar os dados.**
3. **Evitar bufferização indiscriminada.**
4. **Reduzir o número de passagens sobre as tabelas.**
5. **Acumular informações relacionadas durante uma única passagem quando possível.**
6. **Aplicar transformações somente às colunas configuradas.**
7. **Preservar a avaliação lazy do Power Query sempre que possível.**

Esses princípios são complementares.

Uma otimização isolada não deve ser aplicada sem considerar o efeito sobre o restante do pipeline.

---

# 3. Processamento por Coluna

Uma das principais estratégias do framework é executar transformações utilizando operações nativas de tabela.

A camada TRN utiliza:

```powerquery
Table.TransformColumns
```

para aplicar os tratamentos configurados por coluna.

Conceitualmente:

```text
Tabela
  │
  ├── Coluna A → pipeline A
  ├── Coluna B → pipeline B
  ├── Coluna C → pipeline C
  └── Coluna D → sem processamento
```

Em vez de implementar uma sequência de operações manualmente para cada registro, o framework utiliza a estrutura de transformação por coluna.

---

## 3.1 Benefícios

O processamento por coluna permite:

* aplicar somente as transformações necessárias;
* evitar reconstruções completas da tabela;
* manter a lógica associada à configuração de cada coluna;
* utilizar operações nativas do Power Query.

O framework procura, portanto, manter o processamento alinhado à estrutura declarada no Schema.

---

# 4. Processamento Somente das Colunas Configuradas

Nem todas as colunas precisam receber todas as operações.

O pipeline compilado mantém as definições por coluna:

```text
TratamentosPorColuna
ValidaçõesPorColuna
TiposPorColuna
```

Isso permite que o framework determine exatamente quais operações precisam ser aplicadas.

Conceitualmente:

```text
Schema
  │
  ├── Nome → TRIM;PROPER
  ├── CPF  → TRIM;DIGITS
  ├── Cidade → TRIM
  └── Observacao → nenhuma
```

A coluna `Observacao`, por exemplo, não precisa receber uma transformação apenas porque pertence à mesma tabela.

---

# 5. Pré-compilação do Pipeline

A interpretação das configurações é separada da execução dos dados.

O framework compila previamente as definições do Schema em estruturas de execução.

```text
Schema
  │
  ▼
Compilação
  │
  ▼
cfgPipeline
  │
  ├── TiposPorColuna
  ├── TratamentosPorColuna
  ├── ValidaçõesPorColuna
  ├── Ordem
  └── ChavesNegocio
  │
  ▼
Execução
```

Os pipelines por coluna são pré-compilados para reduzir o overhead durante o processamento.

---

## 5.1 Compilação de operadores

Durante a compilação, o framework resolve os operadores definidos no Schema.

Conceitualmente:

```text
TRIM;DIGITS
     │
     ▼
Interpretação
     │
     ▼
Resolução dos operadores
     │
     ▼
Pipeline executável
```

A compilação evita que o mecanismo precise resolver novamente a mesma configuração textual durante o processamento dos dados.

---

# 6. Bufferização Estratégica

O framework utiliza `List.Buffer` e `Table.Buffer`, mas a estratégia é **seletiva**.

Exemplos:

```powerquery
Table.Buffer(stgSchema)
```

e:

```powerquery
List.Buffer(...)
```

A principal regra é:

> **Bufferizar estruturas de configuração e metadados quando isso reduz reavaliações, mas evitar bufferizar indiscriminadamente as tabelas de dados operacionais.**

---

# 7. Bufferização de Metadados

Estruturas como:

* Schema;
* listas de operadores;
* configurações;
* pipelines;

podem ser reutilizadas diversas vezes durante a compilação e execução.

Nesses casos, a bufferização pode evitar avaliações repetidas.

Exemplo:

```powerquery
Table.Buffer(stgSchema)
```

ou:

```powerquery
List.Buffer(Operadores)
```

O objetivo é manter em memória uma estrutura pequena e reutilizável.

---

# 8. Evitar `Table.Buffer` em Dados Operacionais

`Table.Buffer` não deve ser tratado como uma otimização universal.

Aplicar:

```powerquery
Table.Buffer(Tabela)
```

a cada tabela operacional pode aumentar:

* consumo de memória;
* tempo de materialização;
* custo de processamento;
* pressão sobre o mecanismo de avaliação.

Além disso, pode impedir otimizações da origem quando estas estiverem disponíveis.

Por isso, o framework evita utilizar `Table.Buffer` indiscriminadamente sobre tabelas de dados.

A preferência é:

```text
Metadados pequenos
      ↓
Bufferização seletiva


Dados operacionais
      ↓
Manter avaliação lazy quando possível
```

---

# 9. Lazy Evaluation

O Power Query utiliza avaliação lazy.

Isso significa que uma expressão não precisa necessariamente ser materializada no momento em que é definida.

O framework procura preservar esse comportamento sempre que a materialização antecipada não trouxer benefício claro.

Conceitualmente:

```text
Expressão
   │
   ▼
Avaliação lazy
   │
   ▼
Materialização somente quando necessária
```

Essa característica é especialmente relevante quando existem várias etapas encadeadas.

---

# 10. Evitar Materialização Desnecessária

Um dos objetivos da arquitetura é reduzir reconstruções desnecessárias de tabelas.

Em vez de criar múltiplas estruturas intermediárias para cada pequena operação, o framework procura agrupar operações relacionadas.

Exemplo conceitual:

```text
Menos eficiente:

Tabela
  ↓
Transformação 1
  ↓
Tabela
  ↓
Transformação 2
  ↓
Tabela
  ↓
Transformação 3
  ↓
Tabela
```

Preferência arquitetural:

```text
Tabela
  ↓
Pipeline de operações
  ↓
Resultado
```

Isso é particularmente relevante para operações de transformação por coluna.

---

# 11. STG

A camada STG utiliza `Table.TransformColumns` para aplicar os tipos por coluna.

O fluxo é:

```text
Fonte
  │
  ▼
Preparação estrutural
  │
  ▼
Aplicação de tipos
  │
  ▼
Ordenação
  │
  ▼
STG
```

A implementação procura realizar a preparação estrutural e aplicação dos tipos de forma organizada, evitando passagens desnecessárias.

---

## 11.1 Tipagem seletiva

O framework aplica tipos de acordo com `TiposPorColuna`.

Colunas que não precisam de transformação de tipo não devem receber operações de conversão desnecessárias.

Em particular, colunas configuradas como `type any` não recebem transformações de tipo desnecessárias.

---

# 12. TRN

A camada TRN compila os tratamentos por coluna e os aplica utilizando `Table.TransformColumns`.

Exemplo conceitual:

```text
CPF
 │
 ├── TRIM
 └── DIGITS

Nome
 │
 ├── TRIM
 └── PROPER
```

O pipeline compilado permite que essas operações sejam executadas diretamente sem reconstruir a configuração a cada registro.

---

# 13. QA em Uma Única Passagem

A camada QA utiliza uma estratégia específica para reduzir o número de passagens sobre os dados.

`fxQaValidar` percorre cada registro e acumula as ocorrências em uma estrutura `_QA`.

Conceitualmente:

```text
Registro
   │
   ▼
Validações
   │
   ├── ocorrência 1
   ├── ocorrência 2
   └── ocorrência N
   │
   ▼
_QA
   │
   ├── Status
   └── Ocorrencias
```

O status final é determinado durante a mesma passagem.

---

## 13.1 Estrutura intermediária `_QA`

Em vez de criar imediatamente várias colunas auxiliares para cada validação, o framework acumula as informações em um único campo:

```text
_QA
```

Esse campo contém as informações necessárias para determinar:

```text
_QA_Status
_QA_Ocorrencias
```

A expansão ocorre posteriormente, quando necessária.

---

## 13.2 Benefício

A estratégia reduz a quantidade de passagens sobre a tabela e evita criar estruturas intermediárias para cada validação.

O fluxo é:

```text
Tabela
  │
  ▼
Uma passagem
  │
  ▼
_QA
  │
  ├── Status
  └── Ocorrências
```

---

# 14. Filtragem Pós-QA

A validação não precisa remover registros durante sua execução.

O QA pode produzir:

```text
OK
AVISO
ERRO
```

A filtragem ocorre posteriormente por meio de:

```text
fxQaFiltrarPorStatus
```

Isso permite separar:

```text
Validação
     ↓
Diagnóstico
     ↓
Filtragem
```

Essa separação evita misturar a avaliação de qualidade com a decisão de quais registros continuarão no pipeline.

---

# 15. NRM

A camada NRM utiliza `Table.Distinct` para realizar a deduplicação baseada nas chaves de negócio.

Fluxo:

```text
QA
 │
 ▼
Registros selecionados
 │
 ▼
Chaves de negócio
 │
 ▼
Table.Distinct
 │
 ▼
NRM
```

A operação de deduplicação ocorre depois que os registros passaram pelas etapas anteriores do pipeline.

---

# 16. DIM / FATO

A modelagem dimensional pode utilizar estruturas já normalizadas.

Exemplos de operações utilizadas no modelo incluem:

```text
Table.AddIndexColumn
Table.ReorderColumns
Table.NestedJoin
```

O README também documenta o uso de bufferização em consultas dimensionais dependentes.

Exemplo:

```powerquery
Table.Buffer(Reordenada)
```

Essa estratégia deve ser aplicada somente quando a materialização da dimensão apresentar benefício efetivo para as consultas dependentes.

---

# 17. Número de Passagens

Um dos objetivos da arquitetura é reduzir o número de vezes que os mesmos dados precisam ser percorridos.

Uma estrutura conceitual é:

```text
STG
 ├── preparação
 └── tipagem

TRN
 └── tratamentos por coluna

QA
 └── validações em uma passagem

NRM
 └── deduplicação
```

Cada etapa possui uma responsabilidade definida.

A redução de passagens não significa necessariamente juntar todas as etapas em uma única operação.

A separação arquitetural continua sendo necessária para manter:

* legibilidade;
* rastreabilidade;
* testabilidade;
* isolamento de responsabilidades.

---

# 18. Performance x Separação de Responsabilidades

O framework não busca minimizar o número de consultas ou etapas a qualquer custo.

Existe um equilíbrio entre:

```text
Performance
     ↕
Arquitetura
     ↕
Manutenibilidade
```

Por exemplo, juntar STG, TRN e QA em uma única função poderia reduzir etapas intermediárias, mas também eliminaria fronteiras importantes do framework.

A estratégia adotada é otimizar **dentro de cada responsabilidade**, em vez de eliminar responsabilidades.

---

# 19. O que Deve Ser Bufferizado

Preferência:

```text
┌───────────────────────────────┐
│ Estruturas pequenas           │
│                               │
│ Schema                        │
│ Listas de operadores          │
│ Configurações                 │
│ Pipelines                     │
└───────────────┬───────────────┘
                │
                ▼
          Bufferização
```

Essas estruturas tendem a ser menores e podem ser reutilizadas.

---

# 20. O que Deve Ser Evitado

Evite bufferização indiscriminada de:

```text
Tabela operacional completa
        ↓
Table.Buffer()
```

principalmente quando:

* a tabela é grande;
* a fonte possui processamento próprio;
* a materialização não será reutilizada;
* a consulta poderia permanecer lazy;
* a operação pode aumentar significativamente o consumo de memória.

A regra prática é:

> **Não usar `Table.Buffer` apenas porque uma consulta está lenta.**

Primeiro deve-se identificar onde ocorre o custo.

---

# 21. Diagnóstico de Performance

Quando uma consulta estiver lenta, a análise deve considerar:

### 1. Fonte

Verificar se o custo está na origem.

### 2. Transformações

Verificar se existem operações repetidas ou desnecessárias.

### 3. Processamento por linha

Verificar se uma operação que poderia ser feita por coluna está sendo executada registro a registro.

### 4. Bufferização

Verificar se `Table.Buffer` ou outras materializações estão sendo aplicadas sem necessidade.

### 5. Pipeline

Verificar se as configurações estão sendo recompiladas repetidamente.

### 6. Número de passagens

Verificar se a mesma tabela está sendo percorrida várias vezes para obter informações que poderiam ser acumuladas em uma única passagem.

---

# 22. Checklist de Performance

Antes de otimizar uma consulta, verificar:

* [ ] As operações são executadas por coluna quando possível?
* [ ] O pipeline está sendo pré-compilado?
* [ ] Os operadores estão sendo resolvidos uma única vez?
* [ ] Apenas as colunas configuradas recebem transformações?
* [ ] Tipos desnecessários estão sendo evitados?
* [ ] A validação pode ser executada em uma única passagem?
* [ ] Existem tabelas intermediárias desnecessárias?
* [ ] Existe `Table.Buffer` sem justificativa?
* [ ] Listas ou metadados reutilizados deveriam utilizar `List.Buffer`?
* [ ] A avaliação lazy está sendo preservada?
* [ ] A deduplicação está sendo executada somente quando necessária?
* [ ] O problema de performance foi medido antes e depois da alteração?

---

# 23. Princípios para Novos Componentes

Novos componentes do framework devem seguir os seguintes princípios.

### Preferir operações nativas

Quando o Power Query possui uma operação nativa adequada, ela deve ser considerada antes de uma implementação manual.

### Processar por coluna

Quando a regra pertence a uma coluna, preferir processamento por coluna.

### Compilar antes de executar

Quando uma regra pode ser determinada a partir do Schema, preferir compilá-la antes do processamento dos dados.

### Bufferizar seletivamente

Bufferizar somente quando houver benefício identificável.

### Evitar múltiplas passagens

Quando várias operações podem ser combinadas sem prejudicar a arquitetura, considerar uma única passagem.

### Não otimizar prematuramente

Uma alteração de performance deve ser baseada em comportamento observado ou em uma necessidade arquitetural clara.

---

# 24. Performance e Escalabilidade

O framework utiliza Excel e Power Query como plataforma de execução.

Portanto, performance deve ser analisada dentro das características desse ambiente.

Os principais fatores são:

```text
Volume de dados
      │
      ▼
Número de colunas
      │
      ▼
Quantidade de transformações
      │
      ▼
Quantidade de validações
      │
      ▼
Número de passagens
      │
      ▼
Materialização / bufferização
      │
      ▼
Tempo de execução
```

Uma alteração que melhora uma etapa isoladamente pode aumentar o custo total se introduzir materialização ou processamento adicional em outra etapa.

---

# 25. Estratégia de Otimização

A estratégia recomendada é:

```text
1. Medir
   ↓
2. Identificar o gargalo
   ↓
3. Determinar a causa
   ↓
4. Alterar a implementação
   ↓
5. Medir novamente
   ↓
6. Comparar
```

Não se deve assumir que uma técnica conhecida de otimização, como `Table.Buffer`, necessariamente melhora uma consulta.

O benefício depende do padrão de avaliação, da origem dos dados e da reutilização do resultado.

---

# 26. Resumo das Estratégias

| Estratégia               | Aplicação                                        |
| ------------------------ | ------------------------------------------------ |
| `Table.TransformColumns` | Transformações por coluna                        |
| Pré-compilação           | Schema e operadores                              |
| `List.Buffer`            | Listas e estruturas reutilizadas                 |
| `Table.Buffer`           | Metadados ou resultados com benefício comprovado |
| QA em uma passagem       | Acúmulo de ocorrências                           |
| Avaliação lazy           | Evitar materialização antecipada                 |
| Processamento seletivo   | Somente colunas configuradas                     |
| `Table.Distinct`         | Deduplicação na NRM                              |
| Operações nativas        | Reduzir lógica manual                            |

---

# 27. Anti-Patterns

## `Table.Buffer` em todas as consultas

```powerquery
Table.Buffer(Tabela)
```

Não é uma estratégia geral de performance.

---

## Processamento linha a linha desnecessário

Evitar implementar uma transformação de coluna como uma sequência de reconstruções de registros quando uma operação nativa de tabela pode executar a mesma tarefa.

---

## Recompilar o mesmo pipeline

Evitar interpretar repetidamente o Schema durante o processamento dos dados quando o pipeline pode ser compilado e reutilizado.

---

## Criar várias passagens para o mesmo diagnóstico

Quando várias validações podem ser avaliadas durante a mesma passagem, evitar percorrer novamente a tabela apenas para consolidar os resultados.

---

## Otimizar sem medir

Uma mudança de implementação só deve ser considerada uma otimização depois de demonstrar redução de custo ou melhoria mensurável.

---

# 28. Relação com a Arquitetura

As otimizações devem preservar as responsabilidades definidas em `architecture.md`.

```text
SRC
 ↓
STG       → estrutura + tipos
 ↓
TRN       → tratamentos
 ↓
QA        → validações
 ↓
NRM       → normalização
 ↓
DIM/FATO  → modelagem
```

A performance deve ser obtida principalmente pela eficiência interna de cada etapa e pela preparação eficiente do pipeline.

Não se deve eliminar fronteiras arquiteturais apenas para reduzir o número de consultas.

---

# 29. Relação com o Pipeline

A principal otimização estrutural ocorre antes da execução dos dados:

```text
Schema
   │
   ▼
Pré-compilação
   │
   ▼
cfgPipeline
   │
   ▼
Execução
```

Isso reduz o trabalho necessário durante as etapas de processamento.

O detalhamento da compilação está documentado em:

`docs/pipeline.md`

---

# 30. Relação com o Schema

O Schema determina quais operações serão executadas.

Quanto mais precisamente o Schema representar o processamento necessário, menor a tendência de executar operações que não agregam valor.

Exemplo:

```text
Schema

CPF
 └── TRIM;DIGITS

Nome
 └── TRIM;PROPER

Observacao
 └── nenhum tratamento
```

O pipeline deve respeitar essa configuração e evitar aplicar operações genéricas desnecessárias.

---

# 31. Conclusão

A estratégia de performance do Excel Data Framework DBB não depende de uma única técnica.

Ela resulta da combinação de:

```text
                 PERFORMANCE
                      │
        ┌─────────────┼─────────────┐
        │             │             │
        ▼             ▼             ▼
  Processamento   Pré-compilação   Lazy
    por coluna      do pipeline   Evaluation
        │             │             │
        └─────────────┼─────────────┘
                      │
              ┌───────┴───────┐
              ▼               ▼
        Bufferização      Menos passagens
          seletiva
              │               │
              └───────┬───────┘
                      ▼
             Processamento
               eficiente
```

O princípio fundamental é:

> **Otimizar o caminho de execução sem comprometer as fronteiras arquiteturais do framework.**

As otimizações mais importantes do projeto são, portanto, estruturais: **pipeline pré-compilado, processamento por coluna, execução seletiva, QA em uma passagem, bufferização estratégica e preservação da avaliação lazy**.
