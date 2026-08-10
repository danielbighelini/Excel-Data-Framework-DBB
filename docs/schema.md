# Schema

## 1. Visão Geral

O Schema é a principal estrutura de configuração do Excel Data Framework DBB.

Ele define, de forma declarativa, **como cada tabela deve ser processada pelo framework**.

Em vez de implementar regras específicas dentro de cada consulta, o comportamento é descrito no Schema e posteriormente compilado pelo pipeline.

O conceito central é:

```text
Schema
   │
   ├── Estrutura
   ├── Tipos
   ├── Tratamentos
   ├── Validações
   ├── Obrigatoriedade
   ├── Ordem
   └── Chaves de negócio
          │
          ▼
    Pipeline compilado
          │
          ▼
       Execução
```

Essa abordagem permite reutilizar o mesmo mecanismo de processamento para diferentes tabelas.

---

# 2. Objetivo do Schema

O Schema tem quatro objetivos principais:

1. **Descrever a estrutura esperada dos dados.**
2. **Definir como os valores devem ser tratados.**
3. **Definir como os valores devem ser validados.**
4. **Fornecer metadados necessários às etapas posteriores do pipeline.**

O Schema não executa diretamente essas operações.

Ele fornece as definições que serão utilizadas pelo framework.

```text
Schema
  │
  │ define
  ▼
Regras de processamento
  │
  │ compiladas por
  ▼
Pipeline
  │
  │ executado por
  ▼
Camadas do framework
```

---

# 3. Estrutura do Schema

A estrutura documentada utiliza as seguintes informações principais:

| Campo           | Finalidade                                         |
| --------------- | -------------------------------------------------- |
| `Tabela`        | Identifica a tabela à qual a definição pertence    |
| `Coluna`        | Identifica a coluna configurada                    |
| `Tipo`          | Define o tipo esperado da coluna                   |
| `Obrigatório`   | Define se o preenchimento da coluna é obrigatório  |
| `Tratamentos`   | Define os operadores de transformação              |
| `Validações`    | Define os operadores de validação                  |
| `Ordem`         | Define a posição da coluna na estrutura resultante |
| `ChavesNegocio` | Define as colunas utilizadas como chave de negócio |

A implementação atual também utiliza essas informações durante a compilação do pipeline.

---

# 4. Exemplo de Schema

Um exemplo de definição é:

```text
Tabela      | Coluna         | Tipo | Obrigatório | Tratamentos      | Validações
tbClientes  | CPF            | TEXT | SIM         | TRIM;DIGITS      | CPFVAL
tbClientes  | Nome           | TEXT | SIM         | TRIM;PROPER      | SIZE(100)
tbClientes  | DataNascimento | DATE | NÃO         |                  |
tbClientes  | Cidade         | TEXT | NÃO         | TRIM;PROPER      |
tbClientes  | Estado         | TEXT | NÃO         | TRIM;UPPER       | LIST(RS,SP,SC)
```

Esse exemplo representa uma tabela chamada `tbClientes` com cinco colunas configuradas.

Cada linha representa uma definição de coluna.

---

# 5. Tabela

O campo `Tabela` identifica a estrutura à qual a definição pertence.

Exemplo:

```text
Tabela = tbClientes
```

Todas as definições pertencentes a essa tabela são agrupadas durante a preparação do Schema.

Exemplo:

```text
tbClientes | CPF
tbClientes | Nome
tbClientes | DataNascimento
tbClientes | Cidade
tbClientes | Estado
```

O agrupamento permite que o framework gere um pipeline específico para `tbClientes`.

Da mesma forma, outras tabelas podem possuir configurações independentes:

```text
tbClientes
tbProdutos
tbVendas
```

---

# 6. Coluna

O campo `Coluna` identifica a coluna que está sendo configurada.

Exemplo:

```text
Coluna = CPF
```

A combinação:

```text
Tabela + Coluna
```

identifica a definição de uma coluna específica dentro do Schema.

Exemplo:

```text
tbClientes + CPF
```

Essa definição pode determinar:

* tipo;
* obrigatoriedade;
* tratamentos;
* validações.

---

# 7. Tipo

O campo `Tipo` define o tipo esperado para a coluna.

Exemplos utilizados no Schema:

```text
TEXT
DATE
```

Outros tipos podem ser disponibilizados pelo mecanismo de configuração de tipos do framework.

O Schema não precisa armazenar diretamente o tipo M utilizado internamente.

O framework utiliza sua configuração de tipos para realizar o mapeamento para os tipos do Power Query.

Conceitualmente:

```text
Tipo do Schema
      │
      ▼
cfgTiposDados
      │
      ▼
Tipo M
```

A aplicação inicial desses tipos ocorre na camada STG.

---

# 8. Obrigatório

O campo `Obrigatório` determina se uma coluna deve possuir um valor.

Exemplo:

```text
Obrigatório = SIM
```

ou:

```text
Obrigatório = NÃO
```

Quando uma coluna é definida como obrigatória, a compilação das validações pode incluir implicitamente o operador `REQUIRED`.

Conceitualmente:

```text
Obrigatório = SIM
        │
        ▼
REQUIRED
        │
        ▼
Validação QA
```

Essa regra faz parte da compilação das validações por coluna.

---

# 9. Tratamentos

O campo `Tratamentos` define os operadores que devem ser executados na camada TRN.

Os operadores são separados por `;`.

Exemplo:

```text
TRIM;UPPER
```

Outro exemplo:

```text
TRIM;DIGITS
```

Isso representa uma sequência de processamento:

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
Resultado
```

A definição dos operadores disponíveis está documentada em:

`operators.md`

---

# 10. Ordem dos Tratamentos

Quando mais de um tratamento é definido, a sequência declarada representa a ordem em que os operadores são aplicados.

Por exemplo:

```text
TRIM;UPPER
```

é conceitualmente diferente de:

```text
UPPER;TRIM
```

O pipeline compila essa sequência antes da execução.

```text
Schema
  │
  ▼
TRIM;UPPER
  │
  ▼
Pipeline compilado
  │
  ▼
Execução
```

A ordem dos operadores deve, portanto, ser considerada parte da regra de transformação.

---

# 11. Tratamentos com Parâmetros

Operadores parametrizados utilizam a seguinte estrutura:

```text
OPERADOR(param1,param2,...)
```

Exemplos documentados no projeto incluem:

```text
ADDPREFIX(P)
PADLEFT(3,0,P)
```

Durante a compilação, o framework identifica o código do operador e extrai seus parâmetros antes de resolver a função executável.

Exemplo conceitual:

```text
PADLEFT(3,0,P)
      │
      ├── Operador: PADLEFT
      ├── Parâmetro: 3
      ├── Parâmetro: 0
      └── Parâmetro: P
```

A sintaxe e os parâmetros suportados por cada operador devem ser consultados em `operators.md`.

---

# 12. Validações

O campo `Validações` define as regras utilizadas pela camada QA.

Exemplos:

```text
CPFVAL
```

ou:

```text
LIST(RS,SP,SC)
```

ou:

```text
SIZE(100)
```

Os validadores não têm como finalidade modificar o valor.

Eles avaliam o valor e podem gerar ocorrências de qualidade.

```text
Valor
  │
  ▼
Validação
  │
  ├── OK
  └── Ocorrência
```

---

# 13. Validações com Parâmetros

Assim como os tratamentos, as validações podem receber parâmetros.

Exemplos:

```text
SIZE(100)
```

```text
LIST(RS,SP,SC)
```

```text
INTERVAL(...)
```

Durante a compilação, os parâmetros são extraídos da definição textual e incorporados à estrutura executável.

A relação completa dos validadores está documentada em `operators.md`.

---

# 14. Validação Obrigatória

O campo `Obrigatório` possui uma característica importante.

A obrigatoriedade não precisa necessariamente ser declarada como um operador explícito em `Validações`.

Por exemplo:

```text
Tabela      | Coluna | Obrigatório | Validações
tbClientes  | CPF    | SIM         | CPFVAL
```

A configuração pode resultar conceitualmente em:

```text
ValidaçõesPorColuna[CPF]

REQUIRED
CPFVAL
```

Ou seja:

```text
Obrigatório = SIM
        +
CPFVAL
        ↓
REQUIRED + CPFVAL
```

O `REQUIRED` implícito é definido durante a compilação das validações por coluna.

---

# 15. Validações Padrão

O framework também suporta validações padrão associadas ao tipo da coluna.

A compilação considera conceitualmente:

```text
REQUIRED implícito
        +
Validações padrão do tipo
        +
Validações do Schema
        │
        ▼
ValidaçõesPorColuna
```

Isso permite estabelecer comportamentos comuns para determinados tipos sem repetir todas as configurações em cada linha do Schema.

O mecanismo é baseado em `fxOperadoresPadrao(Tipo)`.

---

# 16. Chaves de Negócio

As chaves de negócio identificam as colunas utilizadas pela etapa de normalização.

Exemplo:

```text
ChavesNegocio = {"CPF"}
```

Para uma chave composta:

```text
ChavesNegocio =
{
    "CodigoCliente",
    "CodigoProduto"
}
```

As chaves de negócio são posteriormente disponibilizadas no pipeline compilado e utilizadas pela camada NRM.

```text
Schema
  │
  ▼
ChavesNegocio
  │
  ▼
cfgPipeline
  │
  ▼
NRM
```

A documentação atual identifica `ChavesNegocio` como um dos campos do pipeline compilado.

---

# 17. Ordem das Colunas

O Schema também pode determinar a ordem das colunas da estrutura resultante.

Conceitualmente:

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

Essa informação é compilada e disponibilizada no pipeline.

A camada STG utiliza essa configuração para manter uma estrutura de saída previsível.

---

# 18. Colunas sem Tratamentos

Uma coluna não precisa possuir tratamentos.

Exemplo:

```text
Tabela      | Coluna         | Tipo | Obrigatório | Tratamentos | Validações
tbClientes  | DataNascimento | DATE | NÃO         |             |
```

Nesse caso, a coluna ainda participa do Schema e pode receber:

* tipagem;
* ordenação;
* validações padrão, quando aplicáveis.

A ausência de tratamentos não significa que a coluna seja ignorada pelo framework.

---

# 19. Colunas sem Validações

Da mesma forma, uma coluna pode não possuir validações explícitas.

Exemplo:

```text
Tabela      | Coluna  | Tipo | Obrigatório | Tratamentos | Validações
tbClientes  | Cidade  | TEXT | NÃO         | TRIM;PROPER |
```

A coluna ainda pode participar de:

* tipagem;
* tratamentos;
* ordenação;
* regras padrão associadas ao tipo.

---

# 20. Colunas Obrigatórias sem Validação Explícita

Uma coluna pode ser obrigatória mesmo que o campo `Validações` esteja vazio.

Exemplo:

```text
Tabela      | Coluna | Tipo | Obrigatório | Tratamentos | Validações
tbClientes  | Nome   | TEXT | SIM         | TRIM;PROPER |
```

Nesse cenário, a obrigatoriedade fornece uma regra de validação implícita:

```text
Obrigatório = SIM
        │
        ▼
REQUIRED
```

Esse comportamento evita que a configuração de obrigatoriedade dependa da declaração manual de `REQUIRED`.

---

# 21. Composição da Configuração

Uma coluna pode combinar diferentes tipos de configuração.

Exemplo:

```text
CPF

Tipo:
TEXT

Obrigatório:
SIM

Tratamentos:
TRIM;DIGITS

Validações:
CPFVAL

Chave de negócio:
SIM
```

Conceitualmente:

```text
                  CPF
                   │
       ┌───────────┼───────────┐
       │           │           │
       ▼           ▼           ▼
     Tipo      Tratamentos  Validações
       │           │           │
       ▼           ▼           ▼
     TEXT      TRIM;DIGITS   CPFVAL
                               │
                         REQUIRED implícito
```

---

# 22. Schema e Pipeline

O Schema é uma configuração declarativa.

O pipeline é a estrutura compilada utilizada durante a execução.

```text
Schema
   │
   ▼
Compilação
   │
   ▼
cfgPipeline
```

O resultado pode conter:

```text
Ordem
TiposPorColuna
TratamentosPorColuna
ValidaçõesPorColuna
ChavesNegocio
```

Exemplo conceitual:

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

A estrutura compilada é descrita em maior detalhe em `pipeline.md`.

---

# 23. Schema por Tabela

Cada tabela pode possuir uma configuração independente.

Exemplo:

```text
Schema
│
├── tbClientes
│   ├── CPF
│   ├── Nome
│   ├── Cidade
│   └── Estado
│
├── tbProdutos
│   ├── Codigo
│   ├── Descricao
│   └── Preco
│
└── tbVendas
    ├── Data
    ├── CPF
    ├── CodigoProduto
    └── Valor
```

Cada tabela possui seu próprio conjunto de:

* tipos;
* tratamentos;
* validações;
* ordem;
* chaves.

Isso permite que o mesmo mecanismo de execução seja reutilizado para diferentes estruturas.

---

# 24. Exemplo Completo

Considere a seguinte configuração:

```text
Tabela      | Coluna         | Tipo | Obrigatório | Tratamentos      | Validações
tbClientes  | CPF            | TEXT | SIM         | TRIM;DIGITS      | CPFVAL
tbClientes  | Nome           | TEXT | SIM         | TRIM;PROPER      | SIZE(100)
tbClientes  | DataNascimento | DATE | NÃO         |                  |
tbClientes  | Cidade         | TEXT | NÃO         | TRIM;PROPER      |
tbClientes  | Estado         | TEXT | NÃO         | TRIM;UPPER       | LIST(RS,SP,SC)
```

O comportamento resultante pode ser entendido como:

### CPF

```text
Tipo:
TEXT

Tratamento:
TRIM → DIGITS

Validação:
REQUIRED → CPFVAL
```

### Nome

```text
Tipo:
TEXT

Tratamento:
TRIM → PROPER

Validações:
REQUIRED → SIZE(100)
```

### DataNascimento

```text
Tipo:
DATE

Sem tratamento ou validação explicitamente declarados.
```

### Cidade

```text
Tipo:
TEXT

Tratamento:
TRIM → PROPER
```

### Estado

```text
Tipo:
TEXT

Tratamento:
TRIM → UPPER

Validação:
LIST(RS,SP,SC)
```

---

# 25. Fluxo de uma Definição

Uma definição individual percorre conceitualmente o seguinte fluxo:

```text
Linha do Schema
       │
       ▼
Interpretação
       │
       ├── Tipo
       ├── Obrigatório
       ├── Tratamentos
       ├── Validações
       └── Ordem
       │
       ▼
Compilação
       │
       ▼
Configuração da coluna
       │
       ▼
Pipeline
       │
       ▼
Execução
```

---

# 26. Regras de Configuração

As configurações do Schema devem seguir algumas regras conceituais.

### 26.1 Uma definição deve representar uma coluna

Cada definição deve identificar claramente:

```text
Tabela + Coluna
```

### 26.2 Tratamentos devem representar transformação

Use `Tratamentos` para operações que modificam ou padronizam o valor.

### 26.3 Validações devem representar avaliação

Use `Validações` para operações que verificam uma condição.

### 26.4 Obrigatoriedade deve representar requisito estrutural

Use `Obrigatório` para declarar que uma coluna deve possuir valor.

### 26.5 Chaves devem representar identidade de negócio

Use `ChavesNegocio` para identificar como os registros devem ser distinguidos durante a normalização.

---

# 27. Schema e Separação de Responsabilidades

O Schema não deve conter código de execução.

Ele deve descrever regras.

```text
CORRETO

Schema
   ↓
TRIM;UPPER
   ↓
Pipeline
   ↓
Engine executa


EVITAR

Schema
   ↓
Código M específico da tabela
   ↓
Execução
```

O objetivo da configuração declarativa é manter o comportamento separado da implementação.

---

# 28. Extensão do Schema

Quando uma nova capacidade for adicionada ao framework, deve-se avaliar se ela realmente precisa ser representada no Schema.

Uma nova propriedade deve possuir:

1. significado claro;
2. responsabilidade definida;
3. regra de compilação;
4. consumidor no pipeline;
5. comportamento documentado.

O Schema não deve acumular propriedades apenas porque elas podem ser tecnicamente armazenadas.

Cada campo deve representar uma decisão de configuração necessária ao processamento.

---

# 29. Relação com as Configurações do Framework

O Schema trabalha em conjunto com outras estruturas de configuração.

Entre as estruturas documentadas estão:

```text
srcSchema
stgSchema
cfgSchema
cfgPipeline
cfgTiposDados
cfgTiposBooleanos
cfgParametrosSeveridades
```

Conceitualmente:

```text
Fonte de configuração
        │
        ▼
Schema
        │
        ├──────────────┐
        │              │
        ▼              ▼
Tipos / Parâmetros   Regras
        │              │
        └──────┬───────┘
               ▼
          Compilação
               │
               ▼
          cfgPipeline
```

As responsabilidades específicas dessas estruturas pertencem à arquitetura interna do framework.

---

# 30. Schema e Qualidade

O Schema define as regras que serão utilizadas pela camada QA.

Por exemplo:

```text
CPF
  └── CPFVAL

Estado
  └── LIST(RS,SP,SC)

Nome
  └── SIZE(100)
```

Essas regras são compiladas em `ValidaçõesPorColuna`.

Durante a execução:

```text
Schema
  ↓
ValidaçõesPorColuna
  ↓
fxQaValidar
  ↓
_QA_Ocorrencias
_QA_Status
```

Assim, as regras de qualidade são centralizadas na configuração em vez de ficarem espalhadas pelas consultas.

---

# 31. Schema e Normalização

As chaves de negócio definidas no Schema são utilizadas pela camada NRM.

Exemplo:

```text
Tabela: tbClientes

ChavesNegocio:
CPF
```

Fluxo:

```text
Schema
  │
  ▼
ChavesNegocio
  │
  ▼
cfgPipeline
  │
  ▼
fxNrmAplicar
  │
  ▼
Deduplicação
```

A definição da chave deve representar a identidade lógica do registro dentro do domínio de negócio.

---

# 32. Boas Práticas

### Centralize regras

Se uma regra é configurável e reutilizável, prefira representá-la no Schema em vez de duplicá-la em várias consultas.

### Evite configurações redundantes

Não configure explicitamente uma regra quando ela já é fornecida de forma padronizada pelo framework, salvo quando a intenção for tornar o comportamento explícito.

### Mantenha tratamentos simples

Prefira combinar operadores existentes quando isso mantiver a configuração clara.

### Mantenha validações independentes

Uma validação deve representar uma regra específica e identificável.

### Use nomes consistentes

Os nomes das tabelas e colunas no Schema devem corresponder às estruturas efetivamente processadas.

### Trate o Schema como contrato

Alterações no Schema podem alterar o comportamento do pipeline. Portanto, mudanças estruturais devem ser tratadas como alterações de configuração do sistema, e não apenas como edição de uma tabela auxiliar.

---

# 33. Diagnóstico de Problemas no Schema

Quando um pipeline não se comporta conforme esperado, as primeiras verificações devem ser:

1. A tabela está corretamente identificada?
2. A coluna está corretamente identificada?
3. O tipo está configurado corretamente?
4. O operador está cadastrado?
5. A sintaxe do operador está correta?
6. Os parâmetros estão corretos?
7. A obrigatoriedade está configurada conforme esperado?
8. A validação pertence realmente àquela coluna?
9. A chave de negócio está correta?
10. O pipeline foi compilado com a configuração esperada?

Problemas de compilação relacionados a operadores devem ser investigados também nas estruturas de configuração de tratamentos e validações.

---

# 34. Documentação Relacionada

* [Getting Started](getting-started.md) — configuração inicial do framework.
* [Architecture](architecture.md) — arquitetura e responsabilidades das camadas.
* [Pipeline](pipeline.md) — compilação e execução do Schema.
* [Operators](operators.md) — tratamentos e validações disponíveis.
* [Performance](performance.md) — estratégias de otimização.
* [Examples](examples.md) — exemplos completos.
* [Troubleshooting](troubleshooting.md) — diagnóstico de problemas.

---

# 35. Resumo

O Schema é o **contrato declarativo de processamento** do Excel Data Framework DBB.

Ele define:

```text
┌──────────────────────────────┐
│            SCHEMA            │
├──────────────────────────────┤
│ Tabela                       │
│ Coluna                       │
│ Tipo                         │
│ Obrigatório                  │
│ Tratamentos                  │
│ Validações                   │
│ Ordem                        │
│ Chaves de negócio            │
└──────────────┬───────────────┘
               │
               ▼
       Pipeline compilado
               │
               ▼
     Execução do framework
```

O princípio fundamental é:

> **O Schema descreve o comportamento esperado; o pipeline transforma essa definição em uma estrutura executável; as camadas do framework aplicam as regras aos dados.**

Essa separação permite configurar diferentes tabelas utilizando o mesmo mecanismo de processamento, mantendo as regras específicas de cada estrutura centralizadas e declarativas.
