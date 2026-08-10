# Operators

## 1. Visão Geral

O Excel Data Framework DBB utiliza operadores para definir, de forma declarativa, os tratamentos e as validações aplicados às colunas das tabelas.

Os operadores são configurados no Schema e posteriormente compilados pelo pipeline antes da execução.

Conceitualmente:

```text
Schema
   │
   ▼
Operadores configurados
   │
   ▼
Compilação
   │
   ▼
Operadores executáveis
   │
   ▼
Execução no pipeline
```

Existem dois grupos principais de operadores:

* **Tratamentos** — modificam, normalizam ou convertem valores.
* **Validações** — verificam se os valores atendem às regras definidas.

---

# 2. Tratamentos

Os operadores de tratamento são utilizados na camada **TRN**.

Seu objetivo é transformar os valores de acordo com as regras definidas no Schema.

A execução é realizada por `fxTrnAplicar()`.

Os tratamentos são compilados a partir das definições existentes no Schema e combinados com os operadores padrão associados ao tipo da coluna.

---

## 2.1 Tratamentos básicos

| Operador         | Função                       |
| ---------------- | ---------------------------- |
| `TRIM`           | `fxTratamentoTrim`           |
| `UPPER`          | `fxTratamentoUpper`          |
| `LOWER`          | `fxTratamentoLower`          |
| `PROPER`         | `fxTratamentoProper`         |
| `CLEAN`          | `fxTratamentoClean`          |
| `EMPTYTONULL`    | `fxTratamentoEmptyToNull`    |
| `NULLTOEMPTY`    | `fxTratamentoNullToEmpty`    |
| `SINGLESPACE`    | `fxTratamentoSingleSpace`    |
| `DIGITS`         | `fxTratamentoDigits`         |
| `ALPHANUMERIC`   | `fxTratamentoAlphaNumeric`   |
| `ABS`            | `fxTratamentoAbs`            |
| `ROUND`          | `fxTratamentoRound`          |
| `NORMALIZEBASIC` | `fxTratamentoNormalizeBasic` |
| `NUMBER`         | `fxTratamentoNumber`         |
| `CPF`            | `fxTratamentoDigits`         |
| `CNPJ`           | `fxTratamentoDigits`         |
| `CEP`            | `fxTratamentoDigits`         |
| `PHONE`          | `fxTratamentoDigits`         |

Os operadores `CPF`, `CNPJ`, `CEP` e `PHONE` estão associados ao tratamento de dígitos (`fxTratamentoDigits`) na implementação documentada.

---

## 2.2 Tratamentos de texto avançados

| Operador            | Função                          |
| ------------------- | ------------------------------- |
| `REPLACE`           | `fxTratamentoReplace`           |
| `LEFT`              | `fxTratamentoLeft`              |
| `RIGHT`             | `fxTratamentoRight`             |
| `MID`               | `fxTratamentoMid`               |
| `BEFORE`            | `fxTratamentoBefore`            |
| `AFTER`             | `fxTratamentoAfter`             |
| `ADDPREFIX`         | `fxTratamentoAddPrefix`         |
| `ADDSUFFIX`         | `fxTratamentoAddSuffix`         |
| `REMOVEPREFIX`      | `fxTratamentoRemovePrefix`      |
| `REMOVESUFFIX`      | `fxTratamentoRemoveSuffix`      |
| `PADLEFT`           | `fxTratamentoPadLeft`           |
| `PADRIGHT`          | `fxTratamentoPadRight`          |
| `REMOVECHARS`       | `fxTratamentoRemoveChars`       |
| `KEEPCHARS`         | `fxTratamentoKeepChars`         |
| `REMOVEACCENTS`     | `fxTratamentoRemoveAccents`     |
| `REMOVEPUNCTUATION` | `fxTratamentoRemovePunctuation` |
| `KEEPTEXT`          | `fxTratamentoKeepText`          |

A lista acima corresponde ao catálogo documentado atualmente no projeto.

---

# 3. Sintaxe dos Tratamentos

Os operadores podem ser utilizados isoladamente ou encadeados.

Exemplo documentado no Schema:

```text
TRIM;UPPER;DIGITS
```

Outro exemplo:

```text
TRIM;PROPER
```

Operadores que recebem parâmetros utilizam a sintaxe:

```text
OPERADOR(parâmetro1,parâmetro2)
```

Exemplos documentados no funcionamento do compilador incluem:

```text
ADDPREFIX(P)
PADLEFT(3,0,P)
```

O compilador identifica o código do operador e extrai os parâmetros antes de resolver o operador em `cfgOperadores`.

---

# 4. Tratamentos Padrão

O framework permite associar operadores padrão ao tipo da coluna.

Conceitualmente:

```text
Tipo da coluna
      │
      ▼
fxOperadoresPadrao(Tipo)
      │
      ├── Tratamentos padrão
      └── Validações padrão
```

Os operadores padrão podem ser combinados com os operadores explicitamente definidos no Schema.

A compilação de uma coluna considera:

```text
Tratamentos padrão
        +
Tratamentos definidos no Schema
        ↓
TratamentosPorColuna
```

Essa composição é realizada durante a compilação do pipeline.

---

# 5. Validações

Os operadores de validação são utilizados na camada **QA**.

Seu objetivo é verificar se os valores atendem às regras de qualidade definidas no Schema.

A execução é realizada por `fxQaValidar()`.

Diferentemente dos tratamentos, as validações não têm como objetivo modificar o valor.

Conceitualmente:

```text
TRN
 │
 ▼
Valor tratado
 │
 ▼
QA
 │
 ├── Validação
 ├── Ocorrência
 └── Status
```

---

## 5.1 Validadores disponíveis

| Operador   | Função                |
| ---------- | --------------------- |
| `REQUIRED` | `fxValidacaoREQUIRED` |
| `LIST`     | `fxValidacaoList`     |
| `SIZE`     | `fxValidacaoSize`     |
| `MIN`      | `fxValidacaoMin`      |
| `MAX`      | `fxValidacaoMax`      |
| `INTERVAL` | `fxValidacaoInterval` |
| `EMAIL`    | `fxValidacaoEmail`    |
| `URL`      | `fxValidacaoURL`      |
| `CEPVAL`   | `fxValidacaoCEP`      |
| `CPFVAL`   | `fxValidacaoCPF`      |
| `CNPJVAL`  | `fxValidacaoCNPJ`     |
| `PHONEVAL` | `fxValidacaoPhone`    |

Essa é a relação de validadores documentada atualmente no projeto.

---

# 6. REQUIRED

`REQUIRED` verifica a obrigatoriedade de um valor.

A obrigatoriedade também pode ser declarada diretamente no Schema através do campo `Obrigatório`.

Quando:

```text
Obrigatório = true
```

o pipeline pode incluir implicitamente o operador `REQUIRED`.

Assim:

```text
Obrigatório = true
        │
        ▼
REQUIRED implícito
```

Essa regra faz parte da compilação das validações por coluna.

---

# 7. Validações Padrão

Assim como ocorre com os tratamentos, o framework pode associar validações padrão ao tipo da coluna.

A composição ocorre conceitualmente da seguinte forma:

```text
REQUIRED implícito
        +
Validações padrão do tipo
        +
Validações definidas no Schema
        │
        ▼
ValidaçõesPorColuna
```

Essa estrutura é compilada antes da execução do QA.

---

# 8. Composição de Operadores

Os operadores de uma coluna podem ser combinados para formar uma sequência de processamento.

Exemplo:

```text
CPF

TRIM
  ↓
DIGITS
  ↓
CPFVAL
```

No Schema:

```text
Tratamentos: TRIM;DIGITS
Validações: CPFVAL
```

A distinção entre os dois grupos é importante:

```text
Tratamento
    ↓
modifica o valor

Validação
    ↓
avalia o valor
```

Essa separação permite que uma validação seja executada sobre o valor já tratado.

---

# 9. Compilação de Operadores

O framework não precisa interpretar toda a definição textual do operador durante cada processamento.

A configuração é compilada previamente.

O processo pode ser representado como:

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

O registro compilado contém informações como:

* código do operador;
* parâmetros extraídos;
* referência à função executável.

O README atual descreve `fxPipelineCompilarOperadores` como o componente responsável por resolver cada operador do Schema em uma estrutura de execução.

---

# 10. Compilação por Coluna

A compilação também ocorre no contexto de cada coluna.

Para tratamentos:

```text
Tipo
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
Tipo
 │
 ├── REQUIRED implícito
 ├── Validações padrão
 └── Validações do Schema
          │
          ▼
ValidaçõesPorColuna
```

O resultado é armazenado no pipeline compilado.

---

# 11. Exemplo de Schema

Um Schema pode definir:

```text
Tabela      | Coluna         | Tipo | Obrigatório | Tratamentos      | Validações
tbClientes  | CPF            | TEXT | SIM         | TRIM;DIGITS      | CPFVAL
tbClientes  | Nome           | TEXT | SIM         | TRIM;PROPER      | SIZE(100)
tbClientes  | DataNascimento | DATE | NÃO         |                  |
tbClientes  | Cidade         | TEXT | NÃO         | TRIM;PROPER      |
tbClientes  | Estado         | TEXT | NÃO         | TRIM;UPPER       | LIST(RS,SP,SC)
```

Esse exemplo demonstra a separação entre:

* tipo;
* obrigatoriedade;
* tratamentos;
* validações.

A estrutura está alinhada ao Schema documentado no projeto.

---

# 12. Pipeline Compilado

Depois da compilação, uma tabela pode possuir uma estrutura conceitual como:

```text
cfgPipeline[tbClientes]
```

com:

```text
Ordem
TiposPorColuna
TratamentosPorColuna
ValidaçõesPorColuna
ChavesNegocio
```

Exemplo:

```powerquery
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

Essa estrutura representa o resultado da compilação do Schema em uma configuração adequada para execução.

---

# 13. Tratamento x Validação

A diferença fundamental pode ser resumida assim:

| Característica   | Tratamento          | Validação         |
| ---------------- | ------------------- | ----------------- |
| Camada principal | TRN                 | QA                |
| Objetivo         | Transformar         | Avaliar           |
| Modifica valor   | Sim                 | Não               |
| Gera diagnóstico | Não necessariamente | Sim               |
| Exemplo          | `TRIM`              | `REQUIRED`        |
| Exemplo          | `DIGITS`            | `CPFVAL`          |
| Resultado        | Valor tratado       | Ocorrência/status |

Essa separação evita misturar transformação de dados com regras de qualidade.

---

# 14. Operadores e Tipos

O comportamento dos operadores pode depender do tipo da coluna.

O pipeline utiliza o tipo definido no Schema para determinar as estruturas de processamento associadas à coluna.

Conceitualmente:

```text
Tipo
 │
 ▼
Operadores padrão
 │
 ├── Tratamentos
 └── Validações
```

Isso permite que determinadas regras sejam aplicadas automaticamente conforme o tipo configurado.

A relação completa entre tipos e operadores padrão deve ser documentada junto à configuração de `fxOperadoresPadrao`.

---

# 15. Registro de Ocorrências

As validações produzem informações de qualidade que são armazenadas na camada QA.

O resultado inclui:

```text
_QA_Status
_QA_Ocorrencias
```

As ocorrências podem conter informações como:

```text
Código
Severidade
Coluna
Tipo
Descrição
```

Essas informações permitem identificar não apenas que um registro falhou, mas também qual regra foi responsável pela ocorrência.

---

# 16. Severidade

A classificação das ocorrências utiliza parâmetros de severidade.

O projeto possui:

```text
cfgParametrosSeveridades
```

responsável por controlar classificações como:

```text
AVISO
ERRO
```

A severidade influencia a classificação do registro no QA.

A configuração detalhada das severidades deve ser tratada na documentação específica de parâmetros do framework.

---

# 17. Catálogo Rápido

## Tratamentos

```text
TRIM
UPPER
LOWER
PROPER
CLEAN
EMPTYTONULL
NULLTOEMPTY
SINGLESPACE
DIGITS
ALPHANUMERIC
ABS
ROUND
NORMALIZEBASIC
NUMBER
CPF
CNPJ
CEP
PHONE

REPLACE
LEFT
RIGHT
MID
BEFORE
AFTER
ADDPREFIX
ADDSUFFIX
REMOVEPREFIX
REMOVESUFFIX
PADLEFT
PADRIGHT
REMOVECHARS
KEEPCHARS
REMOVEACCENTS
REMOVEPUNCTUATION
KEEPTEXT
```

## Validações

```text
REQUIRED
LIST
SIZE
MIN
MAX
INTERVAL
EMAIL
URL
CEPVAL
CPFVAL
CNPJVAL
PHONEVAL
```

---

# 18. Referência das Funções

## Tratamentos

Os operadores são associados às seguintes funções:

```text
TRIM             → fxTratamentoTrim
UPPER             → fxTratamentoUpper
LOWER             → fxTratamentoLower
PROPER            → fxTratamentoProper
CLEAN             → fxTratamentoClean
EMPTYTONULL       → fxTratamentoEmptyToNull
NULLTOEMPTY       → fxTratamentoNullToEmpty
SINGLESPACE       → fxTratamentoSingleSpace
DIGITS            → fxTratamentoDigits
ALPHANUMERIC      → fxTratamentoAlphaNumeric
ABS               → fxTratamentoAbs
ROUND             → fxTratamentoRound
NORMALIZEBASIC    → fxTratamentoNormalizeBasic
NUMBER            → fxTratamentoNumber
CPF               → fxTratamentoDigits
CNPJ              → fxTratamentoDigits
CEP               → fxTratamentoDigits
PHONE             → fxTratamentoDigits
REPLACE            → fxTratamentoReplace
LEFT               → fxTratamentoLeft
RIGHT              → fxTratamentoRight
MID                → fxTratamentoMid
BEFORE             → fxTratamentoBefore
AFTER              → fxTratamentoAfter
ADDPREFIX          → fxTratamentoAddPrefix
ADDSUFFIX          → fxTratamentoAddSuffix
REMOVEPREFIX       → fxTratamentoRemovePrefix
REMOVESUFFIX       → fxTratamentoRemoveSuffix
PADLEFT            → fxTratamentoPadLeft
PADRIGHT           → fxTratamentoPadRight
REMOVECHARS        → fxTratamentoRemoveChars
KEEPCHARS          → fxTratamentoKeepChars
REMOVEACCENTS      → fxTratamentoRemoveAccents
REMOVEPUNCTUATION  → fxTratamentoRemovePunctuation
KEEPTEXT            → fxTratamentoKeepText
```

## Validações

```text
REQUIRED    → fxValidacaoREQUIRED
LIST        → fxValidacaoList
SIZE        → fxValidacaoSize
MIN         → fxValidacaoMin
MAX         → fxValidacaoMax
INTERVAL    → fxValidacaoInterval
EMAIL       → fxValidacaoEmail
URL         → fxValidacaoURL
CEPVAL      → fxValidacaoCEP
CPFVAL      → fxValidacaoCPF
CNPJVAL     → fxValidacaoCNPJ
PHONEVAL    → fxValidacaoPhone
```

O catálogo acima corresponde ao conteúdo atualmente documentado no projeto.

---

# 19. Adicionando um Novo Operador

A inclusão de um novo operador deve respeitar a arquitetura de compilação existente.

Conceitualmente:

```text
1. Criar a função
        ↓
2. Registrar o operador
        ↓
3. Definir parâmetros, quando aplicável
        ↓
4. Permitir resolução pelo compilador
        ↓
5. Utilizar no Schema
        ↓
6. Testar a execução
```

O operador deve ser registrado na estrutura utilizada pelo framework para resolução dos operadores.

No estado atual da documentação, os detalhes internos de cadastro e parametrização dos operadores são definidos pelas estruturas de configuração do projeto.

---

# 20. Boas Práticas

### Utilize operadores para regras reutilizáveis

Se uma transformação pode ser aplicada a várias tabelas, ela deve preferencialmente ser implementada como operador reutilizável.

### Evite lógica específica no pipeline

O pipeline deve executar as regras configuradas, e não conter regras específicas para cada tabela.

### Separe tratamento de validação

Quando uma operação altera o valor, ela deve ser tratada como transformação.

Quando uma operação apenas verifica uma condição, ela deve ser tratada como validação.

### Prefira composição

Em vez de criar operadores excessivamente específicos, utilize a composição de operadores existentes quando isso mantiver o comportamento claro.

### Mantenha o Schema declarativo

As configurações devem expressar **o que deve acontecer**, enquanto o framework define **como executar**.

---

# 21. Documentação Relacionada

* [Getting Started](getting-started.md) — primeiro pipeline.
* [Architecture](architecture.md) — arquitetura e responsabilidades das camadas.
* [Pipeline](pipeline.md) — compilação e execução dos operadores.
* [Schema](schema.md) — configuração dos operadores no Schema.
* [Performance](performance.md) — estratégias de otimização.
* [Examples](examples.md) — utilização em cenários completos.
* [Troubleshooting](troubleshooting.md) — diagnóstico de problemas.

---

# 22. Resumo

Os operadores são a unidade básica de comportamento configurável do framework.

```text
Tratamentos
    ↓
Transformam os valores

Validações
    ↓
Avaliam os valores

Schema
    ↓
Define os operadores

Pipeline
    ↓
Compila os operadores

Engine
    ↓
Executa os operadores
```

Essa arquitetura permite que novas regras de tratamento e validação sejam incorporadas ao framework sem duplicar a lógica do pipeline para cada tabela.
