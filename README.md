# Excel Data Framework DBB

> Framework para Power Query que transforma planilhas do Excel em pipelines de dados estruturados, reutilizáveis e orientados por metadados.

**Excel Data Framework DBB** é um framework desenvolvido para padronizar processos de **extração, transformação, validação, normalização e modelagem de dados** utilizando Power Query.

Diferente de soluções baseadas em consultas isoladas, o framework utiliza um **Schema declarativo** para definir como cada tabela deve ser processada, reduzindo código repetitivo e aumentando a padronização entre projetos.

> **Commercial License · Source Available**

O código-fonte Power Query (`.m`) e a documentação são disponibilizados publicamente para inspeção e avaliação técnica. O **projeto Excel completo (`.xlsx`) e demais componentes da distribuição comercial são fornecidos exclusivamente aos usuários licenciados**.

---

# O problema

Projetos desenvolvidos diretamente no Power Query normalmente apresentam alguns problemas recorrentes:

* regras de transformação espalhadas entre consultas;
* código duplicado;
* validações inconsistentes;
* baixa reutilização;
* dificuldade de manutenção;
* ausência de padronização entre projetos;
* pipelines difíceis de entender;
* forte acoplamento entre regra de negócio e implementação.

À medida que o volume e a complexidade dos dados aumentam, esses problemas tornam-se cada vez mais difíceis de administrar.

O Excel Data Framework DBB foi criado para estruturar esse cenário por meio de uma arquitetura baseada em **camadas, metadados e configuração declarativa**.

---

# Principais características

* Arquitetura em camadas (`SRC → STG → TRN → QA → NRM → DIM/FATO`)
* Processamento orientado por Schema
* Tratamentos declarativos
* Validações declarativas
* Pipeline compilado
* Separação entre transformação e validação
* Registro estruturado de ocorrências
* Chaves de negócio e normalização
* Suporte à modelagem dimensional
* Componentes reutilizáveis e desacoplados
* Processamento por coluna
* Uso de operações nativas do Power Query
* Bufferização seletiva de metadados e estruturas reutilizadas
* Arquitetura orientada à manutenção e reutilização

---

# Arquitetura

```text
          EXTRAÇÃO (SRC)

     srcClientes
     srcProdutos
     srcVendas
            │
            ▼

        STAGING (STG)

 Preparação estrutural
 Aplicação de tipos
 Organização das colunas

            │
            ▼

     TRANSFORMAÇÃO (TRN)

 Aplicação dos tratamentos
 Padronização dos dados

            │
            ▼

     QUALIDADE (QA)

 Validação
 Registro de ocorrências
 Classificação dos registros

            │
            ▼

    NORMALIZAÇÃO (NRM)

 Deduplicação
 Chaves de negócio

            │
            ▼

      MODELO DIMENSIONAL

 Dimensões
 Fatos
 Modelo analítico
```

---

# Como funciona

O comportamento do framework é definido através de um **Schema declarativo**.

Em vez de implementar regras específicas para cada tabela, o usuário configura:

* tipo de cada coluna;
* tratamentos;
* validações;
* obrigatoriedade;
* ordem das colunas;
* chaves de negócio.

O framework transforma essa configuração em uma estrutura de execução que pode ser aplicada às tabelas correspondentes.

```text
Schema
   │
   ▼
Compilação
   │
   ▼
Pipeline
   │
   ▼
Execução
```

Isso permite separar:

```text
O QUE fazer
     │
     ▼
Schema

COMO executar
     │
     ▼
Framework
```

---

# Exemplo

Considere uma tabela `tbClientes` configurada no Schema.

O pipeline pode ser utilizado da seguinte forma:

```powerquery
stgClientes =
    fxStgAplicar(
        srcClientes,
        "tbClientes"
    );

trnClientes =
    fxTrnAplicar(
        stgClientes,
        "tbClientes"
    );

qaClientes =
    fxQaValidar(
        trnClientes,
        "tbClientes"
    );

nrmClientes =
    fxNrmAplicar(
        fxQaFiltrarPorStatus(
            qaClientes,
            "OK"
        ),
        "tbClientes"
    );
```

O comportamento de cada etapa é determinado pelo Schema da tabela.

A configuração pode definir, por exemplo:

```text
CPF
 ├── Tipo: TEXT
 ├── Obrigatório: SIM
 ├── Tratamentos: TRIM;DIGITS
 ├── Validações: CPFVAL
 └── Chave de negócio: SIM

Nome
 ├── Tipo: TEXT
 ├── Obrigatório: SIM
 └── Tratamentos: TRIM;PROPER

Estado
 ├── Tipo: TEXT
 ├── Tratamentos: TRIM;UPPER
 └── Validações: LIST(RS,SP,SC)
```

O código do pipeline permanece reutilizável.

---

# Estrutura do Projeto

```text
src/
    Fontes de dados

stg/
    Preparação estrutural

trn/
    Transformações

qa/
    Validação

nrm/
    Normalização

dim/
    Dimensões

fato/
    Tabelas fato

cfg/
    Configurações

fx/
    Funções do framework

docs/
    Documentação
```

A estrutura física pode variar conforme a organização do projeto e a forma de distribuição dos componentes.

---

# Documentação

A documentação técnica está organizada em módulos independentes.

| Documento                                  | Conteúdo                                    |
| ------------------------------------------ | ------------------------------------------- |
| [Getting Started](docs/getting-started.md) | Primeiros passos                            |
| [Architecture](docs/architecture.md)       | Arquitetura e responsabilidades das camadas |
| [Schema](docs/schema.md)                   | Estrutura e configuração do Schema          |
| [Pipeline](docs/pipeline.md)               | Compilação e execução do pipeline           |
| [Operators](docs/operators.md)             | Tratamentos e validações disponíveis        |
| [Performance](docs/performance.md)         | Estratégias de performance                  |
| [Examples](docs/examples.md)               | Exemplos de utilização                      |
| [Troubleshooting](docs/troubleshooting.md) | Diagnóstico de problemas                    |

---

# Source Available

O repositório público disponibiliza parte do projeto para permitir:

* inspeção do código Power Query;
* avaliação da arquitetura;
* compreensão do funcionamento;
* consulta da documentação;
* avaliação técnica antes da aquisição.

O código-fonte público **não representa necessariamente a distribuição comercial completa**.

O produto comercial pode conter componentes adicionais necessários para utilização completa do framework.

---

# Distribuição Comercial

A distribuição comercial é fornecida após a confirmação da aquisição da licença.

O pacote comercial pode incluir:

```text
Excel Data Framework DBB.xlsx
        │
        ├── Estrutura Excel
        ├── Consultas Power Query
        ├── Configurações
        ├── Tabelas do framework
        ├── Funções
        └── Componentes adicionais
```

O arquivo `.xlsx` completo **não faz parte do repositório público**.

Essa separação existe para manter o repositório público como espaço de documentação e avaliação técnica, enquanto o projeto Excel completo permanece como parte da distribuição comercial.

---

# Licenciamento

O Excel Data Framework DBB é distribuído sob uma **licença comercial própria**.

A licença padrão é:

```text
Tipo:
Licença comercial perpétua

Escopo:
Por usuário

Uso:
Comercial interno

Modificação:
Permitida

Transferência:
Não permitida sem autorização

Redistribuição:
Não permitida sem autorização

Sublicenciamento:
Não permitido
```

Uma licença válida concede ao usuário licenciado o direito de utilizar o framework para suas próprias atividades comerciais internas.

O usuário licenciado pode modificar livremente o framework para suas necessidades internas.

A licença não concede o direito de revender, sublicenciar, redistribuir ou disponibilizar publicamente o framework ou uma parte substancial dele.

Consulte [`LICENSE`](LICENSE) para os termos completos.

---

# Código-fonte e Licença

A disponibilização pública de arquivos `.m` neste repositório **não deve ser interpretada como concessão automática de uma licença comercial**.

O repositório utiliza o modelo:

```text
Código-fonte público
        │
        ▼
Inspeção / avaliação
        │
        ▼
Aquisição da licença
        │
        ▼
Distribuição comercial
        │
        ▼
Uso comercial interno
```

A licença comercial é concedida de acordo com os termos definidos no arquivo `LICENSE` e com a confirmação da respectiva aquisição.

---

# Princípios do Framework

O projeto foi desenvolvido seguindo alguns princípios fundamentais:

* Separation of Concerns
* Single Responsibility
* Metadata Driven Architecture
* Reutilização
* Configuração em vez de codificação
* Processamento declarativo
* Pipeline determinístico
* Baixo acoplamento
* Alta legibilidade
* Facilidade de manutenção
* Processamento nativo sempre que apropriado

---

# Performance

O framework foi projetado considerando as características de execução do Power Query.

Entre as principais estratégias estão:

* processamento por coluna;
* utilização de operações nativas do Power Query;
* pré-compilação dos pipelines;
* bufferização seletiva de metadados;
* redução de passagens sobre os dados;
* validação consolidada;
* preservação da avaliação lazy quando apropriado.

Performance não é tratada como uma propriedade absoluta do framework. O desempenho final depende também de fatores como:

* volume de dados;
* origem dos dados;
* complexidade das transformações;
* quantidade de validações;
* estrutura das consultas;
* ambiente de execução.

Consulte [Performance](docs/performance.md) para detalhes.

---

# Situação do Projeto

O framework encontra-se em desenvolvimento ativo.

O objetivo é oferecer uma arquitetura reutilizável para construção de pipelines de dados em Power Query, mantendo simplicidade para projetos menores e uma estrutura organizada para cenários mais complexos.

---

# Roadmap

Planejamentos futuros incluem:

* novos operadores de transformação;
* novos validadores;
* expansão da documentação;
* testes automatizados;
* novos exemplos práticos;
* templates de projetos;
* suporte a múltiplas fontes de dados;
* melhorias contínuas de performance.

O roadmap pode ser atualizado conforme a evolução do projeto.

---

# Contribuindo

Sugestões, correções e propostas de melhoria são bem-vindas.

Antes de contribuir, consulte:

[Contributing Guide](CONTRIBUTING.md)

Para problemas ou propostas de melhoria, utilize as Issues do repositório.

Pull Requests podem ser utilizados para alterações compatíveis com o escopo do projeto e com os termos de contribuição definidos no repositório.

---

# Reportando Vulnerabilidades

Problemas de segurança devem ser reportados de acordo com a política definida em:

[Security Policy](SECURITY.md)

Evite publicar informações sensíveis ou detalhes exploráveis de uma vulnerabilidade em uma Issue pública.

---

# Autor

Desenvolvido por **Daniel Becker Bighelini**.

---

# Licença

Copyright © [2026] Daniel Becker Bighelini.

Este projeto é distribuído sob os termos da **Licença Comercial DBB**.

Consulte [`LICENSE`](LICENSE) para os termos completos de uso, modificação, redistribuição e licenciamento.

---

# Links

* [Documentação](docs/)
* [Getting Started](docs/getting-started.md)
* [Architecture](docs/architecture.md)
* [Schema](docs/schema.md)
* [Pipeline](docs/pipeline.md)
* [Operators](docs/operators.md)
* [Performance](docs/performance.md)
* [Examples](docs/examples.md)
* [Troubleshooting](docs/troubleshooting.md)
* [Contributing](CONTRIBUTING.md)
* [Security](SECURITY.md)
* [License](LICENSE.md)
