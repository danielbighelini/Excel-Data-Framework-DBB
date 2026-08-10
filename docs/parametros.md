# Parâmetros de Configuração

## 1. Visão geral

Os parâmetros de configuração definem o comportamento e os recursos
utilizados pelo framework. Eles são organizados por domínio funcional e
consumidos pelas funções de configuração, validação, tratamento e
carregamento de dados.

A configuração deve ser entendida em dois níveis:

-   **parâmetro:** define uma propriedade individual do framework;
-   **grupo de parâmetros:** define conjuntamente uma fonte, recurso ou
    comportamento.

Um parâmetro pode ser obrigatório em todas as execuções ou
**condicional**, dependendo da fonte ou do recurso selecionado.

------------------------------------------------------------------------

## 2. Regras gerais

### 2.1 Tipo

O valor configurado deve ser compatível com o tipo definido para o
parâmetro.

Tipos utilizados atualmente:

-   `Texto`
-   `Booleano`
-   `Número inteiro`
-   `Data`

### 2.2 Obrigatoriedade

A obrigatoriedade deve ser interpretada no contexto do parâmetro.

Um parâmetro pode ser:

-   **Sempre obrigatório:** deve possuir valor em qualquer configuração;
-   **Condicional:** somente é necessário quando determinado recurso ou
    fonte está sendo utilizado;
-   **Opcional:** pode permanecer sem valor.

### 2.3 Valores permitidos

Quando a coluna `Permitidos` estiver definida, o valor configurado deve
pertencer ao conjunto especificado.

### 2.4 Parâmetros não utilizados

Parâmetros pertencentes a uma fonte que não está sendo utilizada não
precisam ser preenchidos, salvo quando outra regra do framework
determinar o contrário.

------------------------------------------------------------------------

# 3. Parâmetros de dados

## `Dados_Fonte`

  Propriedade          Valor
  -------------------- ----------------------------
  Tipo                 Texto
  Obrigatório          Sim
  Valores permitidos   `Arquivos`; `REST`; `SGBD`

Define a tecnologia utilizada como fonte dos dados.

O valor selecionado determina qual conjunto de parâmetros específicos de
fonte será utilizado.

### Valores

-   `Arquivos` --- dados obtidos a partir de arquivos;
-   `REST` --- dados obtidos por meio de uma API REST;
-   `SGBD` --- dados obtidos a partir de um sistema gerenciador de banco
    de dados.

------------------------------------------------------------------------

## `Dados_Carregar`

  Propriedade          Valor
  -------------------- --------------
  Tipo                 Booleano
  Obrigatório          Não
  Valores permitidos   `Sim`; `Não`

Determina se os dados da fonte configurada devem ser carregados pelo
framework.

------------------------------------------------------------------------

# 4. Parâmetros de arquivos

Os parâmetros deste grupo são aplicáveis quando
`Dados_Fonte = Arquivos`.

## `Arquivo_Origem`

  Propriedade          Valor
  -------------------- ---------------------------------
  Tipo                 Texto
  Obrigatório          Condicional
  Valores permitidos   `Local`; `Remoto`; `Sharepoint`

Define a origem física dos arquivos.

### Dependência

Aplicável quando:

``` text
Dados_Fonte = Arquivos
```

------------------------------------------------------------------------

## `Arquivo_Formato`

  Propriedade          Valor
  -------------------- -------------------------------------
  Tipo                 Texto
  Obrigatório          Condicional
  Valores permitidos   `CSV`; `XLSX`; `XLS`; `XLSM`; `PDF`

Define o formato dos arquivos de entrada.

### Dependência

Aplicável quando:

``` text
Dados_Fonte = Arquivos
```

------------------------------------------------------------------------

## `Arquivo_Delimitador`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional
  Exemplo       `;`

Define o delimitador utilizado na interpretação de arquivos delimitados,
como CSV.

### Dependência

Principalmente aplicável quando:

``` text
Arquivo_Formato = CSV
```

------------------------------------------------------------------------

## `Arquivo_Codificador`

  Propriedade   Valor
  ------------- ----------------
  Tipo          Número inteiro
  Obrigatório   Condicional
  Exemplo       `65001`

Define a codificação utilizada na leitura de arquivos de texto.

O valor `65001` corresponde à codificação UTF-8 no contexto de códigos
de página do Windows.

### Dependência

Aplicável principalmente a formatos de texto, como:

``` text
Arquivo_Formato = CSV
```

------------------------------------------------------------------------

# 5. Parâmetros de localização

## `Local_Pasta`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional

Define o caminho da pasta utilizada quando os arquivos estão armazenados
localmente.

O caminho pode ser normalizado pelo framework antes de ser utilizado
pelas funções de acesso ao sistema de arquivos.

### Dependência

``` text
Dados_Fonte = Arquivos
Arquivo_Origem = Local
```

------------------------------------------------------------------------

## `Remota_Pasta`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional

Define o caminho da pasta utilizada quando os arquivos estão armazenados
em uma localização remota acessível pelo sistema de arquivos.

### Dependência

``` text
Dados_Fonte = Arquivos
Arquivo_Origem = Remoto
```

------------------------------------------------------------------------

# 6. Parâmetros SharePoint

Os parâmetros deste grupo são utilizados para localizar arquivos em um
ambiente SharePoint.

## `SharePoint_Site`

  Propriedade   Valor
  ------------- -------------------------------------------------
  Tipo          Texto
  Obrigatório   Condicional
  Exemplo       `https://tenant.sharepoint.com/sites/NomeSite/`

Define a URL do site SharePoint que contém os arquivos.

### Dependência

``` text
Dados_Fonte = Arquivos
Arquivo_Origem = Sharepoint
```

------------------------------------------------------------------------

## `SharePoint_Biblioteca`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional
  Exemplo       `Documents`

Define a biblioteca de documentos utilizada como origem.

### Dependência

``` text
Arquivo_Origem = Sharepoint
```

------------------------------------------------------------------------

## `SharePoint_Pasta`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional

Define o caminho da pasta dentro da biblioteca SharePoint.

### Dependência

``` text
Arquivo_Origem = Sharepoint
```

------------------------------------------------------------------------

# 7. Parâmetros de SGBD

Os parâmetros deste grupo são utilizados quando os dados são obtidos
diretamente de um sistema gerenciador de banco de dados.

## `SGBD_Tipo`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional
  Exemplo       `ORACLE`

Define o tipo de sistema gerenciador de banco de dados.

### Dependência

``` text
Dados_Fonte = SGBD
```

------------------------------------------------------------------------

## `SGBD_Host`

  Propriedade   Valor
  ------------- --------------------------------
  Tipo          Texto
  Obrigatório   Condicional
  Exemplo       `sgbdoracle.com.br:1525/CLI03`

Define o endereço do servidor e, quando aplicável, as informações de
porta ou serviço necessárias para estabelecer a conexão.

------------------------------------------------------------------------

## `SGBD_Banco`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional

Define o banco, catálogo, serviço ou outra identificação equivalente
utilizada pelo SGBD.

------------------------------------------------------------------------

## `SGBD_SQL`

  Propriedade   Valor
  ------------- --------------------------------------------------------
  Tipo          Texto
  Obrigatório   Condicional
  Exemplo       `SELECT * FROM TABELA_CLIENTES WHERE TIPO = 'CLIENTE'`

Define a consulta SQL utilizada para obtenção dos dados.

### Dependência

``` text
Dados_Fonte = SGBD
```

------------------------------------------------------------------------

# 8. Parâmetros REST

Os parâmetros deste grupo controlam o consumo de APIs.

## `REST_Protocolo`

  Propriedade          Valor
  -------------------- -------------------------
  Tipo                 Texto
  Obrigatório          Condicional
  Exemplo              `REST`
  Valores permitidos   `REST`; `ODATA`; `SOAP`

Define o protocolo ou tecnologia de integração utilizada.

### Dependência

``` text
Dados_Fonte = REST
```

------------------------------------------------------------------------

## `REST_Endpoint_Base`

  Propriedade   Valor
  ------------- -------------------------
  Tipo          Texto
  Obrigatório   Condicional
  Exemplo       `https://dummyjson.com`

Define a URL base do serviço.

------------------------------------------------------------------------

## `REST_Endpoint_Path`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional
  Exemplo       `products`

Define o caminho relativo do endpoint.

Quando combinado com `REST_Endpoint_Base`, forma o endereço utilizado
para a requisição.

------------------------------------------------------------------------

## `REST_Formato`

  Propriedade          Valor
  -------------------- ---------------
  Tipo                 Texto
  Obrigatório          Condicional
  Valores permitidos   `JSON`; `XML`

Define o formato esperado da resposta do serviço.

------------------------------------------------------------------------

## `REST_Autenticacao`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional
  Exemplo       `None`

Define o mecanismo de autenticação utilizado pela requisição.

As opções efetivamente disponíveis devem ser determinadas pela
configuração suportada pela implementação do framework.

------------------------------------------------------------------------

## `REST_Paginacao`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional
  Exemplo       `None`

Define a estratégia utilizada para obtenção de respostas paginadas.

Quando não houver paginação, pode ser utilizado o valor `None`.

------------------------------------------------------------------------

## `REST_Timeout`

  Propriedade   Valor
  ------------- ----------------
  Tipo          Número inteiro
  Obrigatório   Condicional
  Exemplo       `30`

Define o tempo limite utilizado para a requisição.

A unidade deve permanecer consistente com a implementação da função
responsável pela chamada REST.

------------------------------------------------------------------------

## `REST_Tentativas`

  Propriedade   Valor
  ------------- ----------------
  Tipo          Número inteiro
  Obrigatório   Condicional
  Exemplo       `3`

Define a quantidade de tentativas permitidas em operações sujeitas a
nova tentativa.

------------------------------------------------------------------------

## `REST_Token`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional

Armazena o token utilizado para autenticação quando o mecanismo
configurado exigir esse tipo de credencial.

------------------------------------------------------------------------

## `REST_ApiKey`

  Propriedade   Valor
  ------------- -------------
  Tipo          Texto
  Obrigatório   Condicional

Armazena a chave utilizada para autenticação quando o serviço exigir uma
API Key.

> Credenciais e segredos não devem ser registrados em documentação
> pública, exemplos versionados ou arquivos de configuração distribuídos
> com o projeto.

------------------------------------------------------------------------

# 9. Parâmetros de cultura

Os parâmetros de cultura estabelecem as convenções utilizadas pelo
framework para interpretação e apresentação de valores.

## `Cultura_Padrao`

  Propriedade          Valor
  -------------------- ------------------
  Tipo                 Texto
  Obrigatório          Sim
  Exemplo              `pt-BR`
  Valores permitidos   `pt-BR`; `en-US`

Define a cultura padrão utilizada pelo framework.

------------------------------------------------------------------------

## `Cultura_Separador_Lista`

  Propriedade   Valor
  ------------- -------
  Tipo          Texto
  Obrigatório   Sim
  Exemplo       `;`

Define o separador utilizado para representar elementos de listas em
valores textuais de configuração.

------------------------------------------------------------------------

## `Cultura_Separador_Parametros`

  Propriedade   Valor
  ------------- -------
  Tipo          Texto
  Obrigatório   Sim
  Exemplo       `,`

Define o separador utilizado na representação textual de parâmetros.

------------------------------------------------------------------------

## `Cultura_Simbolo_Moeda`

  Propriedade   Valor
  ------------- -------
  Tipo          Texto
  Obrigatório   Sim
  Exemplo       `R$`

Define o símbolo monetário utilizado pelo framework.

------------------------------------------------------------------------

## `Cultura_Simbolo_Percentual`

  Propriedade   Valor
  ------------- -------
  Tipo          Texto
  Obrigatório   Sim
  Exemplo       `%`

Define o símbolo utilizado para representar percentuais.

------------------------------------------------------------------------

## `Cultura_Negativo_Parenteses`

  Propriedade          Valor
  -------------------- --------------
  Tipo                 Booleano
  Obrigatório          Sim
  Exemplo              `Sim`
  Valores permitidos   `Sim`; `Não`

Determina se valores negativos podem ser representados utilizando
parênteses.

------------------------------------------------------------------------

## `Cultura_Formato_Data`

  Propriedade   Valor
  ------------- --------------
  Tipo          Texto
  Obrigatório   Sim
  Exemplo       `dd/mm/aaaa`

Define o formato padrão utilizado para representação de datas.

------------------------------------------------------------------------

## `Cultura_Formato_Hora`

  Propriedade   Valor
  ------------- ------------
  Tipo          Texto
  Obrigatório   Sim
  Exemplo       `hh:mm:ss`

Define o formato padrão utilizado para representação de horários.

------------------------------------------------------------------------

## `Cultura_Formato_DataHora`

  Propriedade   Valor
  ------------- --------------------
  Tipo          Texto
  Obrigatório   Sim
  Exemplo       `dd/mm/aaaa hh:mm`

Define o formato padrão utilizado para representação combinada de data e
hora.

------------------------------------------------------------------------

## `Cultura_Primeiro_Dia_Semana`

  Propriedade   Valor
  ------------- -----------
  Tipo          Texto
  Obrigatório   Sim
  Exemplo       `Domingo`

Define o primeiro dia da semana utilizado nas operações relacionadas a
calendário.

------------------------------------------------------------------------

## `Cultura_Idioma`

  Propriedade   Valor
  ------------- ----------------------
  Tipo          Texto
  Obrigatório   Sim
  Exemplo       `Português (Brasil)`

Define o idioma utilizado pelo framework para operações que dependam de
uma identificação linguística.

------------------------------------------------------------------------

# 10. Parâmetros de calendário

## `Calendario_Data_Inicial`

  Propriedade   Valor
  ------------- --------------
  Tipo          Data
  Obrigatório   Não
  Exemplo       `01/01/2026`

Define a data inicial do calendário gerado pelo framework.

------------------------------------------------------------------------

## `Calendario_Data_Final`

  Propriedade   Valor
  ------------- --------------
  Tipo          Data
  Obrigatório   Não
  Exemplo       `30/06/2026`

Define a data final do calendário gerado pelo framework.

A data final deve ser igual ou posterior à data inicial.

------------------------------------------------------------------------

## `Calendario_Inicio_Fiscal`

  Propriedade          Valor
  -------------------- ----------------
  Tipo                 Número inteiro
  Obrigatório          Não
  Exemplo              `1`
  Valores permitidos   `1` a `12`

Define o mês de início do ano fiscal.

O valor representa o número do mês:

``` text
1  = Janeiro
2  = Fevereiro
...
12 = Dezembro
```

------------------------------------------------------------------------

# 11. Matriz de dependências

A tabela abaixo resume as principais relações entre os parâmetros.

  -----------------------------------------------------------------------
  Condição                            Parâmetros relacionados
  ----------------------------------- -----------------------------------
  `Dados_Fonte = Arquivos`            `Arquivo_Origem`,
                                      `Arquivo_Formato`,
                                      `Arquivo_Delimitador`,
                                      `Arquivo_Codificador`

  `Arquivo_Origem = Local`            `Local_Pasta`

  `Arquivo_Origem = Remoto`           `Remota_Pasta`

  `Arquivo_Origem = Sharepoint`       `SharePoint_Site`,
                                      `SharePoint_Biblioteca`,
                                      `SharePoint_Pasta`

  `Arquivo_Formato = CSV`             `Arquivo_Delimitador`,
                                      `Arquivo_Codificador`

  `Dados_Fonte = SGBD`                `SGBD_Tipo`, `SGBD_Host`,
                                      `SGBD_Banco`, `SGBD_SQL`

  `Dados_Fonte = REST`                parâmetros `REST_*`

  `REST_Autenticacao` exige token     `REST_Token`

  `REST_Autenticacao` exige API Key   `REST_ApiKey`
  
  -----------------------------------------------------------------------

------------------------------------------------------------------------

# 12. Convenção de nomes

Os parâmetros seguem uma convenção baseada em:

``` text
<Categoria>_<Propriedade>
```

Exemplos:

``` text
Dados_Fonte
Arquivo_Formato
SharePoint_Site
SGBD_Host
REST_Timeout
Cultura_Padrao
Calendario_Data_Final
```

Essa convenção permite identificar a área funcional do parâmetro
diretamente pelo nome e facilita sua classificação, validação e consumo
pelas estruturas `cfg*` do framework.

------------------------------------------------------------------------

# 13. Relação com a configuração operacional

Este documento descreve o **contrato dos parâmetros**.

A tabela de configuração contém os valores efetivamente utilizados em
uma determinada execução.

Portanto:

``` text
Documentação
    ↓
define o significado e as regras
    ↓
Tabela de parâmetros
    ↓
define os valores atuais
    ↓
Framework
    ↓
valida e utiliza a configuração
```

A documentação não deve ser considerada a fonte do valor atual de um
parâmetro. A configuração operacional deve permanecer na estrutura de
configuração do framework.

------------------------------------------------------------------------

# 14. Evolução dos parâmetros

Ao adicionar ou alterar um parâmetro, devem ser avaliados:

1.  nome e convenção de nomenclatura;
2.  tipo;
3.  obrigatoriedade;
4.  valores permitidos;
5.  valor padrão, quando aplicável;
6.  dependências;
7.  impacto sobre validações;
8.  impacto sobre funções `fx*`;
9.  impacto sobre estruturas `cfg*`;
10. atualização desta documentação.

Parâmetros removidos ou descontinuados devem ser identificados como
obsoletos antes de sua remoção definitiva, quando houver risco de quebra
de configurações existentes.
