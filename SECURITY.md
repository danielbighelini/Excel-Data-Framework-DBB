# Política de Segurança

## 1. Visão Geral

A segurança do **Excel Data Framework DBB** depende tanto do código do framework quanto da forma como o projeto é configurado e utilizado.

Este documento define como reportar vulnerabilidades de segurança relacionadas ao projeto e quais informações devem ser fornecidas para permitir uma análise adequada.

---

# 2. Reportando uma Vulnerabilidade

Se você identificar uma possível vulnerabilidade de segurança, **não publique os detalhes em uma Issue pública do GitHub**.

O reporte deve ser enviado diretamente ao responsável pelo projeto através do canal privado indicado abaixo:

**E-mail:** danielbighelini@gmail.com

Ao realizar o reporte, utilize um assunto que identifique claramente que se trata de uma questão de segurança.

Exemplo:

```text
[SECURITY] Descrição resumida da vulnerabilidade
```

---

# 3. Informações Recomendadas

Sempre que possível, inclua no reporte:

* descrição da vulnerabilidade;
* componente afetado;
* versão do framework;
* versão do Excel;
* versão do Power Query, quando relevante;
* ambiente em que o problema foi identificado;
* passos necessários para reproduzir o problema;
* comportamento esperado;
* comportamento observado;
* impacto potencial;
* evidências técnicas;
* código ou configuração mínima necessária para reprodução.

Um reporte estruturado pode seguir o seguinte formato:

```text
Componente:
[Nome do componente]

Versão:
[Versão]

Descrição:
[Descrição do problema]

Impacto:
[Impacto potencial]

Reprodução:
[Passos para reproduzir]

Comportamento esperado:
[Comportamento esperado]

Comportamento observado:
[Comportamento observado]

Evidências:
[Logs, código ou outras evidências]
```

---

# 4. Não Envie Informações Sensíveis

Não inclua no reporte:

* senhas;
* tokens;
* chaves de API;
* credenciais;
* dados pessoais reais;
* dados bancários;
* informações financeiras confidenciais;
* arquivos contendo dados de clientes;
* informações proprietárias de terceiros.

Quando for necessário fornecer dados para reprodução, utilize dados fictícios ou anonimizados.

---

# 5. Dados de Exemplo

Para demonstrar uma vulnerabilidade ou problema de segurança, prefira criar um caso mínimo reproduzível.

Por exemplo:

```text
Fonte
  │
  ▼
Schema mínimo
  │
  ▼
Pipeline
  │
  ▼
Comportamento vulnerável
```

O objetivo é fornecer informações suficientes para reproduzir o problema sem expor dados reais.

---

# 6. Componentes Comerciais

O repositório público pode não conter todos os componentes da distribuição comercial.

O arquivo Excel completo (`.xlsx`) e outros componentes comerciais podem ser distribuídos separadamente aos usuários licenciados.

Caso uma vulnerabilidade esteja relacionada a um componente que não esteja disponível no repositório público, informe no reporte:

* qual componente está envolvido;
* versão utilizada;
* contexto em que o componente foi obtido;
* comportamento observado.

Não publique o componente comercial em uma Issue ou Pull Request.

---

# 7. Vulnerabilidades no Código Power Query

Problemas relacionados ao código Power Query podem envolver, entre outros:

* execução de código não esperado;
* processamento inseguro de entradas;
* exposição indevida de informações;
* tratamento inadequado de credenciais;
* consultas que acessam fontes não autorizadas;
* manipulação insegura de parâmetros;
* comportamento inesperado causado por dados malformados.

Quando o problema estiver relacionado a uma função específica, informe:

```text
Função:
[nome da função]

Consulta:
[nome da consulta]

Entrada:
[descrição da entrada]

Resultado observado:
[descrição]

Resultado esperado:
[descrição]
```

---

# 8. Credenciais e Fontes de Dados

O framework pode ser utilizado com diferentes fontes de dados.

Credenciais de acesso às fontes **não devem ser armazenadas no código Power Query ou publicadas no repositório**.

Não inclua no código:

```text
senhas
tokens
API keys
connection strings contendo credenciais
credenciais de banco de dados
```

Quando possível, utilize os mecanismos de autenticação e gerenciamento de credenciais disponibilizados pelo ambiente de execução.

---

# 9. Arquivos Excel

O arquivo `.xlsx` completo do produto comercial não deve ser publicado no repositório público.

Também não devem ser enviados ao GitHub arquivos Excel contendo:

* credenciais;
* dados pessoais;
* dados financeiros;
* dados de clientes;
* configurações confidenciais;
* informações proprietárias;
* componentes comerciais não publicados.

O fato de um arquivo ser tecnicamente necessário para reproduzir determinado comportamento não significa que ele deva ser publicado.

Sempre que possível, substitua o arquivo por um exemplo mínimo e anonimizado.

---

# 10. Divulgação Responsável

Vulnerabilidades devem ser tratadas de forma responsável.

O processo recomendado é:

```text
Identificação
     │
     ▼
Reporte privado
     │
     ▼
Análise
     │
     ▼
Correção
     │
     ▼
Validação
     │
     ▼
Atualização
     │
     ▼
Divulgação apropriada
```

Evite divulgar publicamente detalhes técnicos de uma vulnerabilidade antes que exista uma oportunidade razoável para análise e correção.

---

# 11. Prazo de Resposta

O responsável pelo projeto procurará analisar os relatos de segurança em prazo razoável.

O tempo necessário para resposta e correção pode variar conforme:

* gravidade;
* complexidade;
* possibilidade de reprodução;
* componente afetado;
* necessidade de alteração arquitetural;
* disponibilidade de uma correção segura.

O envio de um reporte não implica garantia de um prazo específico para correção.

---

# 12. Classificação de Severidade

A severidade será avaliada individualmente.

Como referência, podem ser considerados:

### Baixa

Problemas com impacto limitado ou que exigem condições incomuns para exploração.

### Média

Problemas capazes de produzir impacto relevante, mas que dependem de condições adicionais.

### Alta

Problemas que podem comprometer significativamente dados, execução ou segurança do ambiente.

### Crítica

Problemas que podem permitir comprometimento grave do ambiente, exposição significativa de dados ou execução não autorizada em condições realistas.

A classificação final é responsabilidade do mantenedor do projeto e pode depender do contexto de utilização.

---

# 13. Pull Requests de Segurança

Correções de segurança podem ser propostas através de Pull Requests quando isso não expuser informações sensíveis ou detalhes de uma vulnerabilidade ainda não corrigida.

Para vulnerabilidades que exigem divulgação privada, o processo deve começar pelo reporte direto ao mantenedor.

Após a análise, o mantenedor poderá solicitar ou aceitar uma correção através de Pull Request.

---

# 14. Dependências e Componentes de Terceiros

Problemas de segurança podem estar relacionados a componentes externos utilizados pelo projeto.

Quando identificar uma vulnerabilidade originada em uma dependência de terceiros, informe:

* componente;
* versão;
* origem;
* descrição do problema;
* referência pública da vulnerabilidade, quando existente.

Não inclua código de terceiros vulnerável desnecessariamente no reporte.

---

# 15. Responsabilidade do Usuário

O framework é uma ferramenta de processamento de dados.

A segurança do ambiente final também depende da configuração realizada pelo usuário.

O usuário é responsável por:

* proteger suas credenciais;
* controlar o acesso aos arquivos;
* proteger os dados processados;
* utilizar fontes confiáveis;
* revisar conexões externas;
* evitar publicar informações confidenciais;
* configurar adequadamente o ambiente de execução.

Uma configuração insegura do ambiente não deve ser confundida automaticamente com uma vulnerabilidade do framework.

---

# 16. Escopo

Esta política abrange problemas de segurança relacionados ao:

* código Power Query disponibilizado pelo projeto;
* funções do framework;
* pipeline;
* Schema;
* mecanismos de configuração;
* documentação quando esta induzir a uma configuração insegura;
* componentes públicos mantidos neste repositório.

Problemas exclusivamente relacionados ao ambiente do usuário devem ser analisados separadamente.

---

# 17. Contato

**Responsável:** Bighelini

**E-mail de segurança:** [E-MAIL DE SEGURANÇA]

Caso exista um endereço específico para questões de segurança, utilize esse endereço em vez de Issues públicas.

---

# 18. Histórico

| Versão | Data   | Alteração                        |
| ------ | ------ | -------------------------------- |
| 1.0    | [DATA] | Criação da política de segurança |

---

## Princípio

> **Vulnerabilidades devem ser reportadas de forma privada, sem exposição de dados sensíveis ou componentes comerciais, permitindo que o problema seja analisado e corrigido antes de sua divulgação pública.**
