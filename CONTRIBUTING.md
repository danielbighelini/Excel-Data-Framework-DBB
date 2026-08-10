# Contributing to Excel Data Framework DBB

Obrigado pelo interesse em contribuir com o **Excel Data Framework DBB**.

O projeto utiliza uma arquitetura orientada por Schema, pipelines compilados e componentes reutilizáveis em Power Query. Contribuições devem preservar esses princípios e evitar introduzir lógica específica de projeto dentro do framework.

Este documento descreve como propor alterações, corrigir problemas e contribuir com o código e a documentação disponibilizados neste repositório.

---

# 1. O que pode receber contribuições

Contribuições podem ser realizadas principalmente nas seguintes áreas:

* funções Power Query M;
* operadores de transformação;
* validadores;
* pipeline;
* Schema;
* documentação;
* exemplos;
* testes;
* correções de bugs;
* melhorias de performance;
* organização do código.

Contribuições devem permanecer compatíveis com a arquitetura e os objetivos do framework.

---

# 2. Componentes comerciais

O repositório público não contém necessariamente todos os componentes da distribuição comercial.

Em particular, o projeto Excel completo (`.xlsx`) e outros componentes comerciais podem ser distribuídos separadamente aos usuários licenciados.

Contribuições realizadas neste repositório não concedem ao colaborador qualquer direito de acesso, distribuição ou sublicenciamento desses componentes comerciais.

Os termos de utilização do software estão definidos em [`LICENSE`](LICENSE).

---

# 3. Antes de contribuir

Antes de implementar uma alteração:

1. consulte a documentação existente;
2. procure Issues relacionadas;
3. verifique se o comportamento já está documentado;
4. confirme se a alteração realmente pertence ao framework;
5. considere o impacto sobre o Schema e o pipeline;
6. avalie possíveis impactos de performance.

Documentação relevante:

* [Architecture](docs/architecture.md)
* [Schema](docs/schema.md)
* [Pipeline](docs/pipeline.md)
* [Operators](docs/operators.md)
* [Performance](docs/performance.md)
* [Examples](docs/examples.md)
* [Troubleshooting](docs/troubleshooting.md)

---

# 4. Issues

Utilize uma Issue para:

* reportar bugs;
* propor novos operadores;
* propor novas funcionalidades;
* discutir alterações arquiteturais;
* relatar problemas de performance;
* sugerir melhorias na documentação.

Antes de abrir uma nova Issue, verifique se já existe uma discussão sobre o mesmo assunto.

---

# 5. Reportando Bugs

Uma Issue de bug deve conter informações suficientes para reproduzir o problema.

Sempre que possível, informe:

```text
Descrição:
O que aconteceu?

Comportamento esperado:
O que deveria acontecer?

Comportamento observado:
O que aconteceu de fato?

Camada:
SRC / STG / TRN / QA / NRM / DIM / FATO

Reprodução:
Como reproduzir o problema?

Schema:
Qual configuração está envolvida?

Dados:
Qual estrutura de dados produz o problema?

Erro:
Mensagem completa apresentada pelo Power Query.
```

Quando relevante, inclua:

* operador utilizado;
* parâmetros;
* função envolvida;
* `_STG_Ocorrencias`;
* `_QA_Status`;
* `_QA_Ocorrencias`;
* informações de performance.

Não publique dados pessoais, credenciais, informações financeiras ou outros dados confidenciais.

---

# 6. Propondo Novos Operadores

Novos operadores devem possuir uma finalidade clara e reutilizável.

Antes de criar um operador, verifique se a necessidade pode ser atendida pela combinação de operadores existentes.

Um novo operador deve, quando aplicável, possuir:

* nome consistente;
* assinatura definida;
* parâmetros documentados;
* comportamento determinístico;
* tratamento adequado de `null`;
* comportamento documentado;
* integração com o mecanismo de configuração;
* exemplos de utilização;
* testes correspondentes.

Exemplo conceitual:

```text
Schema
   │
   ▼
NOVOOPERADOR(parametros)
   │
   ▼
Pipeline
   │
   ▼
Função M
```

A definição do operador deve permanecer separada da lógica específica de uma única tabela ou projeto.

---

# 7. Alterações no Schema

Alterações no Schema devem ser avaliadas com cuidado porque o Schema determina o comportamento do pipeline.

Antes de alterar uma estrutura de Schema, considere:

* compatibilidade com configurações existentes;
* impacto sobre `cfgPipeline`;
* impacto sobre tratamentos;
* impacto sobre validações;
* impacto sobre chaves de negócio;
* impacto sobre STG, TRN, QA e NRM.

Alterações que modifiquem o significado de uma propriedade existente devem ser discutidas antes de serem implementadas.

---

# 8. Alterações no Pipeline

Alterações no pipeline devem preservar a separação entre:

```text
Configuração
     │
     ▼
Compilação
     │
     ▼
Execução
```

Evite incorporar regras específicas de negócio diretamente no mecanismo de execução quando essas regras puderem ser representadas através do Schema ou de operadores reutilizáveis.

Alterações no pipeline também devem considerar:

* reutilização;
* desempenho;
* avaliação lazy;
* bufferização;
* processamento por coluna;
* compatibilidade com as camadas existentes.

---

# 9. Performance

Alterações de performance devem ser justificadas por comportamento observado ou por uma necessidade arquitetural clara.

Não introduza `Table.Buffer` ou `List.Buffer` indiscriminadamente.

Antes de considerar uma alteração de performance, avalie:

* número de passagens sobre os dados;
* processamento por coluna;
* pré-compilação;
* materialização;
* tamanho das estruturas;
* reutilização de resultados;
* impacto sobre memória.

Quando possível, forneça uma comparação objetiva:

```text
Antes:
Tempo = X

Depois:
Tempo = Y

Volume:
N registros / N colunas
```

Uma otimização só deve ser considerada efetiva quando houver evidência de melhoria.

Consulte [Performance](docs/performance.md).

---

# 10. Documentação

Alterações de comportamento devem ser acompanhadas da documentação correspondente.

Por exemplo:

```text
Novo operador
    ↓
Código
    +
operators.md
    +
examples.md
```

Uma alteração arquitetural pode exigir atualização de:

* `architecture.md`;
* `schema.md`;
* `pipeline.md`;
* `performance.md`;
* `examples.md`;
* `troubleshooting.md`.

Evite duplicar a mesma explicação em vários documentos quando uma referência cruzada for suficiente.

---

# 11. Exemplos

Exemplos devem demonstrar funcionalidades reais do framework.

Prefira exemplos que mostrem:

```text
Schema
   ↓
Pipeline
   ↓
STG
   ↓
TRN
   ↓
QA
   ↓
NRM
```

Evite exemplos excessivamente específicos de um único projeto ou domínio de negócio quando a funcionalidade puder ser demonstrada de forma genérica.

---

# 12. Testes

Alterações que introduzam ou modifiquem comportamento devem ser acompanhadas de testes quando houver infraestrutura de testes disponível.

Novas funções e operadores devem considerar, no mínimo:

* valores válidos;
* valores inválidos;
* `null`;
* valores vazios, quando aplicável;
* parâmetros;
* limites;
* comportamento inesperado.

Para validações, considere também:

* resultado positivo;
* resultado negativo;
* severidade;
* ocorrência gerada.

---

# 13. Pull Requests

Um Pull Request deve apresentar claramente:

* o problema que está sendo resolvido;
* a alteração realizada;
* o impacto esperado;
* testes executados;
* alterações de documentação;
* possíveis impactos de compatibilidade.

Use um título objetivo.

Exemplos:

```text
Add PHONEVAL validator
Fix parameter parsing in pipeline compiler
Improve QA validation performance
Update Schema documentation
Fix null handling in CPF validation
```

---

# 14. Escopo de um Pull Request

Prefira Pull Requests pequenos e focados.

Por exemplo:

```text
PR 1
Novo operador PHONEVAL

PR 2
Documentação do PHONEVAL

PR 3
Correção de performance no QA
```

Evite combinar em um único Pull Request:

* novo operador;
* refatoração completa;
* mudança arquitetural;
* alteração de documentação;
* mudanças não relacionadas.

Isso dificulta revisão, teste e manutenção.

---

# 15. Compatibilidade

Alterações devem evitar quebrar configurações existentes sem necessidade.

Antes de introduzir uma alteração incompatível, avalie:

```text
Comportamento atual
       │
       ▼
Configurações existentes
       │
       ▼
Nova implementação
       │
       ▼
Compatibilidade
```

Quando uma alteração incompatível for necessária, ela deve ser claramente documentada.

---

# 16. Commits

Prefira commits pequenos e semanticamente claros.

Exemplos:

```text
Add CPF validation
Fix pipeline operator parsing
Improve QA status handling
Update pipeline documentation
Optimize column transformation
```

Evite mensagens genéricas como:

```text
fix
update
changes
teste
alterações
```

Um commit deve permitir entender aproximadamente o que foi alterado sem abrir o diff.

---

# 17. Revisão de Código

Pull Requests podem ser revisados considerando:

### Arquitetura

A alteração respeita as responsabilidades das camadas?

### Reutilização

A implementação é reutilizável ou está acoplada a uma tabela específica?

### Schema

A regra deveria estar no Schema em vez do código?

### Pipeline

A alteração interfere na compilação ou execução?

### Performance

A alteração adiciona passagens, materialização ou processamento desnecessário?

### Manutenção

O código permanece legível e consistente com o restante do framework?

### Documentação

O comportamento alterado está documentado?

---

# 18. Segurança

Não publique no repositório:

* senhas;
* tokens;
* chaves de API;
* credenciais;
* dados pessoais;
* dados financeiros;
* arquivos contendo informações confidenciais;
* arquivos comerciais que não fazem parte da distribuição pública.

Problemas de segurança devem ser reportados de acordo com [`SECURITY.md`](SECURITY.md), e não publicados inicialmente em uma Issue pública.

---

# 19. Licenciamento das Contribuições

Ao enviar uma contribuição para este repositório, o colaborador deve possuir os direitos necessários para disponibilizar essa contribuição.

O colaborador não deve enviar:

* código de terceiros sem licença compatível;
* material protegido por direitos autorais sem autorização;
* informações confidenciais;
* código proprietário pertencente a outra organização.

Salvo acordo escrito em contrário, uma contribuição aceita passa a integrar o projeto sob os termos de licenciamento aplicáveis ao projeto.

O envio de uma contribuição não concede ao colaborador qualquer direito sobre componentes comerciais que não estejam presentes no repositório público.

---

# 20. Código de Terceiros

Antes de adicionar uma dependência, biblioteca, função derivada de terceiros ou outro componente externo, verifique sua licença.

A contribuição deve informar, quando aplicável:

* origem;
* licença;
* versão;
* modificações realizadas;
* requisitos de atribuição.

Não inclua código de terceiros sem verificar previamente os direitos de utilização e distribuição.

---

# 21. O que não deve ser enviado

Não envie Pull Requests contendo:

* o arquivo Excel comercial completo (`.xlsx`);
* credenciais;
* dados reais de clientes;
* dados financeiros;
* arquivos de configuração confidenciais;
* componentes comerciais não publicados;
* código sem origem ou licença verificável;
* alterações puramente cosméticas sem benefício relevante.

O `.xlsx` comercial deve permanecer fora do repositório público.

---

# 22. Processo de Contribuição

O fluxo recomendado é:

```text
Identificar problema
        │
        ▼
Verificar Issues existentes
        │
        ▼
Abrir Issue, quando necessário
        │
        ▼
Discutir a solução
        │
        ▼
Criar branch
        │
        ▼
Implementar alteração
        │
        ▼
Executar testes
        │
        ▼
Atualizar documentação
        │
        ▼
Criar Pull Request
        │
        ▼
Revisão
        │
        ▼
Merge
```

---

# 23. Branches

Para alterações independentes, utilize branches específicas.

Exemplos:

```text
feature/phoneval
feature/schema-improvements
fix/pipeline-parser
fix/null-validation
docs/performance
```

Evite realizar alterações diretamente na branch principal.

---

# 24. Antes de Abrir um Pull Request

Verifique:

* [ ] O problema está claramente identificado.
* [ ] A solução está dentro do escopo do framework.
* [ ] O código segue a arquitetura existente.
* [ ] Não existem dados confidenciais.
* [ ] Não existem componentes comerciais.
* [ ] Os testes relevantes foram executados.
* [ ] A documentação foi atualizada quando necessário.
* [ ] A alteração de performance foi medida, quando aplicável.
* [ ] O Pull Request possui escopo claro.
* [ ] Os commits possuem mensagens descritivas.

---

# 25. Dúvidas

Para dúvidas gerais sobre utilização do framework, consulte primeiro:

* [Getting Started](docs/getting-started.md)
* [Examples](docs/examples.md)
* [Troubleshooting](docs/troubleshooting.md)
* [FAQ](docs/faq.md), quando disponível.

Para problemas específicos de implementação, utilize uma Issue com informações suficientes para reprodução.

---

# 26. Resumo

As contribuições devem preservar os princípios fundamentais do projeto:

```text
Configuração
     │
     ▼
Schema
     │
     ▼
Pipeline
     │
     ▼
Camadas especializadas
     │
     ▼
Resultado
```

O objetivo das contribuições é melhorar o framework como uma plataforma reutilizável, e não adicionar lógica específica de um único projeto.

Contribuições devem priorizar:

* reutilização;
* baixo acoplamento;
* clareza;
* previsibilidade;
* desempenho mensurável;
* compatibilidade;
* documentação;
* manutenção de longo prazo.
