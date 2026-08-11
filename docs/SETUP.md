# Configuração do ambiente em outro computador

Este documento descreve como clonar e preparar o **Excel Data Framework DBB** em outro computador Windows.

## 1. Pré-requisitos

Instale:

- Git
- Node.js 26.x

Verifique:

```powershell
git --version
node --version
```

O ambiente utilizado no projeto foi validado com Node.js 26.7.0.

## 2. Clonar o repositório

```powershell
git clone https://github.com/danielbighelini/Excel-Data-Framework-DBB.git
cd Excel-Data-Framework-DBB
git status
```

## 3. Instalar e habilitar Corepack

O projeto fixa Yarn 4.18.0 no `package.json`:

```json
"packageManager": "yarn@4.18.0"
```

No Node.js 26, instale o Corepack:

```powershell
npm install -g corepack@0.35.0
```

Depois:

```powershell
corepack --version
corepack enable
corepack yarn --version
```

A versão esperada do Yarn é:

```text
4.18.0
```

### Se `corepack enable` retornar EPERM

Abra o PowerShell como Administrador e execute:

```powershell
corepack enable
```

Depois volte ao terminal normal.

## 4. Instalar as dependências

O projeto usa **Yarn Plug'n'Play (PnP)**. Não use `npm install`, `npm ci` nem crie `node_modules`.

Execute:

```powershell
corepack yarn install --immutable
```

A configuração PnP está em `.yarnrc.yml`:

```yaml
nodeLinker: pnp
```

O projeto utiliza:

```text
.pnp.cjs
.pnp.loader.mjs
yarn.lock
```

## 5. Validar o ambiente

```powershell
corepack yarn --version
corepack yarn node --version
corepack yarn node -p "require('semantic-release/package.json').version"
```

Atualmente, o Semantic Release esperado é:

```text
25.0.9
```

Também é possível validar a configuração sem criar uma release:

```powershell
corepack yarn semantic-release --dry-run
```

O teste deve carregar:

```text
@semantic-release/changelog
@semantic-release/git
@semantic-release/github
```

> `--dry-run` não cria `CHANGELOG.md`, tag ou release. Ele apenas valida o processo.

## 6. Configurar `.env`

O arquivo `.env` não deve ser versionado.

Se o repositório disponibilizar `.env.example`:

```powershell
Copy-Item .env.example .env
```

Depois preencha os valores necessários.

Não versione tokens, PATs, senhas ou outras credenciais.

Se não houver `.env.example`, consulte a documentação do projeto antes de criar o `.env`; não invente variáveis.

## 7. Configurar Git

Se necessário:

```powershell
git config --global user.name "Seu Nome"
git config --global user.email "seu-email@example.com"
```

Confirme o remoto:

```powershell
git remote -v
```

O remoto esperado é:

```text
https://github.com/danielbighelini/Excel-Data-Framework-DBB.git
```

## 8. Desenvolvimento diário

O fluxo normal é:

```powershell
git pull
```

Faça as alterações, teste e depois:

```powershell
git status
git add .
git commit -m "tipo: descrição"
git push
```

## 9. Conventional Commits

O Semantic Release determina a próxima versão a partir das mensagens dos commits.

### PATCH

Correção de comportamento existente:

```text
fix: corrige validação de CPF
```

Exemplo:

```text
1.3.0 -> 1.3.1
```

### MINOR

Nova funcionalidade compatível:

```text
feat: adiciona nova função de normalização
```

Exemplo:

```text
1.3.0 -> 1.4.0
```

### MAJOR

Mudança incompatível:

```text
feat!: altera assinatura da função
```

ou:

```text
feat: altera assinatura da função

BREAKING CHANGE: o parâmetro X agora é obrigatório.
```

Exemplo:

```text
1.3.0 -> 2.0.0
```

### Commits que normalmente não geram release

```text
chore: ...
docs: ...
refactor: ...
test: ...
style: ...
```

## 10. CHANGELOG e releases

Não crie nem edite `CHANGELOG.md` manualmente.

Ao fazer `push` para `main`, o GitHub Actions executa o Semantic Release:

```text
git push
    ↓
GitHub Actions
    ↓
Node.js 26
    ↓
Corepack
    ↓
Yarn 4.18.0
    ↓
Yarn PnP
    ↓
Semantic Release
    ↓
determina a versão
    ↓
@semantic-release/changelog
    ↓
gera/atualiza CHANGELOG.md
    ↓
@semantic-release/git
    ↓
commita CHANGELOG.md
    ↓
cria tag
    ↓
cria GitHub Release
```

O `CHANGELOG.md` é gerado oficialmente no runner do GitHub Actions. Para recebê-lo no computador local depois da release:

```powershell
git pull
```

O fluxo normal de desenvolvimento não exige um GitHub PAT local para o release; o workflow usa o `GITHUB_TOKEN` do GitHub Actions.

## 11. Yarn Plug'n'Play

Este projeto deliberadamente não usa:

```text
node_modules/
```

Não execute:

```powershell
npm install
npm ci
```

Use:

```powershell
corepack yarn install --immutable
```

O cache do Yarn fica fora do repositório, enquanto os artefatos PnP necessários ao projeto estão versionados.

## 12. Arquivos que não devem ser versionados

Entre os principais itens ignorados pelo `.gitignore` estão:

```text
node_modules/
.yarn/
.venv/
.env
backups/
assets/videos/
examples/data/
```

Respeite também as demais regras existentes no `.gitignore`.

## 13. Problemas comuns

### `corepack` não encontrado

```powershell
npm install -g corepack@0.35.0
```

Depois:

```powershell
corepack enable
```

### `corepack enable` retorna `EPERM`

Abra o PowerShell como Administrador:

```powershell
corepack enable
```

### Yarn está em versão diferente

O projeto exige:

```text
4.18.0
```

Confirme:

```powershell
corepack yarn --version
```

Não substitua a versão definida pelo projeto.

### `node_modules` apareceu

Não é necessário. O projeto usa PnP.

Remova o diretório e execute:

```powershell
corepack yarn install --immutable
```

### Semantic Release não encontra dependências

Execute:

```powershell
corepack yarn install --immutable
```

Depois:

```powershell
corepack yarn semantic-release --dry-run
```

## 14. Checklist de instalação

- [ ] Instalar Git
- [ ] Instalar Node.js 26.x
- [ ] Clonar o repositório
- [ ] Instalar `corepack@0.35.0`
- [ ] Executar `corepack enable`
- [ ] Confirmar Yarn `4.18.0`
- [ ] Executar `corepack yarn install --immutable`
- [ ] Confirmar que não existe `node_modules`
- [ ] Configurar `.env`, se necessário
- [ ] Configurar identidade do Git
- [ ] Confirmar `git remote -v`
- [ ] Executar `corepack yarn semantic-release --dry-run`
- [ ] Usar Conventional Commits
- [ ] Fazer `git push` para `main`
- [ ] Deixar o GitHub Actions executar a release

## 15. Regra operacional

Depois da configuração inicial, o fluxo deve ser:

```powershell
git pull
corepack yarn <comando>
git add .
git commit -m "tipo: descrição"
git push
```

Não altere manualmente:

- versões do projeto;
- `CHANGELOG.md`;
- tags de release.

O Semantic Release controla esses itens automaticamente a partir dos commits.
