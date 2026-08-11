#!/usr/bin/env bash

set -euo pipefail

REQUIRED_NODE_MAJOR=26
REQUIRED_COREPACK_VERSION="0.35.0"
REQUIRED_YARN_VERSION="4.18.0"

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

cd "$PROJECT_ROOT"

step() {
    echo
    echo "==> $1"
}

ok() {
    echo "    [OK] $1"
}

warn() {
    echo "    [!] $1"
}

fail() {
    echo
    echo "[ERRO] $1"
    exit 1
}

echo
echo "Excel Data Framework DBB - Setup"
echo "Projeto: $PROJECT_ROOT"

# ---------------------------------------------------------------------------
# Git
# ---------------------------------------------------------------------------

step "Verificando Git"

if ! command -v git >/dev/null 2>&1; then
    fail "Git não está instalado ou não está disponível no PATH."
fi

ok "$(git --version)"

# ---------------------------------------------------------------------------
# Node.js
# ---------------------------------------------------------------------------

step "Verificando Node.js"

if ! command -v node >/dev/null 2>&1; then
    fail "Node.js não está instalado ou não está disponível no PATH."
fi

NODE_VERSION="$(node --version)"
NODE_MAJOR="$(printf '%s' "$NODE_VERSION" | sed -E 's/^v([0-9]+).*/\1/')"

if [ "$NODE_MAJOR" -ne "$REQUIRED_NODE_MAJOR" ]; then
    fail "O projeto requer Node.js ${REQUIRED_NODE_MAJOR}.x. Versão encontrada: $NODE_VERSION"
fi

ok "Node.js $NODE_VERSION"

# ---------------------------------------------------------------------------
# package.json
# ---------------------------------------------------------------------------

step "Verificando configuração do projeto"

if [ ! -f "package.json" ]; then
    fail "package.json não encontrado."
fi

PACKAGE_MANAGER="$(node -p "require('./package.json').packageManager")"

if [ "$PACKAGE_MANAGER" != "yarn@$REQUIRED_YARN_VERSION" ]; then
    fail "O projeto requer yarn@$REQUIRED_YARN_VERSION. Encontrado: $PACKAGE_MANAGER"
fi

ok "packageManager: $PACKAGE_MANAGER"

# ---------------------------------------------------------------------------
# .yarnrc.yml
# ---------------------------------------------------------------------------

if [ ! -f ".yarnrc.yml" ]; then
    fail ".yarnrc.yml não encontrado."
fi

if ! grep -Eq '^[[:space:]]*nodeLinker:[[:space:]]*pnp[[:space:]]*$' .yarnrc.yml; then
    fail ".yarnrc.yml não está configurado com nodeLinker: pnp."
fi

ok "Yarn Plug'n'Play configurado"

# ---------------------------------------------------------------------------
# Corepack
# ---------------------------------------------------------------------------

step "Verificando Corepack"

if ! command -v corepack >/dev/null 2>&1; then

    warn "Corepack não encontrado."

    echo "    Instalando Corepack $REQUIRED_COREPACK_VERSION..."

    npm install --global "corepack@$REQUIRED_COREPACK_VERSION"

    hash -r

    if ! command -v corepack >/dev/null 2>&1; then
        fail "Corepack foi instalado, mas não está disponível no PATH."
    fi
fi

COREPACK_VERSION="$(corepack --version)"

ok "Corepack $COREPACK_VERSION"

# ---------------------------------------------------------------------------
# Corepack enable
# ---------------------------------------------------------------------------

step "Habilitando Corepack"

if ! corepack enable; then
    fail "Não foi possível executar 'corepack enable'."
fi

ok "Corepack habilitado"

# ---------------------------------------------------------------------------
# Yarn
# ---------------------------------------------------------------------------

step "Verificando Yarn"

YARN_VERSION="$(corepack yarn --version)"

if [ "$YARN_VERSION" != "$REQUIRED_YARN_VERSION" ]; then
    fail "O projeto requer Yarn $REQUIRED_YARN_VERSION. Versão encontrada: $YARN_VERSION"
fi

ok "Yarn $YARN_VERSION"

# ---------------------------------------------------------------------------
# Yarn PnP
# ---------------------------------------------------------------------------

step "Verificando Yarn Plug'n'Play"

if [ -f ".pnp.cjs" ]; then
    ok ".pnp.cjs"
else
    warn ".pnp.cjs ainda não existe. Será criado pelo Yarn."
fi

if [ -f ".pnp.loader.mjs" ]; then
    ok ".pnp.loader.mjs"
else
    warn ".pnp.loader.mjs ainda não existe. Será criado pelo Yarn."
fi

# ---------------------------------------------------------------------------
# Dependências
# ---------------------------------------------------------------------------

step "Instalando dependências"

if ! corepack yarn install --immutable; then
    fail "Falha durante 'yarn install --immutable'."
fi

ok "Dependências instaladas"

# ---------------------------------------------------------------------------
# PnP
# ---------------------------------------------------------------------------

step "Validando artefatos PnP"

if [ ! -f ".pnp.cjs" ]; then
    fail ".pnp.cjs não foi gerado."
fi

if [ ! -f ".pnp.loader.mjs" ]; then
    fail ".pnp.loader.mjs não foi gerado."
fi

ok ".pnp.cjs"
ok ".pnp.loader.mjs"

# ---------------------------------------------------------------------------
# node_modules
# ---------------------------------------------------------------------------

step "Verificando node_modules"

if [ -d "node_modules" ]; then
    warn "node_modules existe. Este projeto utiliza Yarn PnP."
else
    ok "node_modules não existe"
fi

# ---------------------------------------------------------------------------
# Semantic Release
# ---------------------------------------------------------------------------

step "Verificando Semantic Release"

SEMANTIC_RELEASE_VERSION="$(
    corepack yarn node -p \
    "require('semantic-release/package.json').version"
)"

if [ -z "$SEMANTIC_RELEASE_VERSION" ]; then
    fail "Não foi possível determinar a versão do Semantic Release."
fi

ok "semantic-release $SEMANTIC_RELEASE_VERSION"

# ---------------------------------------------------------------------------
# Plugins
# ---------------------------------------------------------------------------

step "Verificando plugins do Semantic Release"

for PLUGIN in \
    "@semantic-release/changelog" \
    "@semantic-release/git" \
    "@semantic-release/github"
do

    if corepack yarn node -e "require('$PLUGIN')" >/dev/null 2>&1; then
        ok "$PLUGIN"
    else
        fail "Plugin não pôde ser carregado: $PLUGIN"
    fi

done

# ---------------------------------------------------------------------------
# Git status
# ---------------------------------------------------------------------------

step "Verificação final"

if [ -n "$(git status --short)" ]; then
    warn "O working tree possui alterações."
    git status --short
else
    ok "Working tree limpo"
fi

# ---------------------------------------------------------------------------
# Final
# ---------------------------------------------------------------------------

echo
echo "=================================================="
echo " Ambiente configurado com sucesso."
echo "=================================================="
echo
echo "Node.js          : $NODE_VERSION"
echo "Corepack         : $COREPACK_VERSION"
echo "Yarn             : $YARN_VERSION"
echo "PnP              : habilitado"
echo "Semantic Release : $SEMANTIC_RELEASE_VERSION"
echo
echo "Para validar o pipeline de release:"
echo
echo "  corepack yarn semantic-release --dry-run"
echo