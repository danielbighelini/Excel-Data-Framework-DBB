#Requires -Version 5.1

$ErrorActionPreference = "Stop"

$RequiredNodeMajor = 26
$RequiredCorepackVersion = "0.35.0"
$RequiredYarnVersion = "4.18.0"

function Write-Step {
    param([string]$Message)
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Message)
    Write-Host "    [OK] $Message" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Message)
    Write-Host "    [!] $Message" -ForegroundColor Yellow
}

function Fail {
    param([string]$Message)
    Write-Host "`n[ERRO] $Message" -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------------------
# Local do projeto
# ---------------------------------------------------------------------------

$ProjectRoot = Split-Path -Parent $PSScriptRoot

Set-Location $ProjectRoot

Write-Host ""
Write-Host "Excel Data Framework DBB - Setup" -ForegroundColor White
Write-Host "Projeto: $ProjectRoot" -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# Verificar Git
# ---------------------------------------------------------------------------

Write-Step "Verificando Git"

$Git = Get-Command git -ErrorAction SilentlyContinue

if (-not $Git) {
    Fail "Git não está instalado ou não está disponível no PATH."
}

$GitVersion = git --version

Write-Ok $GitVersion

# ---------------------------------------------------------------------------
# Verificar Node.js
# ---------------------------------------------------------------------------

Write-Step "Verificando Node.js"

$Node = Get-Command node -ErrorAction SilentlyContinue

if (-not $Node) {
    Fail "Node.js não está instalado ou não está disponível no PATH."
}

$NodeVersion = node --version

if ($NodeVersion -notmatch '^v(\d+)\.') {
    Fail "Não foi possível determinar a versão do Node.js."
}

$NodeMajor = [int]$Matches[1]

if ($NodeMajor -ne $RequiredNodeMajor) {
    Fail "O projeto requer Node.js $RequiredNodeMajor.x. Versão encontrada: $NodeVersion"
}

Write-Ok "Node.js $NodeVersion"

# ---------------------------------------------------------------------------
# Verificar package.json
# ---------------------------------------------------------------------------

Write-Step "Verificando configuração do projeto"

$PackageJsonPath = Join-Path $ProjectRoot "package.json"

if (-not (Test-Path $PackageJsonPath)) {
    Fail "package.json não encontrado."
}

$PackageJson = Get-Content $PackageJsonPath -Raw | ConvertFrom-Json

if (-not $PackageJson.packageManager) {
    Fail "package.json não possui a propriedade packageManager."
}

if ($PackageJson.packageManager -ne "yarn@$RequiredYarnVersion") {
    Fail "O projeto requer $($PackageJson.packageManager), mas o setup espera yarn@$RequiredYarnVersion."
}

Write-Ok "packageManager: $($PackageJson.packageManager)"

# ---------------------------------------------------------------------------
# Verificar .yarnrc.yml
# ---------------------------------------------------------------------------

$YarnRcPath = Join-Path $ProjectRoot ".yarnrc.yml"

if (-not (Test-Path $YarnRcPath)) {
    Fail ".yarnrc.yml não encontrado."
}

$YarnRc = Get-Content $YarnRcPath -Raw

if ($YarnRc -notmatch '(?m)^\s*nodeLinker:\s*pnp\s*$') {
    Fail ".yarnrc.yml não está configurado com nodeLinker: pnp."
}

Write-Ok "Yarn Plug'n'Play configurado"

# ---------------------------------------------------------------------------
# Instalar/verificar Corepack
# ---------------------------------------------------------------------------

Write-Step "Verificando Corepack"

$Corepack = Get-Command corepack -ErrorAction SilentlyContinue

if (-not $Corepack) {
    Write-Warn "Corepack não encontrado. Instalando Corepack $RequiredCorepackVersion..."

    npm install --global "corepack@$RequiredCorepackVersion"

    $Corepack = Get-Command corepack -ErrorAction SilentlyContinue

    if (-not $Corepack) {
        Fail "Corepack foi instalado, mas não está disponível no PATH. Reinicie o PowerShell e execute o setup novamente."
    }
}

$CorepackVersion = corepack --version

Write-Ok "Corepack $CorepackVersion"

# ---------------------------------------------------------------------------
# Habilitar Corepack
# ---------------------------------------------------------------------------

Write-Step "Habilitando Corepack"

try {
    corepack enable
}
catch {
    Fail "Não foi possível executar 'corepack enable'. Tente executar o PowerShell como Administrador."
}

Write-Ok "Corepack habilitado"

# ---------------------------------------------------------------------------
# Verificar Yarn
# ---------------------------------------------------------------------------

Write-Step "Verificando Yarn"

$YarnVersion = corepack yarn --version

if ($YarnVersion.Trim() -ne $RequiredYarnVersion) {
    Fail "O projeto requer Yarn $RequiredYarnVersion. Versão encontrada: $YarnVersion"
}

Write-Ok "Yarn $YarnVersion"

# ---------------------------------------------------------------------------
# Verificar PnP
# ---------------------------------------------------------------------------

Write-Step "Verificando Yarn Plug'n'Play"

$PnpPath = Join-Path $ProjectRoot ".pnp.cjs"
$PnpLoaderPath = Join-Path $ProjectRoot ".pnp.loader.mjs"

if (-not (Test-Path $PnpPath)) {
    Write-Warn ".pnp.cjs não encontrado. Será gerado pelo Yarn."
}

if (-not (Test-Path $PnpLoaderPath)) {
    Write-Warn ".pnp.loader.mjs não encontrado. Será gerado pelo Yarn."
}

# ---------------------------------------------------------------------------
# Instalar dependências
# ---------------------------------------------------------------------------

Write-Step "Instalando dependências"

corepack yarn install --immutable

if ($LASTEXITCODE -ne 0) {
    Fail "Falha durante 'yarn install --immutable'."
}

Write-Ok "Dependências instaladas"

# ---------------------------------------------------------------------------
# Verificar PnP novamente
# ---------------------------------------------------------------------------

Write-Step "Validando artefatos PnP"

if (-not (Test-Path $PnpPath)) {
    Fail ".pnp.cjs não foi gerado."
}

if (-not (Test-Path $PnpLoaderPath)) {
    Fail ".pnp.loader.mjs não foi gerado."
}

Write-Ok ".pnp.cjs"
Write-Ok ".pnp.loader.mjs"

# ---------------------------------------------------------------------------
# Verificar node_modules
# ---------------------------------------------------------------------------

Write-Step "Verificando node_modules"

$NodeModulesPath = Join-Path $ProjectRoot "node_modules"

if (Test-Path $NodeModulesPath) {
    Write-Warn "node_modules existe. O projeto utiliza Yarn PnP e não precisa desse diretório."
}
else {
    Write-Ok "node_modules não existe"
}

# ---------------------------------------------------------------------------
# Verificar Semantic Release
# ---------------------------------------------------------------------------

Write-Step "Verificando Semantic Release"

try {
    $SemanticReleaseVersion = corepack yarn node -p "require('semantic-release/package.json').version"
}
catch {
    Fail "Não foi possível carregar semantic-release através do Yarn PnP."
}

Write-Ok "semantic-release $SemanticReleaseVersion"

# ---------------------------------------------------------------------------
# Validando Semantic Release
# ---------------------------------------------------------------------------

Write-Step "Validando Semantic Release"

try {
    corepack yarn semantic-release --dry-run
}
catch {
    Fail "Falha ao executar o Semantic Release."
}

if ($LASTEXITCODE -ne 0) {
    Fail "Semantic Release retornou código de erro."
}

Write-Ok "Semantic Release executado com sucesso"

# ---------------------------------------------------------------------------
# Status final
# ---------------------------------------------------------------------------

Write-Step "Verificação final"

$GitStatus = git status --short

if ($GitStatus) {
    Write-Warn "O working tree possui alterações:"
    Write-Host $GitStatus
}
else {
    Write-Ok "Working tree limpo"
}

# ---------------------------------------------------------------------------
# Final
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "==================================================" -ForegroundColor Green
Write-Host " Ambiente configurado com sucesso." -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Node.js : $NodeVersion"
Write-Host "Corepack: $CorepackVersion"
Write-Host "Yarn    : $YarnVersion"
Write-Host "PnP     : habilitado"
Write-Host "Semantic Release: $SemanticReleaseVersion"
Write-Host ""
Write-Host "Para validar o Semantic Release sem criar uma release:"
Write-Host "  corepack yarn semantic-release --dry-run" -ForegroundColor Cyan
Write-Host ""