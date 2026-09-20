$ErrorActionPreference = "Stop"

# CBC - Synchronisation automatique Google Apps Script -> GitHub
# Le dossier du script est automatiquement celui de ce fichier.
Set-Location -LiteralPath $PSScriptRoot

function Stop-WithMessage([string]$Message) {
    Write-Host ""
    Write-Host "ERREUR : $Message"
    exit 1
}

Write-Host "============================================================"
Write-Host " CBC : GOOGLE APPS SCRIPT -> GITHUB"
Write-Host "============================================================"
Write-Host ("Dossier : " + $PSScriptRoot)
Write-Host ""

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Stop-WithMessage "Git n'est pas installé ou n'est pas dans le PATH."
}

if (-not (Get-Command clasp -ErrorAction SilentlyContinue)) {
    Stop-WithMessage "clasp n'est pas installé. Installe Node.js puis lance : npm install -g @google/clasp"
}

# Ne mélange jamais des modifications locales manuelles avec une synchro automatique.
$dirty = git status --porcelain
if ($LASTEXITCODE -ne 0) {
    Stop-WithMessage "Impossible de lire l'état Git."
}
if ($dirty) {
    Stop-WithMessage "Le dépôt contient déjà des modifications locales. Fais d'abord un commit ou annule-les."
}

Write-Host "[1/5] Mise à jour locale depuis GitHub..."
git pull --ff-only origin main
if ($LASTEXITCODE -ne 0) {
    Stop-WithMessage "git pull a échoué. Vérifie la connexion ou l'état du dépôt."
}

Write-Host "[2/5] Récupération du projet Apps Script Google..."
clasp pull
if ($LASTEXITCODE -ne 0) {
    Stop-WithMessage "clasp pull a échoué. Vérifie 'clasp login' et l'ID du script."
}

Write-Host "[3/5] Détection des changements..."
git add -A -- .
if ($LASTEXITCODE -ne 0) {
    Stop-WithMessage "git add a échoué."
}

git diff --cached --quiet
$diffExit = $LASTEXITCODE

if ($diffExit -eq 0) {
    Write-Host ""
    Write-Host "Aucun changement : GitHub est déjà à jour."
    exit 0
}

if ($diffExit -ne 1) {
    Stop-WithMessage "Impossible de comparer les changements Git."
}

$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "[4/5] Création du commit..."
git commit -m "CBC auto-sync Google Apps Script - $stamp"
if ($LASTEXITCODE -ne 0) {
    Stop-WithMessage "git commit a échoué."
}

Write-Host "[5/5] Envoi vers GitHub..."
git push origin main
if ($LASTEXITCODE -ne 0) {
    Stop-WithMessage "git push a échoué."
}

Write-Host ""
Write-Host "OK : CBC Google Apps Script est maintenant sauvegardé sur GitHub."
