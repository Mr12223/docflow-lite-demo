[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Message,
    [string]$Branch,
    [switch]$NoPush,
    [switch]$ForceWithLease,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

function Write-Step([string]$Text) {
    Write-Host "==> $Text" -ForegroundColor Cyan
}

function Write-Ok([string]$Text) {
    Write-Host "OK  $Text" -ForegroundColor Green
}

function Fail([string]$Text) {
    Write-Host "ERR $Text" -ForegroundColor Red
    exit 1
}

try {
    git --version *> $null
} catch {
    Fail "Git was not found."
}

$repoRoot = (git rev-parse --show-toplevel 2>$null)
if (-not $repoRoot) {
    Fail "Current folder is not inside a Git repository."
}

Set-Location $repoRoot.Trim()

if (-not $Branch) {
    $Branch = (git branch --show-current 2>$null).Trim()
}
if (-not $Branch) {
    Fail "Cannot detect the current branch."
}

$originUrl = (git remote get-url origin 2>$null)
if (-not $originUrl) {
    Fail "Remote 'origin' is missing. Run: git remote add origin https://github.com/<user>/<repo>.git"
}

Write-Step "Repository: $repoRoot"
Write-Step "Branch: $Branch"
Write-Step "Remote: $originUrl"
Write-Step "Working tree:"
git status --short

Write-Step "Running git add ."
git add .

git diff --cached --quiet --exit-code
if ($LASTEXITCODE -eq 0) {
    Write-Ok "No staged changes to commit."
    exit 0
}

if ([string]::IsNullOrWhiteSpace($Message)) {
    $Message = Read-Host "Enter commit message (leave blank for auto message)"
}
if ([string]::IsNullOrWhiteSpace($Message)) {
    $Message = "chore: update $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

Write-Step "Commit message: $Message"

if ($DryRun) {
    Write-Host "DRY RUN: git commit -m ""$Message"""
    if (-not $NoPush) {
        if ($ForceWithLease) {
            Write-Host "DRY RUN: git push origin $Branch --force-with-lease"
        } else {
            Write-Host "DRY RUN: git push origin $Branch"
        }
    }
    exit 0
}

git commit -m $Message

if ($NoPush) {
    Write-Ok "Commit created. Push skipped."
    exit 0
}

Write-Step "Pushing to origin/$Branch"
if ($ForceWithLease) {
    git push origin $Branch --force-with-lease
} else {
    git push origin $Branch
}

Write-Ok "Commit and push completed. Render will auto-deploy after the new push is detected."
