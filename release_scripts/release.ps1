param(
    [string]$BumpType = "patch",
    [switch]$Staging
)

$ErrorActionPreference = "Stop"

$status = git status --porcelain
if ($status) {
    Write-Host "ERROR: There are uncommitted changes:" -ForegroundColor Red
    Write-Host $status -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please commit or stash changes before releasing." -ForegroundColor Red
    exit 1
}

$tags = git tag -l "v*.*.*" | Sort-Object {
    $v = $_ -replace '^v', ''
    $parts = $v.Split('.')
    [int]$parts[0] * 10000 + [int]$parts[1] * 100 + [int]$parts[2]
} -Descending
$lastTag = $tags | Select-Object -First 1

if ($lastTag) {
    $version = $lastTag -replace '^v', '' -replace '-staging$', ''
    $parts = $version.Split('.')
    $major = [int]$parts[0]
    $minor = [int]$parts[1]
    $patch = [int]$parts[2]

    if ($Staging) {
        # Staging: use current version with -staging suffix
        $newVersion = "$major.$minor.$patch-staging"
    } else {
        # Production: bump version
        switch ($BumpType) {
            "major" { $major++; $minor = 0; $patch = 0 }
            "minor" { $minor++; $patch = 0 }
            "patch" { $patch++ }
        }
        $newVersion = "$major.$minor.$patch"
    }
} else {
    if ($Staging) {
        $newVersion = "1.0.0-staging"
    } else {
        $newVersion = "1.0.0"
    }
}

Write-Host "Last version: $lastTag" -ForegroundColor Cyan
Write-Host "New version:  v$newVersion" -ForegroundColor Green

# Get current branch
$currentBranch = git rev-parse --abbrev-ref HEAD

if ($Staging) {
    Write-Host "Branch: $currentBranch (staging from current branch)" -ForegroundColor Yellow
} else {
    if ($currentBranch -ne "master") {
        Write-Host "ERROR: Production release must be from master branch!" -ForegroundColor Red
        Write-Host "Current branch: $currentBranch" -ForegroundColor Yellow
        Write-Host "Switch to master or use -Staging for test builds" -ForegroundColor Yellow
        exit 1
    }
}

# Get repo root directory
$repoRoot = git rev-parse --show-toplevel
$versionFile = Join-Path $repoRoot "version.yaml"

"version: $newVersion`n" | Out-File -FilePath $versionFile -Encoding UTF8 -NoNewline

git add $versionFile
git commit -m "Release v$newVersion"
git tag "v$newVersion"

Write-Host "Pushing to origin..." -ForegroundColor Cyan
git push origin $currentBranch
git push origin "v$newVersion"

Write-Host "Pushing tag to github..." -ForegroundColor Cyan
if (-not $Staging) {
    git push github master
}
git push github "v$newVersion"

Write-Host ""
if ($Staging) {
    Write-Host "Staging release v$newVersion" -ForegroundColor Yellow
    Write-Host "(Build without server upload)" -ForegroundColor Yellow
} else {
    Write-Host "Released v$newVersion" -ForegroundColor Green
}
Write-Host ""
Write-Host "Build status: https://github.com/fesworkscience/pyrevit_rocket/actions" -ForegroundColor Cyan
