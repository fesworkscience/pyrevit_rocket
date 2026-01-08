param(
    [string]$BumpType = "patch"
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
    $version = $lastTag -replace '^v', ''
    $parts = $version.Split('.')
    $major = [int]$parts[0]
    $minor = [int]$parts[1]
    $patch = [int]$parts[2]

    switch ($BumpType) {
        "major" { $major++; $minor = 0; $patch = 0 }
        "minor" { $minor++; $patch = 0 }
        "patch" { $patch++ }
    }
    $newVersion = "$major.$minor.$patch"
} else {
    $newVersion = "1.0.0"
}

Write-Host "Last version: $lastTag" -ForegroundColor Cyan
Write-Host "New version:  v$newVersion" -ForegroundColor Green

"version: $newVersion`n" | Out-File -FilePath "version.yaml" -Encoding UTF8 -NoNewline

git add version.yaml
git commit -m "Release v$newVersion"
git tag "v$newVersion"

Write-Host "Pushing to origin..." -ForegroundColor Cyan
git push origin master
git push origin "v$newVersion"

Write-Host "Pushing to github..." -ForegroundColor Cyan
git push github master
git push github "v$newVersion"

Write-Host ""
Write-Host "Released v$newVersion" -ForegroundColor Green
