param(
    [string]$BumpType = "patch"
)

$tags = git tag -l "v*.*.*" | Sort-Object { [version]($_ -replace '^v', '') } -Descending
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

Write-Host "Last version: $lastTag"
Write-Host "New version: v$newVersion"

"version: $newVersion`n" | Out-File -FilePath "version.yaml" -Encoding UTF8 -NoNewline

git add version.yaml
git commit -m "Release v$newVersion"
git tag "v$newVersion"
git push origin master
git push github master
git push origin "v$newVersion"
git push github "v$newVersion"

Write-Host "Released v$newVersion"
