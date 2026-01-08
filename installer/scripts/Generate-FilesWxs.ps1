param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Output
)

function Get-SafeId($path) {
    $hash = [System.BitConverter]::ToString([System.Security.Cryptography.MD5]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes($path))).Replace("-","").Substring(0,16)
    return $hash
}

$directories = @{}
$components = @()

Get-ChildItem -Path $Source -Recurse -File | ForEach-Object {
    $relPath = $_.FullName.Substring((Get-Item $Source).FullName.Length + 1)
    $relDir = Split-Path -Parent $relPath

    if ($relDir -and -not $directories.ContainsKey($relDir)) {
        $directories[$relDir] = "D_" + (Get-SafeId $relDir)
    }

    $dirId = if ($relDir) { $directories[$relDir] } else { "EXTENSIONFOLDER" }
    $fileId = "F_" + (Get-SafeId $relPath)
    $compId = "C_" + (Get-SafeId $relPath)
    $srcPath = "`$(var.ExtensionSource)\$relPath"

    $components += @{
        CompId = $compId
        FileId = $fileId
        DirId = $dirId
        Source = $srcPath
    }
}

$xml = '<?xml version="1.0" encoding="UTF-8"?>' + "`n"
$xml += '<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs">' + "`n"
$xml += '  <Fragment>' + "`n"
$xml += '    <DirectoryRef Id="EXTENSIONFOLDER">' + "`n"

$sortedDirs = $directories.GetEnumerator() | Sort-Object { $_.Key.Split('\').Count }, { $_.Key }
$createdDirs = @{}

foreach ($dir in $sortedDirs) {
    $parts = $dir.Key.Split('\')
    $currentPath = ""

    for ($i = 0; $i -lt $parts.Count; $i++) {
        $part = $parts[$i]
        $currentPath = if ($currentPath) { "$currentPath\$part" } else { $part }

        if (-not $createdDirs.ContainsKey($currentPath)) {
            $dirId = if ($i -eq $parts.Count - 1) { $dir.Value } else { "D_" + (Get-SafeId $currentPath) }
            $indent = "      " + ("  " * $i)
            $xml += "$indent<Directory Id=`"$dirId`" Name=`"$part`">`n"
            $createdDirs[$currentPath] = @{ Id = $dirId; Depth = $i }
        }
    }
}

$maxDepth = ($createdDirs.Values | Measure-Object -Property Depth -Maximum).Maximum
if ($null -eq $maxDepth) { $maxDepth = -1 }
for ($d = $maxDepth; $d -ge 0; $d--) {
    $dirsAtDepth = $createdDirs.GetEnumerator() | Where-Object { $_.Value.Depth -eq $d } | Sort-Object { $_.Key } -Descending
    foreach ($dir in $dirsAtDepth) {
        $indent = "      " + ("  " * $dir.Value.Depth)
        $xml += "$indent</Directory>`n"
    }
}

$xml += '    </DirectoryRef>' + "`n`n"
$xml += '    <ComponentGroup Id="ExtensionFiles">' + "`n"

foreach ($comp in $components) {
    $xml += '      <Component Id="' + $comp.CompId + '" Directory="' + $comp.DirId + '" Guid="*">' + "`n"
    $xml += '        <File Id="' + $comp.FileId + '" Source="' + $comp.Source + '" KeyPath="yes" />' + "`n"
    $xml += '      </Component>' + "`n"
}

$xml += '    </ComponentGroup>' + "`n"
$xml += '  </Fragment>' + "`n"
$xml += '</Wix>'

$xml | Out-File -FilePath $Output -Encoding UTF8
Write-Host "Generated $Output with $($components.Count) files and $($directories.Count) directories"
