param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Output
)

function Get-SafeId($path) {
    $hash = [System.BitConverter]::ToString([System.Security.Cryptography.MD5]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes($path))).Replace("-","").Substring(0,16)
    return $hash
}

$tree = @{}
$components = @()

Get-ChildItem -Path $Source -Recurse -File | ForEach-Object {
    $relPath = $_.FullName.Substring((Get-Item $Source).FullName.Length + 1)
    $relDir = Split-Path -Parent $relPath

    $fileId = "F_" + (Get-SafeId $relPath)
    $compId = "C_" + (Get-SafeId $relPath)
    $srcPath = "`$(var.ExtensionSource)\$relPath"

    if ($relDir) {
        $dirId = "D_" + (Get-SafeId $relDir)
        $parts = $relDir.Split('\')
        $current = $tree
        $pathSoFar = ""
        foreach ($part in $parts) {
            $pathSoFar = if ($pathSoFar) { "$pathSoFar\$part" } else { $part }
            if (-not $current.ContainsKey($part)) {
                $current[$part] = @{
                    "_id" = "D_" + (Get-SafeId $pathSoFar)
                    "_children" = @{}
                }
            }
            $current = $current[$part]["_children"]
        }
    } else {
        $dirId = "EXTENSIONFOLDER"
    }

    $components += @{
        CompId = $compId
        FileId = $fileId
        DirId = $dirId
        Source = $srcPath
    }
}

function Write-DirectoryTree($node, $indent) {
    $result = ""
    foreach ($key in $node.Keys | Sort-Object) {
        if ($key -notlike "_*") {
            $dirId = $node[$key]["_id"]
            $children = $node[$key]["_children"]
            $result += "$indent<Directory Id=`"$dirId`" Name=`"$key`">`n"
            $result += Write-DirectoryTree $children ("$indent  ")
            $result += "$indent</Directory>`n"
        }
    }
    return $result
}

$xml = '<?xml version="1.0" encoding="UTF-8"?>' + "`n"
$xml += '<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs">' + "`n"
$xml += '  <Fragment>' + "`n"
$xml += '    <DirectoryRef Id="EXTENSIONFOLDER">' + "`n"
$xml += Write-DirectoryTree $tree "      "
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
Write-Host "Generated $Output with $($components.Count) files"
