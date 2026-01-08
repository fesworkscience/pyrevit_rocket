param([Parameter(Mandatory=$true)][string]$ExtensionPath)

$ErrorActionPreference = "Stop"
$logFile = "$env:TEMP\CPSK_Install.log"

function Write-Log { param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append -Encoding UTF8
    Write-Host $Message
}

function Get-IniContent { param([string]$Path)
    $ini = @{}
    if (-not (Test-Path $Path)) { return $ini }
    $section = "NO_SECTION"
    $ini[$section] = @{}
    Get-Content -Path $Path -Encoding UTF8 | ForEach-Object {
        $line = $_.Trim()
        if ($line -eq "" -or $line.StartsWith(";") -or $line.StartsWith("#")) { return }
        if ($line -match '^\[(.+)\]$') { $section = $matches[1]; if (-not $ini.ContainsKey($section)) { $ini[$section] = @{} }; return }
        if ($line -match '^([^=]+)=(.*)$') { $ini[$section][$matches[1].Trim()] = $matches[2].Trim() }
    }
    return $ini
}

function Set-IniContent { param([string]$Path, [hashtable]$Content)
    $lines = @()
    foreach ($section in $Content.Keys | Sort-Object) {
        if ($section -ne "NO_SECTION") { $lines += "[$section]" }
        foreach ($key in $Content[$section].Keys | Sort-Object) { $lines += "$key = $($Content[$section][$key])" }
        $lines += ""
    }
    $lines | Out-File -FilePath $Path -Encoding UTF8 -Force
}

Write-Log "=== CPSK Tools: Register Extension ==="
Write-Log "Path: $ExtensionPath"

$pyrevitConfigDir = "$env:APPDATA\pyRevit"
$pyrevitConfigFile = "$pyrevitConfigDir\pyRevit_config.ini"
if (-not (Test-Path $pyrevitConfigDir)) { New-Item -ItemType Directory -Path $pyrevitConfigDir -Force | Out-Null }

$extensionParentPath = Split-Path -Parent $ExtensionPath
$config = Get-IniContent -Path $pyrevitConfigFile
if (-not $config.ContainsKey("core")) { $config["core"] = @{} }

$currentExtensions = @()
if ($config["core"].ContainsKey("userextensions")) {
    $extensionsStr = $config["core"]["userextensions"]
    if ($extensionsStr -match '^\[(.+)\]$') {
        $currentExtensions = $matches[1] -split ',' | ForEach-Object { $_.Trim().Trim('"').Trim("'") } | Where-Object { $_ -ne "" }
    }
}

$normalizedNewPath = $extensionParentPath.Replace('\', '/').TrimEnd('/')
$alreadyRegistered = $false
foreach ($existingPath in $currentExtensions) {
    if ($existingPath.Replace('\', '/').TrimEnd('/') -eq $normalizedNewPath) { $alreadyRegistered = $true; break }
}

if (-not $alreadyRegistered) {
    $currentExtensions += $extensionParentPath
    $pathsFormatted = $currentExtensions | ForEach-Object { '"' + $_.Replace('\', '/') + '"' }
    $config["core"]["userextensions"] = "[" + ($pathsFormatted -join ", ") + "]"
    Set-IniContent -Path $pyrevitConfigFile -Content $config
    Write-Log "Extension registered"
} else {
    Write-Log "Extension already registered"
}

Write-Log "=== Registration complete ==="
exit 0
