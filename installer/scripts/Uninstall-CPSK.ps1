param([Parameter(Mandatory=$true)][string]$ExtensionPath)

$ErrorActionPreference = "SilentlyContinue"
$logFile = "$env:TEMP\CPSK_Uninstall.log"

function Write-Log { param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append -Encoding UTF8
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

Write-Log "=== CPSK Tools: Uninstall ==="

$pyrevitConfigFile = "$env:APPDATA\pyRevit\pyRevit_config.ini"
if (Test-Path $pyrevitConfigFile) {
    $config = Get-IniContent -Path $pyrevitConfigFile
    $extensionParentPath = Split-Path -Parent $ExtensionPath
    $normalizedPath = $extensionParentPath.Replace('\', '/').TrimEnd('/')

    if ($config.ContainsKey("core") -and $config["core"].ContainsKey("userextensions")) {
        $extensionsStr = $config["core"]["userextensions"]
        if ($extensionsStr -match '^\[(.+)\]$') {
            $currentExtensions = $matches[1] -split ',' | ForEach-Object { $_.Trim().Trim('"').Trim("'") } | Where-Object { $_ -ne "" }
            $newExtensions = $currentExtensions | Where-Object { $_.Replace('\', '/').TrimEnd('/') -ne $normalizedPath }

            if ($newExtensions.Count -lt $currentExtensions.Count) {
                if ($newExtensions.Count -gt 0) {
                    $pathsFormatted = $newExtensions | ForEach-Object { '"' + $_.Replace('\', '/') + '"' }
                    $config["core"]["userextensions"] = "[" + ($pathsFormatted -join ", ") + "]"
                } else {
                    $config["core"].Remove("userextensions")
                }
                Set-IniContent -Path $pyrevitConfigFile -Content $config
                Write-Log "Extension unregistered"
            }
        }
    }
}

$pyrevitCache = "$env:APPDATA\pyRevit\Cache"
if (Test-Path $pyrevitCache) {
    Get-ChildItem -Path $pyrevitCache -Directory -Filter "*CPSK*" | ForEach-Object {
        Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$markerFile = "$env:LOCALAPPDATA\CPSK\pyrevit_installed_by_cpsk.marker"
if (Test-Path $markerFile) {
    Write-Log "pyRevit was installed by CPSK - removing..."
    Remove-Item $markerFile -Force -ErrorAction SilentlyContinue

    $uninstallPaths = @(
        "${env:ProgramFiles}\pyRevit CLI\unins000.exe",
        "${env:ProgramFiles(x86)}\pyRevit CLI\unins000.exe",
        "$env:LOCALAPPDATA\pyRevit-Master\unins000.exe",
        "$env:APPDATA\pyRevit\unins000.exe"
    )

    foreach ($uninstaller in $uninstallPaths) {
        if (Test-Path $uninstaller) {
            Write-Log "Running pyRevit uninstaller: $uninstaller"
            Start-Process -FilePath $uninstaller -ArgumentList "/VERYSILENT /NORESTART" -Wait -ErrorAction SilentlyContinue
            break
        }
    }

    $pyrevitDirs = @("$env:APPDATA\pyRevit", "$env:LOCALAPPDATA\pyRevit-Master", "$env:ProgramData\pyRevit")
    foreach ($dir in $pyrevitDirs) {
        if (Test-Path $dir) { Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue }
    }

    Write-Log "pyRevit removed"
} else {
    Write-Log "pyRevit was not installed by CPSK - keeping it"
}

# Remove Python virtual environment from AppData\Local
$envsPath = "$env:LOCALAPPDATA\cpsk_envs"
$venvPath = "$envsPath\pyrevit_rocket"
if (Test-Path $venvPath) {
    Write-Log "Removing Python virtual environment: $venvPath"
    Remove-Item $venvPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Virtual environment removed"
}

# Remove cpsk_envs folder if empty
if (Test-Path $envsPath) {
    $items = Get-ChildItem $envsPath -ErrorAction SilentlyContinue
    if ($items.Count -eq 0) {
        Remove-Item $envsPath -Force -ErrorAction SilentlyContinue
        Write-Log "Removed empty cpsk_envs folder"
    }
}

# Also check old location (C:\cpsk_envs) for cleanup
$oldEnvsPath = "C:\cpsk_envs"
$oldVenvPath = "$oldEnvsPath\pyrevit_rocket"
if (Test-Path $oldVenvPath) {
    Write-Log "Removing old Python virtual environment: $oldVenvPath"
    Remove-Item $oldVenvPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Old virtual environment removed"
}
if (Test-Path $oldEnvsPath) {
    $items = Get-ChildItem $oldEnvsPath -ErrorAction SilentlyContinue
    if ($items.Count -eq 0) {
        Remove-Item $oldEnvsPath -Force -ErrorAction SilentlyContinue
        Write-Log "Removed old empty cpsk_envs folder"
    }
}

Write-Log "=== Uninstall complete ==="
exit 0
