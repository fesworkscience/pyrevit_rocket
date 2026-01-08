param(
    [Parameter(Mandatory=$true)][string]$ExtensionPath,
    [string]$LogPath = ""
)

$ErrorActionPreference = "Stop"

if ($LogPath -ne "" -and (Test-Path $LogPath)) {
    $logFile = Join-Path $LogPath "register_extension.log"
} else {
    $logFile = "$env:TEMP\CPSK_Register.log"
}

function Write-Log { param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timestamp - $Message"
    $line | Out-File -FilePath $logFile -Append -Encoding UTF8
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

try {
    Write-Log "=========================================="
    Write-Log "=== CPSK Tools: Register Extension ==="
    Write-Log "=========================================="
    Write-Log "Extension Path: $ExtensionPath"
    Write-Log "Log File: $logFile"
    Write-Log ""

    # Check if extension path exists
    Write-Log "Checking if extension path exists..."
    if (Test-Path $ExtensionPath) {
        Write-Log "  [OK] Extension path exists: $ExtensionPath"
        $items = Get-ChildItem -Path $ExtensionPath -ErrorAction SilentlyContinue
        Write-Log "  Contents: $($items.Count) items"
    } else {
        Write-Log "  [ERROR] Extension path does NOT exist: $ExtensionPath"
    }
    Write-Log ""

    # pyRevit config
    $pyrevitConfigDir = "$env:APPDATA\pyRevit"
    $pyrevitConfigFile = "$pyrevitConfigDir\pyRevit_config.ini"

    Write-Log "pyRevit config directory: $pyrevitConfigDir"
    Write-Log "pyRevit config file: $pyrevitConfigFile"
    Write-Log ""

    # Create config directory if needed
    if (-not (Test-Path $pyrevitConfigDir)) {
        Write-Log "Creating pyRevit config directory..."
        New-Item -ItemType Directory -Path $pyrevitConfigDir -Force | Out-Null
        Write-Log "  [OK] Created: $pyrevitConfigDir"
    } else {
        Write-Log "  [OK] pyRevit config directory exists"
    }
    Write-Log ""

    # Get extension parent path (the folder containing CPSK.extension)
    $extensionParentPath = Split-Path -Parent $ExtensionPath
    Write-Log "Extension parent path (to register): $extensionParentPath"
    Write-Log ""

    # Read current config
    Write-Log "Reading pyRevit config file..."
    if (Test-Path $pyrevitConfigFile) {
        Write-Log "  [OK] Config file exists"
        $configContent = Get-Content -Path $pyrevitConfigFile -Raw -ErrorAction SilentlyContinue
        Write-Log "  Current content:"
        Write-Log "  ----------------"
        $configContent -split "`n" | ForEach-Object { Write-Log "  $_" }
        Write-Log "  ----------------"
    } else {
        Write-Log "  [!] Config file does not exist, will create new one"
    }
    Write-Log ""

    $config = Get-IniContent -Path $pyrevitConfigFile
    if (-not $config.ContainsKey("core")) {
        $config["core"] = @{}
        Write-Log "Created [core] section"
    }

    # Helper function to normalize path (convert all slashes to forward, remove doubles)
    function Normalize-Path { param([string]$Path)
        # Replace backslashes with forward slashes, then collapse multiple slashes
        $normalized = $Path -replace '[/\\]+', '/'
        return $normalized.TrimEnd('/')
    }

    # Parse current extensions
    $currentExtensions = @()
    if ($config["core"].ContainsKey("userextensions")) {
        $extensionsStr = $config["core"]["userextensions"]
        Write-Log "Current userextensions value: $extensionsStr"
        if ($extensionsStr -match '^\[(.+)\]$') {
            $innerContent = $matches[1]
            # Split by comma, trim quotes and whitespace, normalize paths
            $currentExtensions = $innerContent -split ',' | ForEach-Object {
                $path = $_.Trim().Trim('"').Trim("'")
                if ($path -ne "") {
                    Normalize-Path $path
                }
            } | Where-Object { $_ -ne $null -and $_ -ne "" }
            Write-Log "Parsed extensions:"
            foreach ($ext in $currentExtensions) {
                Write-Log "  - $ext"
            }
        }
    } else {
        Write-Log "No userextensions key found in config"
    }
    Write-Log ""

    # Check if already registered
    $normalizedNewPath = Normalize-Path $extensionParentPath
    Write-Log "Normalized path to register: $normalizedNewPath"

    $alreadyRegistered = $false
    foreach ($existingPath in $currentExtensions) {
        Write-Log "  Comparing with: $existingPath"
        if ($existingPath -eq $normalizedNewPath) {
            $alreadyRegistered = $true
            Write-Log "  [MATCH] Already registered!"
            break
        }
    }
    Write-Log ""

    # Register if needed
    if (-not $alreadyRegistered) {
        Write-Log "Registering new extension path..."
        $currentExtensions += $normalizedNewPath
        # Format each path with quotes
        $pathsFormatted = @()
        foreach ($p in $currentExtensions) {
            $pathsFormatted += "`"$p`""
        }
        $newValue = "[" + ($pathsFormatted -join ", ") + "]"
        $config["core"]["userextensions"] = $newValue
        Write-Log "New userextensions value: $newValue"

        Write-Log "Writing config file..."
        Set-IniContent -Path $pyrevitConfigFile -Content $config
        Write-Log "  [OK] Config file updated"

        # Verify write
        if (Test-Path $pyrevitConfigFile) {
            $verifyContent = Get-Content -Path $pyrevitConfigFile -Raw
            Write-Log "  Verification - new content:"
            Write-Log "  ----------------"
            $verifyContent -split "`n" | ForEach-Object { Write-Log "  $_" }
            Write-Log "  ----------------"
        }
    } else {
        Write-Log "Extension already registered, skipping"
    }
    Write-Log ""

    Write-Log "=========================================="
    Write-Log "=== Registration complete ==="
    Write-Log "=========================================="
    exit 0
}
catch {
    Write-Log "=========================================="
    Write-Log "=== ERROR ==="
    Write-Log "=========================================="
    Write-Log "Exception: $($_.Exception.Message)"
    Write-Log "Stack trace: $($_.ScriptStackTrace)"
    exit 1
}
