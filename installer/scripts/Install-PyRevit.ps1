param(
    [Parameter(Mandatory=$true)][string]$RequiredVersion,
    [Parameter(Mandatory=$true)][string]$DownloadUrl
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$logFile = "$env:TEMP\CPSK_Install.log"

function Write-Log { param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append -Encoding UTF8
    Write-Host $Message
}

function Get-PyRevitInfo {
    $regPath = "HKCU:\Software\pyRevit"
    if (Test-Path $regPath) {
        $installPath = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).InstallPath
        $version = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).Version
        return @{ Installed = $true; Path = $installPath; Version = $version }
    }

    $commonPaths = @("$env:APPDATA\pyRevit-Master", "$env:APPDATA\pyRevit", "$env:ProgramData\pyRevit", "C:\pyRevit")
    foreach ($path in $commonPaths) {
        if (Test-Path "$path\bin\pyrevit.exe") {
            try {
                $versionOutput = & "$path\bin\pyrevit.exe" --version 2>&1
                if ($versionOutput -match "(\d+\.\d+\.\d+)") { $version = $matches[1] }
            } catch { $version = "unknown" }
            return @{ Installed = $true; Path = $path; Version = $version }
        }
    }
    return @{ Installed = $false; Path = $null; Version = $null }
}

function Compare-Versions { param([string]$Current, [string]$Required)
    if ([string]::IsNullOrEmpty($Current)) { return -1 }
    try {
        $currentClean = $Current -replace '[^\d\.]', '' -replace '\.+$', ''
        $requiredClean = $Required -replace '[^\d\.]', '' -replace '\.+$', ''
        $currentParts = $currentClean.Split('.') | ForEach-Object { [int]$_ }
        $requiredParts = $requiredClean.Split('.') | ForEach-Object { [int]$_ }
        while ($currentParts.Count -lt 4) { $currentParts += 0 }
        while ($requiredParts.Count -lt 4) { $requiredParts += 0 }
        for ($i = 0; $i -lt 4; $i++) {
            if ($currentParts[$i] -lt $requiredParts[$i]) { return -1 }
            if ($currentParts[$i] -gt $requiredParts[$i]) { return 1 }
        }
        return 0
    } catch { return -1 }
}

Write-Log "=== CPSK Tools: PyRevit Installation ==="
Write-Log "Required version: $RequiredVersion"

$pyrevitInfo = Get-PyRevitInfo
Write-Log "Installed: $($pyrevitInfo.Installed), Current version: $($pyrevitInfo.Version)"

$markerFile = "$env:LOCALAPPDATA\CPSK\pyrevit_installed_by_cpsk.marker"

if ($pyrevitInfo.Installed) {
    $comparison = Compare-Versions -Current $pyrevitInfo.Version -Required $RequiredVersion
    if ($comparison -ge 0) {
        Write-Log "pyRevit version $($pyrevitInfo.Version) is same or newer - skipping installation"
        exit 0
    }
    Write-Log "pyRevit version $($pyrevitInfo.Version) is older than $RequiredVersion - updating"
} else {
    Write-Log "pyRevit not found - installing"
}

$installerPath = "$env:TEMP\pyRevit_installer.exe"
Write-Log "Downloading: $DownloadUrl"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$webClient = New-Object System.Net.WebClient
$webClient.DownloadFile($DownloadUrl, $installerPath)

Write-Log "Running installer..."
$process = Start-Process -FilePath $installerPath -ArgumentList "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES" -Wait -PassThru
Write-Log "Installer exit code: $($process.ExitCode)"

if (Test-Path $installerPath) { Remove-Item $installerPath -Force -ErrorAction SilentlyContinue }

if (-not $pyrevitInfo.Installed) {
    $markerDir = Split-Path -Parent $markerFile
    if (-not (Test-Path $markerDir)) { New-Item -ItemType Directory -Path $markerDir -Force | Out-Null }
    "installed=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $markerFile -Encoding UTF8
    Write-Log "Marker created (pyRevit was not installed before)"
}

Write-Log "=== PyRevit installation complete ==="
exit 0
