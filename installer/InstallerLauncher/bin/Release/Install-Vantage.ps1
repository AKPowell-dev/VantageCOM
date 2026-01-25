param(
    [string]$SourceDir = "",
    [string]$InstallDir = "",
    [switch]$Force32,
    [switch]$Force64,
    [switch]$NoRegister,
    [switch]$NoXlam,
    [string]$LogPath = "",
    [string]$StatusPath = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($SourceDir)) {
    $SourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$ProgId = "VantagePackageHolder.Addin"
$FriendlyName = "Vantage"
$Description = "Vantage Excel Add-in"

$transcriptStarted = $false
if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
    try {
        $logDir = Split-Path -Parent $LogPath
        if (-not [string]::IsNullOrWhiteSpace($logDir)) {
            New-Item -ItemType Directory -Force -Path $logDir | Out-Null
        }
        Start-Transcript -Path $LogPath -Append | Out-Null
        $transcriptStarted = $true
    } catch {
        # ignore logging failures
    }
}

try {
    function Write-Status([string]$message) {
        if ([string]::IsNullOrWhiteSpace($StatusPath)) {
            return
        }
        try {
            Set-Content -Path $StatusPath -Value $message -Encoding UTF8
        } catch {
            # ignore status failures
        }
    }

    Write-Status "Preparing installer..."
    try {
        Get-ChildItem -Path $SourceDir -Recurse -File | Unblock-File
    } catch {
        # ignore unblock failures
    }

function Get-OfficeBitness {
    $candidates = @(
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"; Name = "Platform" },
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"; Name = "OfficeClientEdition" },
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook"; Name = "Bitness" },
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Office\15.0\Outlook"; Name = "Bitness" },
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Office\14.0\Outlook"; Name = "Bitness" }
    )

    foreach ($candidate in $candidates) {
        try {
            $value = (Get-ItemProperty -Path $candidate.Path -ErrorAction Stop).$($candidate.Name)
            if ($value -match "64") { return "x64" }
            if ($value -match "86") { return "x86" }
        } catch {
            # ignore
        }
    }

    return "x64"
}

function Get-RegasmPath([string]$bitness) {
    $frameworkRoot = if ($bitness -eq "x86") { "Framework" } else { "Framework64" }
    $regasm = Join-Path $env:WINDIR ("Microsoft.NET\" + $frameworkRoot + "\v4.0.30319\RegAsm.exe")
    if (!(Test-Path $regasm)) {
        throw "RegAsm.exe not found at $regasm"
    }
    return $regasm
}

function Register-ExcelComAddin([string]$bitness) {
    $baseKey = "HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins\" + $ProgId
    if ($bitness -eq "x86" -and (Test-Path "HKLM:\SOFTWARE\WOW6432Node")) {
        $baseKey = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Excel\Addins\" + $ProgId
    }

    New-Item -Path $baseKey -Force | Out-Null
    New-ItemProperty -Path $baseKey -Name "FriendlyName" -Value $FriendlyName -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $baseKey -Name "Description" -Value $Description -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $baseKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $baseKey -Name "CommandLineSafe" -Value 0 -PropertyType DWord -Force | Out-Null
}

function Add-ExcelOpenEntry([string]$xlamPath) {
    $versions = @("16.0", "15.0", "14.0")
    $targetKey = $null
    foreach ($v in $versions) {
        $candidate = "HKCU:\Software\Microsoft\Office\$v\Excel\Options"
        if (Test-Path $candidate) {
            $targetKey = $candidate
            break
        }
    }
    if ($null -eq $targetKey) {
        $targetKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
        New-Item -Path $targetKey -Force | Out-Null
    }

    $existing = Get-ItemProperty -Path $targetKey
    $openNames = @("OPEN")
    for ($i = 1; $i -le 20; $i++) { $openNames += "OPEN$i" }

    foreach ($name in $openNames) {
        if ($existing.PSObject.Properties.Name -contains $name) {
            if ($existing.$name -eq $xlamPath) { return }
        }
    }

    foreach ($name in $openNames) {
        if (!($existing.PSObject.Properties.Name -contains $name)) {
            New-ItemProperty -Path $targetKey -Name $name -Value $xlamPath -PropertyType String -Force | Out-Null
            return
        }
    }

    Set-ItemProperty -Path $targetKey -Name "OPEN" -Value $xlamPath -Force | Out-Null
}

Write-Status "Detecting Office bitness..."
$bitness = if ($Force32) { "x86" } elseif ($Force64) { "x64" } else { Get-OfficeBitness }
Write-Status ("Installing for " + $bitness + " Office...")

if ([string]::IsNullOrWhiteSpace($InstallDir)) {
    $base = $env:ProgramFiles
    if ($bitness -eq "x86" -and $env:ProgramFiles(x86)) { $base = $env:ProgramFiles(x86) }
    $InstallDir = Join-Path $base "Vantage"
}

$comDllSource = Join-Path $SourceDir "VantagePackageHolder.dll"
$extensibilitySource = Join-Path $SourceDir "Extensibility.dll"
$xlamSource = Join-Path $SourceDir "Vantage.xlam"
$resourcesSource = Join-Path $SourceDir "Resources"

if (!(Test-Path $comDllSource)) {
    throw "Missing VantagePackageHolder.dll in $SourceDir"
}

Write-Status "Copying add-in files..."
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
Copy-Item -Path $comDllSource -Destination $InstallDir -Force
if (Test-Path $extensibilitySource) {
    Copy-Item -Path $extensibilitySource -Destination $InstallDir -Force
}
if (Test-Path $resourcesSource) {
    Copy-Item -Path $resourcesSource -Destination $InstallDir -Recurse -Force
}

if (-not $NoRegister) {
    Write-Status "Registering COM add-in..."
    $regasm = Get-RegasmPath $bitness
    $dllPath = Join-Path $InstallDir "VantagePackageHolder.dll"
    $tlbPath = Join-Path $InstallDir "VantagePackageHolder.tlb"
    & $regasm /nologo /codebase /tlb:$tlbPath $dllPath
    if ($LASTEXITCODE -ne 0) {
        throw "RegAsm failed with exit code $LASTEXITCODE"
    }
    Write-Status "Registering Excel add-in..."
    Register-ExcelComAddin $bitness
}

if (-not $NoXlam) {
    Write-Status "Installing Vantage.xlam..."
    if (!(Test-Path $xlamSource)) {
        throw "Missing Vantage.xlam in $SourceDir"
    }
    $addinDir = Join-Path $env:APPDATA "Microsoft\AddIns"
    New-Item -ItemType Directory -Force -Path $addinDir | Out-Null
    $xlamDest = Join-Path $addinDir "Vantage.xlam"
    Copy-Item -Path $xlamSource -Destination $xlamDest -Force
    Add-ExcelOpenEntry $xlamDest
}

Write-Host "Vantage install complete."
Write-Status "Install complete."
}
catch {
    Write-Status ("FAILED: " + $_.Exception.Message)
    throw
}
finally {
    if ($transcriptStarted) {
        try { Stop-Transcript | Out-Null } catch { }
    }
}
