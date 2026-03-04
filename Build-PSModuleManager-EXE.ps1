<#
.SYNOPSIS
    Builds PSModuleManager.exe from PSModuleManager.ps1 using ps2exe.
.DESCRIPTION
    Run this script ONCE on your Windows machine (as Administrator recommended).
    It will:
      1. Install ps2exe if not already installed
      2. Compile PSModuleManager.ps1 → PSModuleManager.exe with icon
      3. Place the .exe next to this script
.NOTES
    Requirements: PowerShell 5.1 or 7, internet access (first run)
    Run from the folder containing PSModuleManager.ps1 and PSModuleManager.ico
#>


$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path $MyInvocation.MyCommand.Path
$SourcePS1 = Join-Path $ScriptDir 'PSModuleManager.ps1'
$IconFile  = Join-Path $ScriptDir 'PSModuleManager.ico'
$OutputEXE = Join-Path $ScriptDir 'PSModuleManager.exe'

Write-Host "`n=== PSModuleManager EXE Builder ===" -ForegroundColor Cyan

if (-not (Test-Path $SourcePS1)) {
    Write-Host "[ERROR] PSModuleManager.ps1 not found in: $ScriptDir" -ForegroundColor Red; exit 1
}
if (-not (Test-Path $IconFile)) {
    Write-Host "[WARN]  PSModuleManager.ico not found - building without icon." -ForegroundColor Yellow
    $IconFile = $null
}

# Install ps2exe if needed
Write-Host "`n[1/3] Checking ps2exe..." -ForegroundColor White
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "      Installing ps2exe from PSGallery..." -ForegroundColor Yellow
    Install-Module ps2exe -Scope CurrentUser -Force -AllowClobber
    Write-Host "      [OK] Installed." -ForegroundColor Green
} else {
    Write-Host "      [OK] ps2exe found." -ForegroundColor Green
}
Import-Module ps2exe -Force

# Compile
Write-Host "`n[2/3] Compiling..." -ForegroundColor White

$params = @{
    inputFile    = $SourcePS1
    outputFile   = $OutputEXE
    title        = 'PowerShell Module Manager'
    description  = 'GUI tool for managing PS modules across Windows PowerShell 5.1 and PS7'
    company      = 'karanik'
    product      = 'PSModuleManager'
    version      = '7.0'
    copyright    = '© 2026 Nikolaos Karanikolas'
    requireAdmin = $true
    noConsole    = $true
    x64          = $true
    STA          = $true
    noOutput     = $true   # suppress ps2exe progress messages
    noError      = $false
}
if ($IconFile) { $params['iconFile'] = $IconFile }

try {
    Invoke-ps2exe @params
} catch {
    Write-Host "[ERROR] $_ " -ForegroundColor Red; exit 1
}

# Verify
Write-Host "`n[3/3] Verifying..." -ForegroundColor White
if (Test-Path $OutputEXE) {
    $sz = [math]::Round((Get-Item $OutputEXE).Length / 1MB, 2)
    Write-Host "      [OK] PSModuleManager.exe  ($sz MB)" -ForegroundColor Green
    Write-Host "      Path: $OutputEXE" -ForegroundColor DarkGray
} else {
    Write-Host "[ERROR] EXE not found after build." -ForegroundColor Red; exit 1
}

Write-Host "`n=== Done! Run PSModuleManager.exe directly. ===" -ForegroundColor Cyan
