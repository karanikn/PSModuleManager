# build.ps1 - KeePassNetworkChecker
param([string]$KeePassPath = "")

if ($KeePassPath -eq "" -and $env:KEEPASS_PATH) { $KeePassPath = $env:KEEPASS_PATH }
if (-not (Test-Path $KeePassPath)) {
    Write-Host "Set: `$env:KEEPASS_PATH = 'C:\...\KeePass.exe'" -ForegroundColor Red; exit 1
}
Write-Host "KeePass  : $KeePassPath" -ForegroundColor Green

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" 2>$null | Select-Object -First 1
if (-not (Test-Path $msbuild)) { Write-Host "MSBuild not found" -ForegroundColor Red; exit 1 }
Write-Host "MSBuild  : $msbuild" -ForegroundColor Green

$srcDir = (Get-Location).Path
$outDir = Join-Path $srcDir "bin\Release"
$dll    = Join-Path $outDir "KeePassNetworkChecker.dll"

# ── 1. Clean build artifacts ─────────────────────────────────────────────────
Remove-Item -Path obj, bin -Recurse -Force -ErrorAction SilentlyContinue

# ── 2. Build DLL ──────────────────────────────────────────────────────────────
Write-Host "Building..." -ForegroundColor Cyan
& $msbuild "KeePassNetworkChecker.csproj" /p:Configuration=Release /p:KEEPASS_PATH="$KeePassPath" /nologo /verbosity:minimal
if ($LASTEXITCODE -ne 0) { Write-Host "Build FAILED" -ForegroundColor Red; exit 1 }
Write-Host "DLL      -> $dll" -ForegroundColor Green

# ── 3. Copy DLL to Plugins ────────────────────────────────────────────────────
$pluginsDir = Join-Path (Split-Path $KeePassPath) "Plugins"
if (-not (Test-Path $pluginsDir)) { New-Item -ItemType Directory -Path $pluginsDir | Out-Null }

# Remove any leftover plgx that would cause compile errors
Remove-Item (Join-Path $pluginsDir "KeePassNetworkChecker.plgx") -Force -ErrorAction SilentlyContinue
Copy-Item $dll $pluginsDir -Force
Write-Host "Installed -> $pluginsDir" -ForegroundColor Green

# ── 4. Clear KeePass plugin cache ─────────────────────────────────────────────
Write-Host "Clearing plugin cache..." -ForegroundColor Cyan
Remove-Item "$env:LOCALAPPDATA\KeePass\PluginCache\*" -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Done. Restart KeePass." -ForegroundColor Cyan
