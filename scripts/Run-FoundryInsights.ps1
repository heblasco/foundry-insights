<#
.SYNOPSIS
    One-command FoundryInsights run: checks prerequisites, extracts per-minute metrics and builds the report.

.DESCRIPTION
    1. Checks PowerShell 7, Python 3.10+ and the Python packages in requirements.txt (-InstallPythonDeps
       installs them with pip).
    2. Runs Export-FoundryTokenMetrics.ps1 into a new timestamped folder (<OutputRoot>\<RunName>_<yyyyMMdd-HHmmss>Z).
    3. Runs foundry_token_stats.py on that folder (CSVs + Excel report).
    Any other Export-FoundryTokenMetrics.ps1 parameter can be passed through -ExportArgs: a hashtable when the
    script is invoked in-process (& .\scripts\Run-FoundryInsights.ps1), or a JSON object string with pwsh -File.

.EXAMPLE
    pwsh -File .\scripts\Run-FoundryInsights.ps1 -TenantId <tenant> -Days 30 -FetchPrices

.EXAMPLE
    pwsh -File .\scripts\Run-FoundryInsights.ps1 -TenantId <tenant> -AuthMode ManagedIdentity -ExportArgs '{"IncludePtuUtilization": true, "Parallel": 8}'

.EXAMPLE
    & .\scripts\Run-FoundryInsights.ps1 -TenantId <tenant> -ExportArgs @{ IncludePtuUtilization = $true; Parallel = 8 }
#>
[CmdletBinding()]
param(
    [string[]] $TenantId,
    [string[]] $SubscriptionId,
    [ValidateRange(1, 93)] [int] $Days = 30,
    [string]   $OutputRoot = (Join-Path (Get-Location) 'output'),
    [ValidatePattern('^[\w.-]+$')] [string] $RunName = 'foundry',
    [ValidateSet('Auto', 'AzCli', 'AzPowerShell', 'ManagedIdentity')] [string] $AuthMode = 'Auto',
    [object]   $ExportArgs = @{},
    [string]   $PricesFile,
    [switch]   $FetchPrices,
    [switch]   $NoExcel,
    [switch]   $InstallPythonDeps,
    [string]   $Python
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

# -ExportArgs: a hashtable when called in-process (& .\Run-FoundryInsights.ps1), a JSON object with pwsh -File.
if ($ExportArgs -is [string]) {
    try { $ExportArgs = $ExportArgs | ConvertFrom-Json -AsHashtable } catch { throw "-ExportArgs is not valid JSON: $ExportArgs" }
}
if ($ExportArgs -isnot [System.Collections.IDictionary]) { throw '-ExportArgs must be a hashtable or a JSON object string.' }

if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7+ is required (pwsh).' }
$versionCheck = 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)'
$depsCheck = 'import duckdb, pandas, openpyxl'
if (-not $Python) {
    # Prefer the first Python 3.10+ that already has the packages; "python3" may be the Store stub or a bare install.
    $valid = foreach ($candidate in 'python', 'python3', 'py') {
        if (-not (Get-Command $candidate -CommandType Application -ErrorAction SilentlyContinue)) { continue }
        & $candidate -c $versionCheck 2>$null
        if ($LASTEXITCODE -eq 0) { $candidate }
    }
    if (-not $valid) { throw 'Python 3.10+ is required (python, python3 or py on PATH, or pass -Python).' }
    $Python = @($valid | Where-Object { & $_ -c $depsCheck 2>$null; $LASTEXITCODE -eq 0 } | Select-Object -First 1)[0]
    if (-not $Python) { $Python = @($valid)[0] }
}
else {
    & $Python -c $versionCheck
    if ($LASTEXITCODE) { throw "$Python is not Python 3.10+." }
}
& $Python -c $depsCheck 2>$null
if ($LASTEXITCODE) {
    if (-not $InstallPythonDeps) { throw "Missing Python packages. Run: $Python -m pip install -r requirements.txt (or use -InstallPythonDeps)." }
    & $Python -m pip install --quiet -r (Join-Path $root 'requirements.txt')
    if ($LASTEXITCODE) { throw 'pip install failed.' }
}

$runDir = Join-Path $OutputRoot ('{0}_{1:yyyyMMdd-HHmmss}Z' -f $RunName, [datetime]::UtcNow)
$export = @{ OutputDirectory = $runDir; Days = $Days; AuthMode = $AuthMode } + $ExportArgs
if ($TenantId) { $export.TenantId = $TenantId }
if ($SubscriptionId) { $export.SubscriptionId = $SubscriptionId }
& (Join-Path $PSScriptRoot 'Export-FoundryTokenMetrics.ps1') @export
if (-not (Test-Path (Join-Path $runDir 'manifest.csv'))) { throw "Extraction did not produce $runDir\manifest.csv." }
if ($ExportArgs['InventoryOnly']) { Write-Host "Inventory only: $runDir"; return }

$statsArgs = @((Join-Path $PSScriptRoot 'foundry_token_stats.py'), '--run-dir', $runDir)
if ($PricesFile) { $statsArgs += @('--prices', $PricesFile) }
if ($FetchPrices) { $statsArgs += '--fetch-prices' }
if ($NoExcel) { $statsArgs += '--no-excel' }
& $Python @statsArgs
if ($LASTEXITCODE) { throw "foundry_token_stats.py failed (exit $LASTEXITCODE)." }
Write-Host "Done: $runDir"
