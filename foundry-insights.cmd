@echo off
rem FoundryInsights launcher for cmd.exe (Windows).
rem Runs scripts\Run-FoundryInsights.ps1 with PowerShell 7.2+ (pwsh); arguments are passed through unchanged:
rem   foundry-insights.cmd -TenantId <tenantId> -Days 30 -FetchPrices
rem   foundry-insights.cmd --help
rem Set FI_PWSH to use a specific pwsh executable. From PowerShell, run the .ps1 script directly instead.
setlocal
set "FI_SCRIPT=%~dp0scripts\Run-FoundryInsights.ps1"
if not defined FI_PWSH (
    where pwsh >nul 2>&1 || (
        echo PowerShell 7.2+ ^(pwsh^) is required but was not found on PATH. 1>&2
        echo Install it with: winget install --id Microsoft.PowerShell 1>&2
        exit /b 9009
    )
    set "FI_PWSH=pwsh"
)
if /i "%~1"=="--help" goto help
if /i "%~1"=="-h" goto help
if /i "%~1"=="/?" goto help
"%FI_PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%FI_SCRIPT%" %*
exit /b %ERRORLEVEL%

:help
"%FI_PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%FI_SCRIPT%" -?
exit /b %ERRORLEVEL%
