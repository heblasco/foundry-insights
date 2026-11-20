#!/usr/bin/env bash
# FoundryInsights launcher for bash (Linux, macOS, WSL, Git Bash).
# Runs scripts/Run-FoundryInsights.ps1 with PowerShell 7.2+ (pwsh); arguments are passed through unchanged:
#   ./foundry-insights.sh -TenantId <tenantId> -Days 30 -FetchPrices
#   ./foundry-insights.sh --help
# Set FI_PWSH to use a specific pwsh executable.
set -euo pipefail

# Resolve the repository folder, following symlinks (no readlink -f on macOS).
src="${BASH_SOURCE[0]}"
while [ -L "$src" ]; do
    dir="$(cd -P "$(dirname "$src")" && pwd)"
    src="$(readlink "$src")"
    case "$src" in /*) ;; *) src="$dir/$src" ;; esac
done
root="$(cd -P "$(dirname "$src")" && pwd)"
script="$root/scripts/Run-FoundryInsights.ps1"
# Git Bash / MSYS / Cygwin: pwsh.exe is a Windows program and needs a Windows path.
if command -v cygpath >/dev/null 2>&1; then script="$(cygpath -w "$script")"; fi

pwsh_bin="${FI_PWSH:-}"
if [ -z "$pwsh_bin" ]; then
    for candidate in pwsh pwsh-preview; do
        if command -v "$candidate" >/dev/null 2>&1; then pwsh_bin="$candidate"; break; fi
    done
fi
if [ -z "$pwsh_bin" ]; then
    echo "PowerShell 7.2+ (pwsh) is required but was not found on PATH." >&2
    case "$(uname -s)" in
        Darwin) echo "Install it with: brew install --cask powershell" >&2 ;;
        Linux) echo "Install it: https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-linux" >&2 ;;
        MINGW* | MSYS* | CYGWIN*) echo "Install it with: winget install --id Microsoft.PowerShell" >&2 ;;
    esac
    exit 127
fi

if [ "$#" -eq 1 ] && { [ "$1" = "-h" ] || [ "$1" = "--help" ]; }; then
    exec "$pwsh_bin" -NoProfile -File "$script" '-?'
fi
exec "$pwsh_bin" -NoProfile -File "$script" "$@"
