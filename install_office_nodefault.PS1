<#
.SYNOPSIS
    Installs Office LTSC 2024 using only the latest Office Deployment Tool from the
    Microsoft Download Center (no pinned fallback build).

.DESCRIPTION
    Kept for compatibility with the original install_office_nodefault.PS1 entry point.
    It runs Install_Office.ps1 (which must be in the same folder) with -OdtSource Latest.
    Every Install_Office.ps1 parameter is supported, for example:

        .\install_office_nodefault.PS1
        .\install_office_nodefault.PS1 -Unattended -Edition ProPlus -AcceptEula

    Run "Get-Help .\Install_Office.ps1 -Detailed" for all parameters and exit codes.
#>

$installer = Join-Path $PSScriptRoot 'Install_Office.ps1'
if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) {
    Write-Host "[!] Install_Office.ps1 was not found next to this script ($PSScriptRoot)." -ForegroundColor Red
    exit 1
}

& $installer -OdtSource Latest @args
exit $LASTEXITCODE
