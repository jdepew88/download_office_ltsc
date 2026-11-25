<#
.SYNOPSIS
    Optional, separate activation helper for volume licensed Office LTSC 2024.

.DESCRIPTION
    Uses Microsoft's Office Software Protection Platform script (OSPP.VBS), which is
    installed with volume licensed Office, to check or activate Office LTSC 2024.

    Supported, Microsoft-authorised activation methods:
      - KMS:  Office installs with Microsoft's Generic Volume License Key (GVLK) and
              activates automatically against a KMS host published in DNS. Use
              -KmsHost/-KmsPort only to point Office at a specific KMS host.
      - Active Directory-based activation: automatic, no action required.
      - MAK:  supply your organisation's Multiple Activation Key with -ProductKey.

    No product keys are stored in this script or repository. Keys are only accepted
    at run time and are masked in all output. The Volume Activation Management Tool
    (VAMT) can also be used to manage activation centrally.

.PARAMETER Status
    Show the license status (OSPP.VBS /dstatus). This is the default action.

.PARAMETER KmsHost
    KMS host name or IP address to use (OSPP.VBS /sethst), followed by activation.

.PARAMETER KmsPort
    KMS host port (OSPP.VBS /setprt). Default port on the KMS host is 1688.

.PARAMETER ClearKmsHost
    Remove a manually configured KMS host so DNS auto-discovery is used (OSPP.VBS /remhst).

.PARAMETER ProductKey
    Your organisation's MAK product key (OSPP.VBS /inpkey), followed by activation.

.PARAMETER Activate
    Attempt activation of installed Office products now (OSPP.VBS /act).

.PARAMETER OsppPath
    Full path to OSPP.VBS if Office is not in the default location.

.PARAMETER DryRun
    Print the OSPP.VBS commands that would be run (product key masked) without running them.

.EXAMPLE
    .\Activate_Office.ps1
    Show the current license status.

.EXAMPLE
    .\Activate_Office.ps1 -KmsHost kms.contoso.com
    Point Office at a KMS host and activate.

.EXAMPLE
    .\Activate_Office.ps1 -ProductKey $env:OFFICE_MAK
    Install a MAK from a deployment secret and activate.

.NOTES
    Exit codes:
       0  Success
       1  Invalid arguments
       2  OSPP.VBS not found (volume licensed Office not installed?)
       4  Administrator rights required
       5  OSPP.VBS reported an error (including an activation error listed by /dstatus)
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [switch]$Status,

    [ValidatePattern('^[A-Za-z0-9]([A-Za-z0-9.-]*[A-Za-z0-9])?$|^\[?[0-9A-Fa-f:.]+\]?$')]
    [string]$KmsHost,

    [ValidateRange(1, 65535)]
    [int]$KmsPort,

    [switch]$ClearKmsHost,

    [ValidatePattern('^[A-Za-z0-9]{5}(-[A-Za-z0-9]{5}){4}$')]
    [string]$ProductKey,

    [switch]$Activate,

    [string]$OsppPath,

    [switch]$DryRun
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Get-MaskedProductKey {
    param([string]$Key)
    return "*****-*****-*****-*****-$($Key.Substring($Key.Length - 5))"
}

function Find-Ospp {
    param([string]$Path)
    if ($Path) {
        if (Test-Path -LiteralPath $Path -PathType Leaf) { return (Resolve-Path -LiteralPath $Path).ProviderPath }
        return $null
    }
    $candidates = @(
        (Join-Path $env:ProgramFiles 'Microsoft Office\Office16\OSPP.VBS')
    )
    if (${env:ProgramFiles(x86)}) {
        $candidates += (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\Office16\OSPP.VBS')
    }
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return $null
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal $identity
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if ($KmsHost -and $ClearKmsHost) {
    Write-Host '[!] -KmsHost and -ClearKmsHost cannot be used together.' -ForegroundColor Red
    exit 1
}
if ($KmsPort -and -not $KmsHost) {
    Write-Host '[!] -KmsPort requires -KmsHost.' -ForegroundColor Red
    exit 1
}

# Build the list of OSPP.VBS operations. Each has the real argument and a display form.
$operations = @()
if ($ClearKmsHost) {
    $operations += @{ Argument = '/remhst'; Display = '/remhst' }
}
if ($KmsHost) {
    $operations += @{ Argument = "/sethst:$KmsHost"; Display = "/sethst:$KmsHost" }
    if ($KmsPort) {
        $operations += @{ Argument = "/setprt:$KmsPort"; Display = "/setprt:$KmsPort" }
    }
}
if ($ProductKey) {
    $operations += @{ Argument = "/inpkey:$($ProductKey.ToUpperInvariant())"; Display = "/inpkey:$(Get-MaskedProductKey $ProductKey)" }
}
if ($Activate -or $KmsHost -or $ProductKey) {
    $operations += @{ Argument = '/act'; Display = '/act' }
}
$changesLicensing = ($operations.Count -gt 0)
if ($Status -or -not $changesLicensing) {
    $operations += @{ Argument = '/dstatus'; Display = '/dstatus' }
}

$ospp = Find-Ospp -Path $OsppPath
if (-not $ospp) {
    if ($DryRun) {
        Write-Warning 'OSPP.VBS was not found; showing commands with a placeholder path.'
        $ospp = '<path to OSPP.VBS>'
    } else {
        Write-Host '[!] OSPP.VBS was not found. Is volume licensed Office LTSC 2024 installed? Use -OsppPath to specify its location.' -ForegroundColor Red
        exit 2
    }
}

if ($DryRun) {
    Write-Host '[*] Dry run: the following commands would be run:'
    foreach ($operation in $operations) {
        Write-Host "  cscript.exe //NoLogo `"$ospp`" $($operation.Display)"
    }
    exit 0
}

if ($changesLicensing -and -not (Test-IsAdministrator)) {
    Write-Host '[!] Changing Office licensing requires an elevated PowerShell session.' -ForegroundColor Red
    exit 4
}

$failed = $false
foreach ($operation in $operations) {
    Write-Host "[*] cscript.exe //NoLogo `"$ospp`" $($operation.Display)"
    $output = & cscript.exe //NoLogo $ospp $operation.Argument 2>&1 | ForEach-Object { "$_" }
    $text = ($output -join [Environment]::NewLine)
    if ($ProductKey) {
        $text = $text.Replace($ProductKey.ToUpperInvariant(), (Get-MaskedProductKey $ProductKey))
    }
    Write-Host $text
    if ($LASTEXITCODE -ne 0 -or $text -match 'ERROR CODE:') {
        $failed = $true
        if ($operation.Argument -ne '/dstatus') {
            Write-Host "[!] OSPP.VBS $($operation.Display) reported an error." -ForegroundColor Red
            break
        }
    }
}

if ($failed) {
    exit 5
}
exit 0
