<#
.SYNOPSIS
    Downloads and installs Microsoft Office LTSC 2024 (Standard or Professional Plus)
    with Microsoft's Office Deployment Tool (ODT).

.DESCRIPTION
    Installs volume licensed Office LTSC Standard 2024 (Standard2024Volume) or
    Office LTSC Professional Plus 2024 (ProPlus2024Volume) from the
    PerpetualVL2024 channel, or any custom ODT configuration.xml.

    Run without parameters for the interactive menu (the original behaviour).
    Use -Unattended for silent, prompt-free deployments, for example from
    Windows Deployment Builder, a task sequence or a provisioning package.

    The script:
      1. Builds and validates the ODT configuration.xml.
      2. Downloads the current Office Deployment Tool from Microsoft and
         verifies its Microsoft Authenticode signature.
      3. Runs "setup.exe /download" and "setup.exe /configure".
      4. Optionally removes the downloaded files.

    Activation is NOT performed by this script. Volume licensed Office LTSC 2024
    installs with Microsoft's Generic Volume License Key (GVLK) and activates
    automatically against a KMS host or Active Directory-based activation.
    For MAK activation or to set a KMS host, run Activate_Office.ps1 afterwards.

.PARAMETER Edition
    Standard, ProPlus or Custom. Required with -Unattended (unless -ConfigPath is
    given, which implies Custom). If omitted interactively, a menu is shown.

.PARAMETER ConfigPath
    Path to a custom ODT configuration XML. Implies -Edition Custom.
    Default for Custom: Configuration.xml next to this script.

.PARAMETER Unattended
    Never prompt. Sets <Display Level="None"> in the configuration, defaults
    -Cleanup to Yes and requires -Edition or -ConfigPath.

.PARAMETER AcceptEula
    Sets AcceptEULA="TRUE" so users are not shown the license terms on first
    launch. Only use this if your organisation has accepted the license terms.

.PARAMETER RemovePreviousVersions
    Adds <RemoveMSI /> to remove MSI-based Office (2016 and earlier) before
    installing. Interactively you are asked if this switch is not given;
    pass -RemovePreviousVersions:$false to skip the question.

.PARAMETER Architecture
    Overrides OfficeClientEdition (64 or 32).

.PARAMETER Language
    Replaces the Language IDs of every product, e.g. -Language en-us,fr-fr
    or -Language MatchOS.

.PARAMETER SourcePath
    Sets the Add SourcePath: where "/download" stores the installation files
    and where "/configure" installs from (local folder or \\server\share).

.PARAMETER DownloadOnly
    Only run "setup.exe /download" (for building offline media). Files are kept.

.PARAMETER SkipDownload
    Skip "setup.exe /download" and install directly from -SourcePath or the
    Office CDN.

.PARAMETER OdtSource
    Auto   (default) Use the latest ODT listed on the Microsoft Download Center,
           falling back to the pinned, hash-verified build if it cannot be found.
    Latest Use the latest ODT from the Download Center; fail if it cannot be found.
    Pinned Use the pinned, hash-verified ODT build only.

.PARAMETER OdtUrl
    Explicit HTTPS URL of an ODT executable. Its Microsoft signature is still verified.

.PARAMETER WorkingDirectory
    Parent folder for the OfficeDeploymentTool_<timestamp> folder.
    Default: the script folder (or %TEMP% if the script folder is unknown).

.PARAMETER LogDirectory
    Folder for the script transcript and ODT logs.
    Default: %ProgramData%\OfficeLTSC2024Installer\Logs

.PARAMETER Cleanup
    Ask (default interactively), Yes or No: delete the downloaded files after a
    successful installation. Files are always kept after a failure.

.PARAMETER MinimumFreeSpaceGB
    Minimum free space required on the system drive and the download drive. Default 10.

.PARAMETER DryRun
    Validate options and build the final configuration.xml, then print what
    would happen. Nothing is downloaded, installed or removed.

.EXAMPLE
    .\Install_Office.ps1
    Interactive installation (menu, prompts and confirmation).

.EXAMPLE
    .\Install_Office.ps1 -Unattended -Edition ProPlus -AcceptEula -RemovePreviousVersions
    Silent installation of Office LTSC Professional Plus 2024.

.EXAMPLE
    .\Install_Office.ps1 -Unattended -ConfigPath D:\Deploy\Office2024.xml -LogDirectory D:\Logs
    Silent installation from a custom configuration.

.EXAMPLE
    .\Install_Office.ps1 -Unattended -Edition Standard -DownloadOnly -SourcePath \\server\share\Office2024
    Stage offline installation files without installing.

.EXAMPLE
    .\Install_Office.ps1 -Unattended -Edition Standard -DryRun
    Show what an unattended installation would do, without making changes.

.NOTES
    Exit codes:
       0  Success (or dry run completed)
       1  Invalid arguments / input
       2  Configuration file missing or invalid
       3  Insufficient disk space
       4  Administrator rights required
       5  Office Deployment Tool download or verification failed
       6  Office Deployment Tool extraction failed
       7  Office download (setup.exe /download) failed
       8  Office installation (setup.exe /configure) failed
      10  Cancelled by the user
      99  Unexpected error

    Compatible with Windows PowerShell 5.1 and PowerShell 7.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [ValidateSet('Standard', 'ProPlus', 'Custom')]
    [string]$Edition,

    [string]$ConfigPath,

    [switch]$Unattended,

    [switch]$AcceptEula,

    [switch]$RemovePreviousVersions,

    [ValidateSet('64', '32')]
    [string]$Architecture,

    # Accepts an array or a comma-separated string (powershell.exe -File passes "a,b" as one string).
    [string[]]$Language,

    [string]$SourcePath,

    [switch]$DownloadOnly,

    [switch]$SkipDownload,

    [ValidateSet('Auto', 'Latest', 'Pinned')]
    [string]$OdtSource = 'Auto',

    [string]$OdtUrl,

    [string]$WorkingDirectory,

    [string]$LogDirectory,

    [ValidateSet('Ask', 'Yes', 'No')]
    [string]$Cleanup = 'Ask',

    [ValidateRange(0, 1024)]
    [int]$MinimumFreeSpaceGB = 10,

    [switch]$DryRun,

    # Backwards compatibility with the original "newurl" argument form:
    #   .\Install_Office.ps1 newurl <anything> <odt-url>
    [Parameter(ValueFromRemainingArguments = $true, DontShow = $true)]
    [string[]]$LegacyArguments
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Office LTSC 2024 channel and supported built-in editions.
$OfficeChannel = 'PerpetualVL2024'
$OfficeEditions = @{
    Standard = @{
        ProductId   = 'Standard2024Volume'
        FileName    = 'config_standard.xml'
        DisplayName = 'Office LTSC Standard 2024'
    }
    ProPlus  = @{
        ProductId   = 'ProPlus2024Volume'
        FileName    = 'config_ProPlus.xml'
        DisplayName = 'Office LTSC Professional Plus 2024'
    }
}

# Fallback location of the built-in configurations when the script is run on its own.
$ConfigBaseUrl = 'https://raw.githubusercontent.com/jdepew88/download_office_ltsc/main'

# Microsoft Download Center page for the Office Deployment Tool.
$OdtDownloadPage = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'

# Known-good ODT build (verified Microsoft signature, released 2026-11-17).
$PinnedOdt = @{
    Version = '16.0.20527.20150'
    Url     = 'https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_20527-20150.exe'
    Sha256  = '28BC762075C758B5EF841FE94487565168DEEED59F78DED5D5DFDFB08A9E827A'
}

$ExitCodes = @{
    Success           = 0
    InvalidArguments  = 1
    ConfigError       = 2
    DiskSpace         = 3
    NotAdministrator  = 4
    OdtDownload       = 5
    OdtExtract        = 6
    OfficeDownload    = 7
    OfficeInstall     = 8
    Cancelled         = 10
    Unexpected        = 99
}

# ---------------------------------------------------------------------------
# Output and error helpers
# ---------------------------------------------------------------------------

function Write-Step {
    param([string]$Message)
    Write-Host "[*] $Message"
}

function Write-Success {
    param([string]$Message)
    Write-Host "[+] $Message" -ForegroundColor Green
}

function Write-Failure {
    param([string]$Message)
    Write-Host "[!] $Message" -ForegroundColor Red
}

function New-InstallerError {
    # Creates an exception carrying the script exit code to return.
    param(
        [Parameter(Mandatory = $true)][int]$ExitCode,
        [Parameter(Mandatory = $true)][string]$Message
    )
    $exception = New-Object System.InvalidOperationException $Message
    $exception.Data['InstallerExitCode'] = $ExitCode
    return $exception
}

function Get-InstallerExitCode {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    $exception = $ErrorRecord.Exception
    while ($exception) {
        if ($exception.Data.Contains('InstallerExitCode')) {
            return [int]$exception.Data['InstallerExitCode']
        }
        $exception = $exception.InnerException
    }
    return $ExitCodes.Unexpected
}

function Read-InstallerInput {
    # Single entry point for every prompt, so unattended mode can never block.
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [bool]$Unattended
    )
    if ($Unattended) {
        throw (New-InstallerError -ExitCode $ExitCodes.Unexpected -Message "Internal error: a prompt was requested in unattended mode ($Prompt).")
    }
    try {
        return (Read-Host $Prompt)
    } catch {
        throw (New-InstallerError -ExitCode $ExitCodes.InvalidArguments -Message "Cannot prompt for input in this session. Run with -Unattended and -Edition (or -ConfigPath). Details: $($_.Exception.Message)")
    }
}

function Read-YesNo {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [bool]$Unattended
    )
    $answer = Read-InstallerInput -Prompt "$Prompt (Y/N)" -Unattended $Unattended
    return ([string]$answer).Trim() -match '^(?i:y|yes)$'
}

function Resolve-FullPath {
    # Resolves a path relative to the current PowerShell location (not the process directory).
    param([Parameter(Mandatory = $true)][string]$Path)
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal $identity
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ---------------------------------------------------------------------------
# Options
# ---------------------------------------------------------------------------

function Resolve-LegacyArguments {
    # Returns the ODT URL from the original "newurl" argument form, or $null.
    param([string[]]$Arguments)
    if (-not $Arguments -or $Arguments.Count -eq 0) {
        return $null
    }
    if ($Arguments[0] -eq 'newurl') {
        if ($Arguments.Count -ge 3) { return $Arguments[2] }
        if ($Arguments.Count -eq 2) { return $Arguments[1] }
    }
    throw (New-InstallerError -ExitCode $ExitCodes.InvalidArguments -Message "Unrecognised argument(s): $($Arguments -join ' '). Run Get-Help .\Install_Office.ps1 -Detailed for usage.")
}

function Read-EditionSelection {
    param([bool]$Unattended)
    Write-Host 'Select the Office LTSC 2024 edition to install:'
    Write-Host "1. $($OfficeEditions.Standard.DisplayName)"
    Write-Host "2. $($OfficeEditions.ProPlus.DisplayName)"
    Write-Host '3. Custom (place your Configuration.xml file in the same directory as this script)'
    $selection = ([string](Read-InstallerInput -Prompt 'Enter the number of your choice (1, 2, or 3)' -Unattended $Unattended)).Trim()
    switch ($selection) {
        '1' { return 'Standard' }
        '2' { return 'ProPlus' }
        '3' { return 'Custom' }
        default {
            throw (New-InstallerError -ExitCode $ExitCodes.InvalidArguments -Message "Invalid selection '$selection'. Exiting.")
        }
    }
}

function Resolve-InstallerOptions {
    param([Parameter(Mandatory = $true)][hashtable]$Settings)

    $options = @{} + $Settings

    $legacyUrl = Resolve-LegacyArguments -Arguments $options.LegacyArguments
    if ($legacyUrl) {
        if ($options.OdtUrl) {
            throw (New-InstallerError -ExitCode $ExitCodes.InvalidArguments -Message 'Specify the ODT URL either with -OdtUrl or with the legacy "newurl" form, not both.')
        }
        $options.OdtUrl = $legacyUrl
    }

    if ($options.OdtUrl) {
        $uri = $null
        if (-not [Uri]::TryCreate($options.OdtUrl, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') {
            throw (New-InstallerError -ExitCode $ExitCodes.InvalidArguments -Message "The ODT URL must be an absolute https:// URL: $($options.OdtUrl)")
        }
        Write-Step "Custom ODT URL provided: $($options.OdtUrl)"
    }

    if ($options.Language) {
        $languages = @($options.Language | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        foreach ($id in $languages) {
            if ($id -notmatch '^(?:[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})+|MatchOS|MatchInstalled|MatchPreviousMSI)$') {
                throw (New-InstallerError -ExitCode $ExitCodes.InvalidArguments -Message "Invalid language ID '$id'. Use IDs such as en-us, fr-fr or MatchOS.")
            }
        }
        $options.Language = $languages
    }

    if ($options.DownloadOnly -and $options.SkipDownload) {
        throw (New-InstallerError -ExitCode $ExitCodes.InvalidArguments -Message '-DownloadOnly and -SkipDownload cannot be used together.')
    }

    if ($options.ConfigPath) {
        if (-not $options.Edition) {
            $options.Edition = 'Custom'
        } elseif ($options.Edition -ne 'Custom') {
            throw (New-InstallerError -ExitCode $ExitCodes.InvalidArguments -Message '-ConfigPath can only be used with -Edition Custom.')
        }
    }

    if (-not $options.Edition) {
        if ($options.Unattended) {
            throw (New-InstallerError -ExitCode $ExitCodes.InvalidArguments -Message 'Unattended mode requires -Edition (Standard, ProPlus or Custom) or -ConfigPath.')
        }
        $options.Edition = Read-EditionSelection -Unattended $false
    }

    $scriptRoot = $options.ScriptRoot
    if (-not $scriptRoot) {
        $scriptRoot = (Get-Location).ProviderPath
    }

    if ($options.Edition -eq 'Custom') {
        if (-not $options.ConfigPath) {
            $options.ConfigPath = Join-Path $scriptRoot 'Configuration.xml'
        }
        $options.ConfigPath = Resolve-FullPath $options.ConfigPath
    }

    if (-not $options.WorkingDirectory) {
        if ($options.ScriptRoot) {
            $options.WorkingDirectory = $options.ScriptRoot
        } else {
            $options.WorkingDirectory = $env:TEMP
        }
    }
    $options.WorkingDirectory = Resolve-FullPath $options.WorkingDirectory

    if ($options.SourcePath -and -not $options.SourcePath.StartsWith('\\')) {
        $options.SourcePath = Resolve-FullPath $options.SourcePath
    }

    if ($options.LogDirectory) {
        $options.LogDirectory = Resolve-FullPath $options.LogDirectory
    }

    if ($options.Cleanup -eq 'Ask' -and $options.Unattended) {
        $options.Cleanup = 'Yes'
    }

    return $options
}

# ---------------------------------------------------------------------------
# Disk space
# ---------------------------------------------------------------------------

function Get-FreeSpaceGB {
    # Returns free space in GB for the drive holding $Path, or $null for UNC paths.
    param([Parameter(Mandatory = $true)][string]$Path)
    $root = [IO.Path]::GetPathRoot($Path)
    if (-not $root -or $root.StartsWith('\\')) {
        return $null
    }
    $drive = New-Object System.IO.DriveInfo $root
    if (-not $drive.IsReady) {
        return $null
    }
    return [math]::Round($drive.AvailableFreeSpace / 1GB, 2)
}

function Assert-FreeSpace {
    param(
        [Parameter(Mandatory = $true)][string[]]$Paths,
        [Parameter(Mandatory = $true)][int]$MinimumGB
    )
    $checked = @{}
    foreach ($path in $Paths) {
        $root = [IO.Path]::GetPathRoot($path)
        if (-not $root -or $checked.ContainsKey($root.ToUpperInvariant())) {
            continue
        }
        $checked[$root.ToUpperInvariant()] = $true
        $free = Get-FreeSpaceGB -Path $path
        if ($null -eq $free) {
            Write-Warning "Free space on $root could not be checked."
            continue
        }
        if ($free -lt $MinimumGB) {
            throw (New-InstallerError -ExitCode $ExitCodes.DiskSpace -Message "Insufficient disk space on $root ($free GB free). At least $MinimumGB GB of free space is required.")
        }
        Write-Step "Sufficient disk space detected on ${root}: $free GB available."
    }
}

# ---------------------------------------------------------------------------
# Configuration XML
# ---------------------------------------------------------------------------

function Import-ConfigurationDocument {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw (New-InstallerError -ExitCode $ExitCodes.ConfigError -Message "Configuration file not found: $Path")
    }

    $readerSettings = New-Object System.Xml.XmlReaderSettings
    $readerSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $readerSettings.XmlResolver = $null

    $document = New-Object System.Xml.XmlDocument
    $reader = $null
    try {
        $reader = [System.Xml.XmlReader]::Create((Resolve-FullPath $Path), $readerSettings)
        $document.Load($reader)
    } catch {
        throw (New-InstallerError -ExitCode $ExitCodes.ConfigError -Message "Configuration file is not valid XML: $Path. $($_.Exception.Message)")
    } finally {
        if ($reader) { $reader.Close() }
    }

    if (-not $document.DocumentElement -or $document.DocumentElement.Name -ne 'Configuration') {
        throw (New-InstallerError -ExitCode $ExitCodes.ConfigError -Message "Configuration file must have a <Configuration> root element: $Path")
    }

    # The leading comma stops PowerShell from enumerating the XmlDocument's child nodes.
    return , $document
}

function Get-ConfigurationSummary {
    param([Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Document)

    $root = $Document.DocumentElement
    $add = $root.SelectSingleNode('Add')
    $display = $root.SelectSingleNode('Display')

    $productIds = @()
    $languages = @()
    if ($add) {
        foreach ($product in $add.SelectNodes('Product')) {
            $productIds += $product.GetAttribute('ID')
            foreach ($languageNode in $product.SelectNodes('Language')) {
                $languages += $languageNode.GetAttribute('ID')
            }
        }
    }

    return [pscustomobject]@{
        HasAdd        = [bool]$add
        HasRemove     = [bool]$root.SelectSingleNode('Remove')
        Channel       = if ($add) { $add.GetAttribute('Channel') } else { '' }
        Architecture  = if ($add) { $add.GetAttribute('OfficeClientEdition') } else { '' }
        SourcePath    = if ($add) { $add.GetAttribute('SourcePath') } else { '' }
        ProductIds    = $productIds
        Languages     = @($languages | Select-Object -Unique)
        HasProductKey = ($root.SelectNodes('Add/Product[@PIDKEY]').Count -gt 0)
        RemoveMSI     = [bool]$root.SelectSingleNode('RemoveMSI')
        DisplayLevel  = if ($display) { $display.GetAttribute('Level') } else { '' }
        AcceptEULA    = if ($display) { $display.GetAttribute('AcceptEULA') } else { '' }
    }
}

function Assert-BuiltInConfiguration {
    # Built-in configurations must target the expected Office LTSC 2024 product and carry no key.
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Document,
        [Parameter(Mandatory = $true)][string]$Edition
    )
    $expected = $OfficeEditions[$Edition]
    $summary = Get-ConfigurationSummary -Document $Document

    if (-not $summary.HasAdd) {
        throw (New-InstallerError -ExitCode $ExitCodes.ConfigError -Message "Built-in configuration for $Edition has no <Add> element.")
    }
    if ($summary.Channel -ne $OfficeChannel) {
        throw (New-InstallerError -ExitCode $ExitCodes.ConfigError -Message "Built-in configuration for $Edition uses channel '$($summary.Channel)'; expected '$OfficeChannel'.")
    }
    if ($summary.ProductIds -notcontains $expected.ProductId) {
        throw (New-InstallerError -ExitCode $ExitCodes.ConfigError -Message "Built-in configuration for $Edition does not contain product '$($expected.ProductId)' (found: $($summary.ProductIds -join ', ')).")
    }
    if ($summary.HasProductKey) {
        throw (New-InstallerError -ExitCode $ExitCodes.ConfigError -Message "Built-in configuration for $Edition must not contain a PIDKEY. Use Activate_Office.ps1 for activation.")
    }
}

function Test-CustomConfiguration {
    # Custom configurations are flexible; only reject ones ODT cannot act on and warn about the rest.
    param([Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Document)
    $summary = Get-ConfigurationSummary -Document $Document

    if (-not $summary.HasAdd -and -not $summary.HasRemove) {
        throw (New-InstallerError -ExitCode $ExitCodes.ConfigError -Message 'Custom configuration contains neither an <Add> nor a <Remove> element.')
    }
    if ($summary.HasAdd) {
        if ($summary.Channel -ne $OfficeChannel) {
            Write-Warning "Custom configuration uses channel '$($summary.Channel)'. Office LTSC 2024 uses '$OfficeChannel'."
        }
        if (-not ($summary.ProductIds | Where-Object { $_ -like '*2024Volume' })) {
            Write-Warning "Custom configuration contains no Office LTSC 2024 volume product (found: $($summary.ProductIds -join ', '))."
        }
    }
    if ($summary.HasProductKey) {
        Write-Step 'Custom configuration contains a product key (PIDKEY); its value is not displayed or logged.'
    }
}

function Set-ConfigurationOptions {
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Document,
        [bool]$RemovePreviousVersions,
        [bool]$Unattended,
        [bool]$AcceptEula,
        [string]$Architecture,
        [string[]]$Language,
        [string]$SourcePath,
        [string]$LogDirectory
    )

    $root = $Document.DocumentElement
    $add = $root.SelectSingleNode('Add')

    if (($Architecture -or $Language -or $SourcePath) -and -not $add) {
        throw (New-InstallerError -ExitCode $ExitCodes.ConfigError -Message '-Architecture, -Language and -SourcePath require a configuration with an <Add> element.')
    }

    if ($Architecture) {
        $add.SetAttribute('OfficeClientEdition', $Architecture)
    }

    if ($SourcePath) {
        $add.SetAttribute('SourcePath', $SourcePath)
    }

    if ($Language) {
        foreach ($product in $add.SelectNodes('Product')) {
            foreach ($existing in @($product.SelectNodes('Language'))) {
                [void]$product.RemoveChild($existing)
            }
            $insertBefore = $product.FirstChild
            foreach ($id in $Language) {
                $languageNode = $Document.CreateElement('Language')
                $languageNode.SetAttribute('ID', $id)
                if ($insertBefore) {
                    [void]$product.InsertBefore($languageNode, $insertBefore)
                } else {
                    [void]$product.AppendChild($languageNode)
                }
            }
        }
    }

    if ($RemovePreviousVersions -and -not $root.SelectSingleNode('RemoveMSI')) {
        [void]$root.AppendChild($Document.CreateElement('RemoveMSI'))
    }

    if ($Unattended -or $AcceptEula) {
        $display = $root.SelectSingleNode('Display')
        if (-not $display) {
            $display = $Document.CreateElement('Display')
            [void]$root.AppendChild($display)
        }
        if ($Unattended) {
            $display.SetAttribute('Level', 'None')
        }
        if ($AcceptEula) {
            $display.SetAttribute('AcceptEULA', 'TRUE')
        } elseif (-not $display.HasAttribute('AcceptEULA')) {
            $display.SetAttribute('AcceptEULA', 'FALSE')
        }
    }

    if ($LogDirectory -and -not $root.SelectSingleNode('Logging')) {
        $logging = $Document.CreateElement('Logging')
        $logging.SetAttribute('Level', 'Standard')
        $logging.SetAttribute('Path', $LogDirectory)
        [void]$root.AppendChild($logging)
    }
}

function Save-ConfigurationDocument {
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Document,
        [Parameter(Mandatory = $true)][string]$Path
    )
    $writerSettings = New-Object System.Xml.XmlWriterSettings
    $writerSettings.Indent = $true
    $writerSettings.OmitXmlDeclaration = $true
    $writerSettings.Encoding = New-Object System.Text.UTF8Encoding $false
    $writerSettings.NewLineChars = "`r`n"
    $writer = [System.Xml.XmlWriter]::Create($Path, $writerSettings)
    try {
        $Document.Save($writer)
    } finally {
        $writer.Close()
    }
}

function Get-MaskedConfigurationText {
    # Returns the configuration as text with any product keys masked.
    param([Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Document)
    $copy = $Document.Clone()
    foreach ($product in $copy.SelectNodes('/Configuration/Add/Product[@PIDKEY]')) {
        $key = $product.GetAttribute('PIDKEY')
        $suffix = if ($key.Length -ge 5) { $key.Substring($key.Length - 5) } else { '' }
        $product.SetAttribute('PIDKEY', "*****-*****-*****-*****-$suffix")
    }
    $stringWriter = New-Object System.IO.StringWriter
    $writerSettings = New-Object System.Xml.XmlWriterSettings
    $writerSettings.Indent = $true
    $writerSettings.OmitXmlDeclaration = $true
    $xmlWriter = [System.Xml.XmlWriter]::Create($stringWriter, $writerSettings)
    try {
        $copy.Save($xmlWriter)
    } finally {
        $xmlWriter.Close()
    }
    return $stringWriter.ToString()
}

function Get-BuiltInConfigurationPath {
    # Uses the configuration shipped next to the script, or downloads it from the repository.
    param(
        [Parameter(Mandatory = $true)][string]$Edition,
        [string]$ScriptRoot,
        [Parameter(Mandatory = $true)][string]$DownloadDirectory
    )
    $fileName = $OfficeEditions[$Edition].FileName
    if ($ScriptRoot) {
        $localPath = Join-Path $ScriptRoot $fileName
        if (Test-Path -LiteralPath $localPath -PathType Leaf) {
            Write-Step "Using configuration file: $localPath"
            return $localPath
        }
    }

    $url = "$ConfigBaseUrl/$fileName"
    $destination = Join-Path $DownloadDirectory $fileName
    Write-Step "$fileName not found next to the script; downloading $url"
    try {
        Invoke-FileDownload -Uri $url -OutFile $destination
    } catch {
        throw (New-InstallerError -ExitCode $ExitCodes.ConfigError -Message "Failed to download $url. $($_.Exception.Message)")
    }
    return $destination
}

# ---------------------------------------------------------------------------
# Office Deployment Tool
# ---------------------------------------------------------------------------

function Invoke-FileDownload {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$OutFile
    )
    # Windows PowerShell 5.1 may not enable TLS 1.2 by default.
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $previousProgress = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'   # The progress bar slows downloads dramatically in 5.1.
    try {
        Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
    } finally {
        $ProgressPreference = $previousProgress
    }
}

function Get-LatestOdtUrlFromContent {
    # Extracts the ODT executable link from Download Center page content.
    param([string]$Content)
    if (-not $Content) {
        return $null
    }
    # Absolute links keep their host (validated by the caller); href="/download/..." links are page-relative.
    $pattern = '(?i)(?:https?://[A-Za-z0-9.-]+)?/download/[^"''\s<>]*?/officedeploymenttool_[0-9-]+\.exe'
    foreach ($match in [regex]::Matches($Content, $pattern)) {
        $url = $match.Value
        if ($url.StartsWith('/')) {
            # Page-relative link: resolve against the Download Center page (the original script's bug).
            $url = (New-Object Uri ((New-Object Uri $OdtDownloadPage), $url)).AbsoluteUri
        }
        return $url
    }
    return $null
}

function Get-OdtDownloadPlan {
    # Decides which ODT executable to download. Returns Url, Sha256 (optional) and Description.
    param(
        [Parameter(Mandatory = $true)][string]$OdtSource,
        [string]$OdtUrl
    )
    $pinned = @{ Url = $PinnedOdt.Url; Sha256 = $PinnedOdt.Sha256; Description = "pinned build $($PinnedOdt.Version)" }

    if ($OdtUrl) {
        return @{ Url = $OdtUrl; Sha256 = $null; Description = 'custom URL' }
    }
    if ($OdtSource -eq 'Pinned') {
        return $pinned
    }

    $latest = $null
    $failure = $null
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $page = Invoke-WebRequest -Uri $OdtDownloadPage -UseBasicParsing -ErrorAction Stop
        $latest = Get-LatestOdtUrlFromContent -Content $page.Content
    } catch {
        $failure = $_.Exception.Message
    }

    if ($latest) {
        $uri = New-Object Uri $latest
        if ($uri.Scheme -ne 'https' -or -not ($uri.Host -eq 'microsoft.com' -or $uri.Host.EndsWith('.microsoft.com'))) {
            $failure = "unexpected download location $latest"
            $latest = $null
        }
    }

    if ($latest) {
        Write-Step "Latest ODT found on the Microsoft Download Center: $latest"
        if ($latest -eq $PinnedOdt.Url) {
            return $pinned
        }
        return @{ Url = $latest; Sha256 = $null; Description = 'latest from the Microsoft Download Center' }
    }

    if (-not $failure) {
        $failure = 'no ODT download link found on the page'
    }
    if ($OdtSource -eq 'Latest') {
        throw (New-InstallerError -ExitCode $ExitCodes.OdtDownload -Message "Could not determine the latest ODT from $OdtDownloadPage ($failure).")
    }
    Write-Warning "Could not determine the latest ODT ($failure). Using the $($pinned.Description)."
    return $pinned
}

function Test-MicrosoftSignature {
    # True when the file has a valid Authenticode signature from Microsoft Corporation.
    param([Parameter(Mandatory = $true)][string]$Path)
    $signature = Get-AuthenticodeSignature -FilePath $Path
    if ($signature.Status -ne [System.Management.Automation.SignatureStatus]::Valid -or -not $signature.SignerCertificate) {
        return $false
    }
    return ($signature.SignerCertificate.Subject -match '(^|,\s*)O=Microsoft Corporation(,|$)')
}

function Get-OfficeDeploymentTool {
    # Downloads, verifies and extracts the ODT. Returns the path of setup.exe.
    param(
        [Parameter(Mandatory = $true)][hashtable]$Plan,
        [Parameter(Mandatory = $true)][string]$Directory
    )
    $installer = Join-Path $Directory 'officedeploymenttool.exe'

    Write-Step "Downloading the Office Deployment Tool ($($Plan.Description)): $($Plan.Url)"
    try {
        Invoke-FileDownload -Uri $Plan.Url -OutFile $installer
    } catch {
        throw (New-InstallerError -ExitCode $ExitCodes.OdtDownload -Message "Failed to download the Office Deployment Tool. $($_.Exception.Message)")
    }
    if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) {
        throw (New-InstallerError -ExitCode $ExitCodes.OdtDownload -Message 'Office Deployment Tool download failed. File not found.')
    }

    if ($Plan.Sha256) {
        $hash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash
        if ($hash -ne $Plan.Sha256) {
            throw (New-InstallerError -ExitCode $ExitCodes.OdtDownload -Message "Office Deployment Tool hash mismatch (expected $($Plan.Sha256), got $hash).")
        }
    }
    if (-not (Test-MicrosoftSignature -Path $installer)) {
        throw (New-InstallerError -ExitCode $ExitCodes.OdtDownload -Message 'The downloaded Office Deployment Tool is not validly signed by Microsoft Corporation.')
    }
    Write-Success 'Downloaded the Office Deployment Tool and verified its Microsoft signature.'

    try {
        $process = Start-Process -FilePath $installer -ArgumentList @('/quiet', "/extract:`"$Directory`"") -Wait -PassThru
    } catch {
        throw (New-InstallerError -ExitCode $ExitCodes.OdtExtract -Message "Failed to extract the Office Deployment Tool. $($_.Exception.Message)")
    }
    $setup = Join-Path $Directory 'setup.exe'
    if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $setup -PathType Leaf)) {
        throw (New-InstallerError -ExitCode $ExitCodes.OdtExtract -Message "Extraction failed (exit code $($process.ExitCode)); setup.exe not found.")
    }
    if (-not (Test-MicrosoftSignature -Path $setup)) {
        throw (New-InstallerError -ExitCode $ExitCodes.OdtExtract -Message 'The extracted setup.exe is not validly signed by Microsoft Corporation.')
    }
    Write-Success "Extracted the Office Deployment Tool (setup.exe $((Get-Item -LiteralPath $setup).VersionInfo.FileVersion))."
    return $setup
}

function Invoke-OdtSetup {
    # Runs setup.exe in download or configure mode and returns its exit code.
    param(
        [Parameter(Mandatory = $true)][string]$SetupPath,
        [Parameter(Mandatory = $true)][ValidateSet('download', 'configure')][string]$Mode,
        [Parameter(Mandatory = $true)][string]$ConfigFile
    )
    $process = Start-Process -FilePath $SetupPath -ArgumentList @("/$Mode", "`"$ConfigFile`"") `
        -WorkingDirectory (Split-Path -Parent $SetupPath) -NoNewWindow -Wait -PassThru
    return [int]$process.ExitCode
}

function Test-OfficeProductInstalled {
    # Checks the Click-to-Run registration for a product ID after installation.
    param([Parameter(Mandatory = $true)][string]$ProductId)
    $key = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $item = Get-ItemProperty -Path $key -Name 'ProductReleaseIds' -ErrorAction SilentlyContinue
    if (-not $item) {
        return $false
    }
    return (($item.ProductReleaseIds -split ',') -contains $ProductId)
}

# ---------------------------------------------------------------------------
# Main flow
# ---------------------------------------------------------------------------

function Invoke-Main {
    param([Parameter(Mandatory = $true)][hashtable]$Settings)

    $outputDir = $null
    $keepOutput = $false
    try {
        $options = Resolve-InstallerOptions -Settings $Settings
        $isCustom = ($options.Edition -eq 'Custom')
        if ($isCustom) {
            $displayName = 'Office LTSC 2024 (custom configuration)'
        } else {
            $displayName = $OfficeEditions[$options.Edition].DisplayName
        }
        $runConfigure = -not $options.DownloadOnly
        $runDownload = -not $options.SkipDownload

        Write-Step "Edition: $displayName"
        if ($options.Unattended) { Write-Step 'Mode: unattended' }
        if ($options.DryRun) { Write-Step 'Mode: dry run (nothing will be downloaded, installed or removed)' }

        # Elevation is required to install Office.
        if ($runConfigure -and -not (Test-IsAdministrator)) {
            if ($options.DryRun) {
                Write-Warning 'Not running as Administrator; a real installation would stop here (exit code 4).'
            } else {
                throw (New-InstallerError -ExitCode $ExitCodes.NotAdministrator -Message 'Installing Office requires an elevated PowerShell session. Right-click PowerShell and choose "Run as administrator".')
            }
        }

        if ($isCustom -and -not (Test-Path -LiteralPath $options.ConfigPath -PathType Leaf)) {
            throw (New-InstallerError -ExitCode $ExitCodes.ConfigError -Message "Custom configuration file not found: $($options.ConfigPath). Place your Configuration.xml in the same directory as this script or pass -ConfigPath.")
        }

        # Ask about previous versions when not decided on the command line.
        $removePrevious = $options.RemovePreviousVersions
        if ($runConfigure -and $null -eq $removePrevious) {
            if ($options.Unattended) {
                $removePrevious = $false
            } else {
                Write-Host 'NOTE: Removing previous versions of Office is optional.'
                Write-Host 'This removes Windows Installer (MSI) versions of Office, such as Office 2016 and earlier.'
                Write-Host 'You may want to retain the existing installation if:'
                Write-Host '1. You are performing an update and want to ensure settings and customizations are preserved.'
                Write-Host '2. You are installing a different edition or version side-by-side with an existing Office installation.'
                Write-Host '3. You are troubleshooting or comparing features between different Office versions.'
                $removePrevious = Read-YesNo -Prompt 'Do you want to remove previous versions of Office on this computer?' -Unattended $false
            }
        }
        $removePrevious = [bool]$removePrevious

        # Disk space: download location and system drive (install location).
        $spacePaths = @()
        if ($runDownload) {
            if ($options.SourcePath) { $spacePaths += $options.SourcePath } else { $spacePaths += $options.WorkingDirectory }
        }
        if ($runConfigure) { $spacePaths += "$env:SystemDrive\" }
        if ($spacePaths.Count -gt 0 -and $options.MinimumFreeSpaceGB -gt 0) {
            Assert-FreeSpace -Paths $spacePaths -MinimumGB $options.MinimumFreeSpaceGB
        }

        # Unique working folder for this run.
        if (-not (Test-Path -LiteralPath $options.WorkingDirectory -PathType Container)) {
            New-Item -ItemType Directory -Path $options.WorkingDirectory -Force | Out-Null
        }
        $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
        $outputDir = Join-Path $options.WorkingDirectory "OfficeDeploymentTool_$timestamp"
        try {
            New-Item -ItemType Directory -Path $outputDir -ErrorAction Stop | Out-Null
        } catch {
            $outputDir = $null
            throw (New-InstallerError -ExitCode $ExitCodes.InvalidArguments -Message "Failed to create the working directory under $($options.WorkingDirectory). $($_.Exception.Message)")
        }
        Write-Step "Created working directory: $outputDir"

        if ($options.LogDirectory -and -not (Test-Path -LiteralPath $options.LogDirectory -PathType Container)) {
            New-Item -ItemType Directory -Path $options.LogDirectory -Force | Out-Null
        }

        # Build and validate the configuration.
        if ($isCustom) {
            Write-Step "Using custom configuration: $($options.ConfigPath)"
            $configDocument = Import-ConfigurationDocument -Path $options.ConfigPath
            Test-CustomConfiguration -Document $configDocument
        } else {
            $sourceConfig = Get-BuiltInConfigurationPath -Edition $options.Edition -ScriptRoot $options.ScriptRoot -DownloadDirectory $outputDir
            $configDocument = Import-ConfigurationDocument -Path $sourceConfig
            Assert-BuiltInConfiguration -Document $configDocument -Edition $options.Edition
        }

        Set-ConfigurationOptions -Document $configDocument `
            -RemovePreviousVersions $removePrevious `
            -Unattended $options.Unattended `
            -AcceptEula $options.AcceptEula `
            -Architecture $options.Architecture `
            -Language $options.Language `
            -SourcePath $options.SourcePath `
            -LogDirectory $options.LogDirectory

        $configFile = Join-Path $outputDir 'configuration.xml'
        Save-ConfigurationDocument -Document $configDocument -Path $configFile
        # Re-read the saved file to prove what ODT will receive is well-formed.
        $null = Import-ConfigurationDocument -Path $configFile

        $summary = Get-ConfigurationSummary -Document $configDocument
        Write-Step "Configuration written to $configFile"
        Write-Step ("Products: {0}; channel: {1}; architecture: {2}-bit; languages: {3}" -f ($summary.ProductIds -join ', '), $summary.Channel, $summary.Architecture, ($summary.Languages -join ', '))
        if ($runConfigure) {
            if ($summary.RemoveMSI) {
                Write-Step 'Previous MSI versions of Office will be removed (<RemoveMSI />).'
            } else {
                Write-Step 'Previous Office versions will not be removed.'
            }
        }

        if ($options.DryRun) {
            Write-Host ''
            Write-Host 'Final configuration.xml (product keys masked):'
            Write-Host (Get-MaskedConfigurationText -Document $configDocument)
            Write-Host ''
            Write-Step 'Dry run: the following steps would be performed:'
            if ($options.OdtUrl) {
                Write-Host "  - Download the ODT from $($options.OdtUrl) and verify its Microsoft signature"
            } elseif ($options.OdtSource -eq 'Pinned') {
                Write-Host "  - Download the pinned ODT $($PinnedOdt.Version) and verify its SHA-256 hash and Microsoft signature"
            } else {
                Write-Host "  - Resolve the latest ODT from $OdtDownloadPage (OdtSource=$($options.OdtSource)), then verify its Microsoft signature"
            }
            if ($runDownload) { Write-Host "  - setup.exe /download `"$configFile`"" }
            if ($runConfigure) { Write-Host "  - setup.exe /configure `"$configFile`"" }
            Write-Host "  - Cleanup: $($options.Cleanup)"
            Write-Success 'Dry run completed. No changes were made to Office on this computer.'
            $keepOutput = $true
            return $ExitCodes.Success
        }

        if (-not $options.Unattended) {
            if ($runConfigure) { $action = 'install' } else { $action = 'download' }
            if (-not (Read-YesNo -Prompt "Ready to $action $displayName. Do you want to proceed?" -Unattended $false)) {
                Write-Step 'Installation canceled by the user.'
                return $ExitCodes.Cancelled
            }
        }

        $plan = Get-OdtDownloadPlan -OdtSource $options.OdtSource -OdtUrl $options.OdtUrl
        $setup = Get-OfficeDeploymentTool -Plan $plan -Directory $outputDir

        if ($runDownload) {
            Write-Step 'Running setup.exe /download (this can take a while; there is no progress output)...'
            $code = Invoke-OdtSetup -SetupPath $setup -Mode 'download' -ConfigFile $configFile
            if ($code -ne 0) {
                throw (New-InstallerError -ExitCode $ExitCodes.OfficeDownload -Message "setup.exe /download failed with exit code $code.")
            }
            Write-Success 'Office LTSC 2024 files have been downloaded successfully.'
        }

        if (-not $runConfigure) {
            $keepOutput = $true
            if ($options.SourcePath) { $location = $options.SourcePath } else { $location = $outputDir }
            Write-Success "Download completed. Installation files are in: $location"
            return $ExitCodes.Success
        }

        Write-Step 'Running setup.exe /configure...'
        $code = Invoke-OdtSetup -SetupPath $setup -Mode 'configure' -ConfigFile $configFile
        if ($code -ne 0) {
            throw (New-InstallerError -ExitCode $ExitCodes.OfficeInstall -Message "setup.exe /configure failed with exit code $code.")
        }
        Write-Success 'Office installation completed successfully.'

        foreach ($productId in $summary.ProductIds) {
            if ($productId -like '*Volume' -and -not (Test-OfficeProductInstalled -ProductId $productId)) {
                Write-Warning "$productId was not found in the Click-to-Run product list; check the ODT logs."
            }
        }

        Write-Step 'Activation: volume licensed Office LTSC 2024 uses a Generic Volume License Key and activates automatically against your KMS host or Active Directory.'
        Write-Step 'For MAK activation or to set a KMS host, run Activate_Office.ps1 (see README).'

        # Cleanup after a successful installation.
        $doCleanup = $false
        switch ($options.Cleanup) {
            'Yes' { $doCleanup = $true }
            'No' { $doCleanup = $false }
            default {
                Write-Host 'IMPORTANT: Please ensure that the Office setup window is closed before proceeding.'
                $doCleanup = Read-YesNo -Prompt 'Do you want to delete the downloaded files to free up space?' -Unattended $false
            }
        }
        if ($doCleanup) {
            Remove-Item -LiteralPath $outputDir -Recurse -Force
            Write-Step 'Downloaded files have been deleted.'
            $outputDir = $null
        } else {
            $keepOutput = $true
            Write-Step "Downloaded files have been retained in $outputDir"
        }

        return $ExitCodes.Success
    } catch {
        $exitCode = Get-InstallerExitCode -ErrorRecord $_
        Write-Failure $_.Exception.Message
        if ($exitCode -eq $ExitCodes.Unexpected) {
            Write-Failure $_.ScriptStackTrace
        }
        if ($outputDir -and (Test-Path -LiteralPath $outputDir)) {
            $keepOutput = $true
            Write-Step "Files retained for troubleshooting in $outputDir"
        }
        Write-Step "Exiting with code $exitCode."
        return $exitCode
    } finally {
        # Remove an unused working folder after cancellation.
        if (-not $keepOutput -and $outputDir -and (Test-Path -LiteralPath $outputDir)) {
            Remove-Item -LiteralPath $outputDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if ($PSBoundParameters.ContainsKey('RemovePreviousVersions')) {
    $removePreviousSetting = [bool]$RemovePreviousVersions
} else {
    $removePreviousSetting = $null
}
if ($LogDirectory) {
    $logDirectorySetting = $LogDirectory
} else {
    $logDirectorySetting = Join-Path $env:ProgramData 'OfficeLTSC2024Installer\Logs'
}

$InstallerSettings = @{
    Edition                = $Edition
    ConfigPath             = $ConfigPath
    Unattended             = [bool]$Unattended
    AcceptEula             = [bool]$AcceptEula
    RemovePreviousVersions = $removePreviousSetting
    Architecture           = $Architecture
    Language               = $Language
    SourcePath             = $SourcePath
    DownloadOnly           = [bool]$DownloadOnly
    SkipDownload           = [bool]$SkipDownload
    OdtSource              = $OdtSource
    OdtUrl                 = $OdtUrl
    WorkingDirectory       = $WorkingDirectory
    LogDirectory           = $logDirectorySetting
    Cleanup                = $Cleanup
    MinimumFreeSpaceGB     = $MinimumFreeSpaceGB
    DryRun                 = [bool]$DryRun
    LegacyArguments        = $LegacyArguments
    ScriptRoot             = $PSScriptRoot
}

# When dot-sourced (for example by the tests), only load the functions and settings.
if ($MyInvocation.InvocationName -eq '.') {
    return
}

$transcriptStarted = $false
try {
    if (-not (Test-Path -LiteralPath $logDirectorySetting -PathType Container)) {
        New-Item -ItemType Directory -Path $logDirectorySetting -Force | Out-Null
    }
    $transcriptPath = Join-Path $logDirectorySetting ("Install_Office_{0}.log" -f (Get-Date -Format 'yyyyMMddHHmmss'))
    Start-Transcript -Path $transcriptPath -Append | Out-Null
    $transcriptStarted = $true
} catch {
    Write-Warning "Could not start a transcript in ${logDirectorySetting}: $($_.Exception.Message)"
}

try {
    $installerExitCode = Invoke-Main -Settings $InstallerSettings
} finally {
    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }
}
exit $installerExitCode
