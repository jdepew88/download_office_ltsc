<#
.SYNOPSIS
    Tests for Install_Office.ps1, Activate_Office.ps1, install_office_nodefault.PS1
    and the Office LTSC 2024 configuration files.

.DESCRIPTION
    Safe to run on any computer: Office is never downloaded, installed, uninstalled
    or activated.
      - Child-process tests always use -DryRun (installer/activation) or fail
        before any action is taken.
      - In-process tests dot-source Install_Office.ps1 and replace setup.exe, the ODT
        download, Start-Process, Invoke-WebRequest and Read-Host with stubs.

    -Online additionally downloads the current Office Deployment Tool from Microsoft
    into a temporary folder, verifies its signature and extracts it (setup.exe is
    never run).

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-OfficeInstaller.ps1
#>
[CmdletBinding()]
param(
    [switch]$Online
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$RepoRoot = Split-Path -Parent $PSScriptRoot
$InstallerPath = Join-Path $RepoRoot 'Install_Office.ps1'
$ActivatorPath = Join-Path $RepoRoot 'Activate_Office.ps1'
$NoDefaultPath = Join-Path $RepoRoot 'install_office_nodefault.PS1'
$TestRoot = Join-Path ([IO.Path]::GetTempPath()) ("OfficeInstallerTests_{0}" -f [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $TestRoot | Out-Null
$LogDir = Join-Path $TestRoot 'logs'

# A syntactically valid but fake product key, built at run time so no key-shaped
# literal exists anywhere in the repository.
$FakeKey = (@('TESTA', 'TESTB', 'TESTC', 'TESTD', 'FAKE9') -join '-')
$KeyPattern = '\b[A-Z0-9]{5}(-[A-Z0-9]{5}){4}\b'

$Results = New-Object System.Collections.Generic.List[object]

# ---------------------------------------------------------------------------
# Minimal test framework
# ---------------------------------------------------------------------------

function Test-Case {
    param([string]$Name, [scriptblock]$Body)
    try {
        & $Body
        $Results.Add([pscustomobject]@{ Name = $Name; Passed = $true; Message = '' })
        Write-Host "  PASS  $Name" -ForegroundColor Green
    } catch {
        $Results.Add([pscustomobject]@{ Name = $Name; Passed = $false; Message = $_.Exception.Message })
        Write-Host "  FAIL  $Name" -ForegroundColor Red
        Write-Host "        $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Assert-True {
    param($Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    if ($Expected -ne $Actual) { throw "Assertion failed: $Message (expected '$Expected', got '$Actual')" }
}

function New-TestDirectory {
    $path = Join-Path $TestRoot ([guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Path $path | Out-Null
    return $path
}

function New-TestFile {
    param([string]$Content, [string]$Name = 'Configuration.xml')
    $path = Join-Path (New-TestDirectory) $Name
    [IO.File]::WriteAllText($path, $Content)
    return $path
}

function Invoke-ScriptProcess {
    # Runs a script in a separate, non-interactive Windows PowerShell process.
    param([string]$Path, [string[]]$Arguments)
    # Parameter-binding errors arrive on stderr; do not let them throw here.
    $ErrorActionPreference = 'Continue'
    $output = & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Path @Arguments 2>&1 | Out-String
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $output }
}

function Get-OdtFolder {
    param([string]$WorkingDirectory)
    return @(Get-ChildItem -LiteralPath $WorkingDirectory -Directory -Filter 'OfficeDeploymentTool_*')
}

function Import-Xml {
    param([string]$Path)
    $document = New-Object System.Xml.XmlDocument
    $document.Load($Path)
    return , $document
}

function Invoke-StubbedInstaller {
    <#
      Dot-sources Install_Office.ps1 with the given parameters, replaces every
      external action with a stub and runs Invoke-Main.
    #>
    param(
        [hashtable]$Parameters,
        [hashtable]$SetupExitCodes = @{ download = 0; configure = 0 },
        [bool]$IsAdmin = $true,
        [string[]]$Answers = @()
    )
    $calls = New-Object System.Collections.Generic.List[string]
    $configs = New-Object System.Collections.Generic.List[string]
    $answerQueue = New-Object System.Collections.Queue
    foreach ($answer in $Answers) { $answerQueue.Enqueue($answer) }

    if (-not $Parameters.ContainsKey('LogDirectory')) { $Parameters.LogDirectory = $LogDir }
    if (-not $Parameters.ContainsKey('MinimumFreeSpaceGB')) { $Parameters.MinimumFreeSpaceGB = 0 }

    . $InstallerPath @Parameters

    # Safety net: nothing in the tests may start a process or use the network.
    function Start-Process { throw 'Start-Process must not be called by the tests.' }
    function Invoke-WebRequest { throw 'Invoke-WebRequest must not be called by the tests.' }

    function Test-IsAdministrator { return $IsAdmin }
    function Get-OdtDownloadPlan {
        param($OdtSource, $OdtUrl)
        $calls.Add("plan:$OdtSource")
        return @{ Url = 'https://download.microsoft.com/stub/officedeploymenttool_0-0.exe'; Sha256 = $null; Description = 'stub' }
    }
    function Get-OfficeDeploymentTool {
        param($Plan, $Directory)
        $calls.Add('odt')
        $setup = Join-Path $Directory 'setup.exe'
        Set-Content -LiteralPath $setup -Value 'stub'
        return $setup
    }
    function Invoke-OdtSetup {
        param($SetupPath, $Mode, $ConfigFile)
        $calls.Add("setup:$Mode")
        $configs.Add([IO.File]::ReadAllText($ConfigFile))
        return [int]$SetupExitCodes[$Mode]
    }
    function Test-OfficeProductInstalled { param($ProductId) return $true }
    function Read-Host {
        param($Prompt)
        $calls.Add("prompt:$Prompt")
        if ($answerQueue.Count -eq 0) { throw "Unexpected prompt: $Prompt" }
        return $answerQueue.Dequeue()
    }

    $exitCode = Invoke-Main -Settings $InstallerSettings 6>$null 3>$null
    return [pscustomobject]@{
        ExitCode = $exitCode
        Calls    = @($calls)
        Configs  = @($configs)
    }
}

# Dot-source this scriptblock (". $ImportInstaller") to load the installer's
# functions into the current scope; the dot-source guard stops anything from running.
$ImportInstaller = { . $InstallerPath -LogDirectory $LogDir }

Write-Host "Office LTSC 2024 installer tests"
Write-Host "Repository: $RepoRoot"
Write-Host "PowerShell: $($PSVersionTable.PSVersion)"
Write-Host "Temp:       $TestRoot"
Write-Host ''

try {

# ---------------------------------------------------------------------------
Write-Host 'PowerShell syntax'
# ---------------------------------------------------------------------------

foreach ($scriptFile in @(Get-ChildItem -LiteralPath $RepoRoot -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1' })) {
    Test-Case "Parses without errors: $($scriptFile.FullName.Substring($RepoRoot.Length + 1))" {
        $tokens = $null
        $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($scriptFile.FullName, [ref]$tokens, [ref]$errors)
        Assert-Equal 0 $errors.Count ("parse errors: " + (($errors | ForEach-Object { "line $($_.Extent.StartLineNumber): $($_.Message)" }) -join '; '))
    }
}

Test-Case 'PSScriptAnalyzer reports no errors (skipped if not installed)' {
    if (-not (Get-Module -ListAvailable -Name PSScriptAnalyzer)) {
        Write-Host '        (PSScriptAnalyzer not installed - skipped)' -ForegroundColor DarkGray
        return
    }
    $findings = @(Invoke-ScriptAnalyzer -Path $RepoRoot -Recurse -Severity Error)
    Assert-Equal 0 $findings.Count (($findings | ForEach-Object { "$($_.ScriptName):$($_.Line) $($_.RuleName)" }) -join '; ')
}

Test-Case 'Every script parameter is forwarded in $InstallerSettings' {
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($InstallerPath, [ref]$tokens, [ref]$errors)
    $names = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
    . $ImportInstaller
    foreach ($name in $names) {
        Assert-True $InstallerSettings.ContainsKey($name) "parameter $name missing from InstallerSettings"
    }
}

Test-Case 'Script help documents every parameter' {
    $help = Get-Help $InstallerPath -Full
    $documented = @($help.parameters.parameter | Where-Object { $_.description } | ForEach-Object { $_.name })
    foreach ($name in 'Edition', 'ConfigPath', 'Unattended', 'AcceptEula', 'RemovePreviousVersions', 'Architecture', 'Language',
        'SourcePath', 'DownloadOnly', 'SkipDownload', 'OdtSource', 'OdtUrl', 'WorkingDirectory', 'LogDirectory', 'Cleanup',
        'MinimumFreeSpaceGB', 'DryRun') {
        Assert-True ($documented -contains $name) "parameter $name has no help text"
    }
}

# ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'Configuration files'
# ---------------------------------------------------------------------------

$expectedConfigs = @(
    @{ File = 'config_standard.xml'; ProductId = 'Standard2024Volume' },
    @{ File = 'config_ProPlus.xml'; ProductId = 'ProPlus2024Volume' }
)
foreach ($expected in $expectedConfigs) {
    $path = Join-Path $RepoRoot $expected.File

    Test-Case "$($expected.File): well-formed Office LTSC 2024 configuration" {
        $xml = Import-Xml $path
        Assert-Equal 'Configuration' $xml.DocumentElement.Name 'root element'
        $adds = $xml.SelectNodes('/Configuration/Add')
        Assert-Equal 1 $adds.Count 'number of <Add> elements'
        Assert-Equal 'PerpetualVL2024' $adds[0].GetAttribute('Channel') 'channel'
        Assert-Equal '64' $adds[0].GetAttribute('OfficeClientEdition') 'architecture'
        $products = @($xml.SelectNodes('/Configuration/Add/Product') | ForEach-Object { $_.GetAttribute('ID') })
        Assert-Equal 1 $products.Count 'number of products'
        Assert-Equal $expected.ProductId $products[0] 'product ID'
        $languages = @($xml.SelectNodes('/Configuration/Add/Product/Language') | ForEach-Object { $_.GetAttribute('ID') })
        Assert-True ($languages -contains 'en-us') 'en-us language present'
        Assert-Equal 'TRUE' $xml.SelectSingleNode('/Configuration/Updates').GetAttribute('Enabled') 'updates enabled'
    }

    Test-Case "$($expected.File): no product key, activation or MSI removal by default" {
        $raw = [IO.File]::ReadAllText($path)
        $xml = Import-Xml $path
        Assert-Equal 0 $xml.SelectNodes('//*[@PIDKEY]').Count 'PIDKEY attributes'
        Assert-True (-not ($raw -cmatch $KeyPattern)) 'no product-key-shaped text'
        Assert-Equal 0 $xml.SelectNodes("//Property[@Name='AUTOACTIVATE']").Count 'AUTOACTIVATE property'
        Assert-Equal 0 $xml.SelectNodes("//Property[@Name='SharedComputerLicensing' or @Name='SCLCacheOverride']").Count 'shared computer properties (not applicable to LTSC 2024)'
        Assert-Equal 0 $xml.SelectNodes('/Configuration/RemoveMSI').Count '<RemoveMSI /> (opt-in only)'
        Assert-Equal 0 $xml.SelectNodes('/Configuration/Display').Count '<Display /> (added at run time)'
        Assert-True (-not ($raw -match '2021')) 'no Office 2021 references'
    }

    Test-Case "$($expected.File): accepted by the installer's built-in validation" {
        . $ImportInstaller
        $edition = if ($expected.ProductId -like 'Standard*') { 'Standard' } else { 'ProPlus' }
        $document = Import-ConfigurationDocument -Path $path
        Assert-BuiltInConfiguration -Document $document -Edition $edition
    }
}

Test-Case 'No product keys anywhere in the repository' {
    $files = @(Get-ChildItem -LiteralPath $RepoRoot -Recurse -File | Where-Object { $_.FullName -notmatch '\\\.git(\\|$)' })
    foreach ($file in $files) {
        $text = [IO.File]::ReadAllText($file.FullName)
        Assert-True (-not ($text -cmatch $KeyPattern)) "product-key-shaped text in $($file.Name)"
    }
}

# ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'Configuration processing'
# ---------------------------------------------------------------------------

$legacy2021 = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume" PIDKEY="$FakeKey"><Language ID="en-us" /></Product>
  </Add>
</Configuration>
"@

Test-Case 'Built-in validation rejects an Office 2021 configuration (exit code 2)' {
    . $ImportInstaller
    $document = Import-ConfigurationDocument -Path (New-TestFile $legacy2021)
    try { Assert-BuiltInConfiguration -Document $document -Edition ProPlus; throw 'not rejected' }
    catch { Assert-Equal 2 (Get-InstallerExitCode $_) "exit code ($($_.Exception.Message))" }
}

Test-Case 'Built-in validation rejects the wrong 2024 product' {
    . $ImportInstaller
    $document = Import-ConfigurationDocument -Path (Join-Path $RepoRoot 'config_standard.xml')
    try { Assert-BuiltInConfiguration -Document $document -Edition ProPlus; throw 'not rejected' }
    catch { Assert-Equal 2 (Get-InstallerExitCode $_) "exit code ($($_.Exception.Message))" }
}

Test-Case 'Built-in validation rejects a PIDKEY' {
    . $ImportInstaller
    $document = Import-ConfigurationDocument -Path (Join-Path $RepoRoot 'config_ProPlus.xml')
    $document.SelectSingleNode('//Product').SetAttribute('PIDKEY', $FakeKey)
    try { Assert-BuiltInConfiguration -Document $document -Edition ProPlus; throw 'not rejected' }
    catch { Assert-Equal 2 (Get-InstallerExitCode $_) "exit code ($($_.Exception.Message))" }
}

foreach ($bad in @(
        @{ Name = 'malformed XML'; Content = '<Configuration><Add></Configuration>' },
        @{ Name = 'wrong root element'; Content = '<Config><Add Channel="PerpetualVL2024" /></Config>' },
        @{ Name = 'DTD / external entity'; Content = '<?xml version="1.0"?><!DOCTYPE c [<!ENTITY x SYSTEM "file:///c:/windows/win.ini">]><Configuration>&x;</Configuration>' },
        @{ Name = 'empty file'; Content = '' }
    )) {
    Test-Case "Import rejects $($bad.Name) (exit code 2)" {
        . $ImportInstaller
        try { $null = Import-ConfigurationDocument -Path (New-TestFile $bad.Content); throw 'not rejected' }
        catch { Assert-Equal 2 (Get-InstallerExitCode $_) "exit code ($($_.Exception.Message))" }
    }
}

Test-Case 'Custom validation rejects a configuration with neither Add nor Remove' {
    . $ImportInstaller
    $document = Import-ConfigurationDocument -Path (New-TestFile '<Configuration><Updates Enabled="TRUE" /></Configuration>')
    try { Test-CustomConfiguration -Document $document 6>$null; throw 'not rejected' }
    catch { Assert-Equal 2 (Get-InstallerExitCode $_) "exit code ($($_.Exception.Message))" }
}

Test-Case 'Custom validation warns (but allows) a non-2024 channel' {
    . $ImportInstaller
    $document = Import-ConfigurationDocument -Path (New-TestFile $legacy2021)
    $warnings = @()
    Test-CustomConfiguration -Document $document -WarningVariable warnings -WarningAction SilentlyContinue 6>$null
    Assert-True ($warnings.Count -ge 1) 'warning emitted'
}

Test-Case 'Options: RemoveMSI, unattended display, EULA, architecture, languages, source path, logging' {
    . $ImportInstaller
    $document = Import-ConfigurationDocument -Path (Join-Path $RepoRoot 'config_standard.xml')
    Set-ConfigurationOptions -Document $document -RemovePreviousVersions $true -Unattended $true -AcceptEula $true `
        -Architecture '32' -Language @('de-de', 'MatchOS') -SourcePath '\\server\share\Office' -LogDirectory 'C:\Logs'
    $summary = Get-ConfigurationSummary -Document $document
    Assert-True $summary.RemoveMSI 'RemoveMSI added'
    Assert-Equal 'None' $summary.DisplayLevel 'Display Level'
    Assert-Equal 'TRUE' $summary.AcceptEULA 'AcceptEULA'
    Assert-Equal '32' $summary.Architecture 'architecture'
    Assert-Equal 'de-de,MatchOS' ($summary.Languages -join ',') 'languages'
    Assert-Equal '\\server\share\Office' $summary.SourcePath 'SourcePath'
    Assert-Equal 'C:\Logs' $document.SelectSingleNode('/Configuration/Logging').GetAttribute('Path') 'logging path'
    Assert-Equal 'ExcludeApp' $document.SelectSingleNode('//Product').LastChild.Name 'ExcludeApp kept after languages'
}

Test-Case 'Options: unattended without -AcceptEula sets AcceptEULA=FALSE; no RemoveMSI unless requested' {
    . $ImportInstaller
    $document = Import-ConfigurationDocument -Path (Join-Path $RepoRoot 'config_ProPlus.xml')
    Set-ConfigurationOptions -Document $document -RemovePreviousVersions $false -Unattended $true -AcceptEula $false
    $summary = Get-ConfigurationSummary -Document $document
    Assert-Equal 'None' $summary.DisplayLevel 'Display Level'
    Assert-Equal 'FALSE' $summary.AcceptEULA 'AcceptEULA'
    Assert-True (-not $summary.RemoveMSI) 'RemoveMSI not added'
}

Test-Case 'Options: existing RemoveMSI, Display and Logging in a custom file are not duplicated' {
    . $ImportInstaller
    $custom = '<Configuration><Add Channel="PerpetualVL2024"><Product ID="ProPlus2024Volume" /></Add><RemoveMSI /><Display Level="Full" AcceptEULA="TRUE" /><Logging Level="Standard" Path="D:\x" /></Configuration>'
    $document = Import-ConfigurationDocument -Path (New-TestFile $custom)
    Set-ConfigurationOptions -Document $document -RemovePreviousVersions $true -Unattended $true -AcceptEula $false -LogDirectory 'C:\Logs'
    Assert-Equal 1 $document.SelectNodes('/Configuration/RemoveMSI').Count 'RemoveMSI count'
    Assert-Equal 1 $document.SelectNodes('/Configuration/Display').Count 'Display count'
    Assert-Equal 1 $document.SelectNodes('/Configuration/Logging').Count 'Logging count'
    Assert-Equal 'None' $document.SelectSingleNode('/Configuration/Display').GetAttribute('Level') 'unattended forces Level=None'
    Assert-Equal 'TRUE' $document.SelectSingleNode('/Configuration/Display').GetAttribute('AcceptEULA') 'existing AcceptEULA kept'
    Assert-Equal 'D:\x' $document.SelectSingleNode('/Configuration/Logging').GetAttribute('Path') 'existing logging kept'
}

Test-Case 'Options: -Architecture on a Remove-only configuration is rejected' {
    . $ImportInstaller
    $document = Import-ConfigurationDocument -Path (New-TestFile '<Configuration><Remove All="TRUE" /></Configuration>')
    try { Set-ConfigurationOptions -Document $document -Architecture '64'; throw 'not rejected' }
    catch { Assert-Equal 2 (Get-InstallerExitCode $_) "exit code ($($_.Exception.Message))" }
}

Test-Case 'Saved configuration is UTF-8 without BOM, CRLF, and reloads' {
    . $ImportInstaller
    $document = Import-ConfigurationDocument -Path (Join-Path $RepoRoot 'config_standard.xml')
    $out = Join-Path (New-TestDirectory) 'configuration.xml'
    Save-ConfigurationDocument -Document $document -Path $out
    $bytes = [IO.File]::ReadAllBytes($out)
    Assert-True (-not ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB)) 'no UTF-8 BOM'
    $text = [IO.File]::ReadAllText($out)
    Assert-True ($text.Contains("`r`n")) 'CRLF line endings'
    Assert-True (-not ($text -match "[^`r]`n")) 'no bare LF'
    $null = Import-ConfigurationDocument -Path $out
}

Test-Case 'Product keys are masked in displayed configuration' {
    . $ImportInstaller
    $document = Import-ConfigurationDocument -Path (New-TestFile $legacy2021)
    $text = Get-MaskedConfigurationText -Document $document
    Assert-True (-not $text.Contains($FakeKey)) 'key not shown'
    Assert-True ($text.Contains('*****-*****-*****-*****-FAKE9')) 'masked key shown'
    Assert-Equal $FakeKey $document.SelectSingleNode('//Product').GetAttribute('PIDKEY') 'original document unchanged'
}

# ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'Office Deployment Tool handling'
# ---------------------------------------------------------------------------

Test-Case 'Latest ODT link: absolute, page-relative, and missing' {
    . $ImportInstaller
    $absolute = '<a href="https://download.microsoft.com/download/abc/officedeploymenttool_20527-20150.exe">x</a>'
    Assert-Equal 'https://download.microsoft.com/download/abc/officedeploymenttool_20527-20150.exe' (Get-LatestOdtUrlFromContent $absolute) 'absolute link'
    $relative = '<a href="/download/abc/officedeploymenttool_1-2.exe">x</a>'
    Assert-Equal 'https://www.microsoft.com/download/abc/officedeploymenttool_1-2.exe' (Get-LatestOdtUrlFromContent $relative) 'relative link'
    Assert-Equal $null (Get-LatestOdtUrlFromContent '<html>nothing here</html>') 'no link'
}

Test-Case 'ODT plan: pinned build matches the expected version and hash format' {
    . $ImportInstaller
    $plan = Get-OdtDownloadPlan -OdtSource Pinned
    Assert-Equal $PinnedOdt.Url $plan.Url 'pinned URL'
    Assert-True ($plan.Sha256 -match '^[0-9A-F]{64}$') 'pinned SHA-256'
    Assert-True ($plan.Url -like 'https://download.microsoft.com/*') 'pinned URL is on download.microsoft.com'
}

Test-Case 'ODT plan: Auto uses the latest link from the Download Center' {
    . $ImportInstaller
    function Invoke-WebRequest { return [pscustomobject]@{ Content = '<a href="https://download.microsoft.com/download/new/officedeploymenttool_99999-99999.exe">' } }
    $plan = Get-OdtDownloadPlan -OdtSource Auto 6>$null
    Assert-Equal 'https://download.microsoft.com/download/new/officedeploymenttool_99999-99999.exe' $plan.Url 'latest URL'
    Assert-Equal $null $plan.Sha256 'no pinned hash for a newer build'
}

Test-Case 'ODT plan: Auto falls back to the pinned build when the page is unavailable' {
    . $ImportInstaller
    function Invoke-WebRequest { throw 'network down' }
    $plan = Get-OdtDownloadPlan -OdtSource Auto -WarningAction SilentlyContinue 6>$null
    Assert-Equal $PinnedOdt.Url $plan.Url 'fallback URL'
}

Test-Case 'ODT plan: Auto ignores a link that is not on microsoft.com' {
    . $ImportInstaller
    function Invoke-WebRequest { return [pscustomobject]@{ Content = 'https://download.microsoft.com.evil.example/download/x/officedeploymenttool_1-1.exe' } }
    $plan = Get-OdtDownloadPlan -OdtSource Auto -WarningAction SilentlyContinue 6>$null
    Assert-Equal $PinnedOdt.Url $plan.Url 'fallback URL'
}

Test-Case 'ODT plan: Latest fails with exit code 5 when the page is unavailable' {
    . $ImportInstaller
    function Invoke-WebRequest { throw 'network down' }
    try { $null = Get-OdtDownloadPlan -OdtSource Latest 6>$null; throw 'not rejected' }
    catch { Assert-Equal 5 (Get-InstallerExitCode $_) "exit code ($($_.Exception.Message))" }
}

Test-Case 'Signature check accepts a Microsoft-signed binary and rejects an unsigned file' {
    . $ImportInstaller
    Assert-True (Test-MicrosoftSignature -Path (Join-Path $env:SystemRoot 'System32\cmd.exe')) 'cmd.exe accepted'
    $unsigned = New-TestFile 'not a real executable' 'fake.exe'
    Assert-True (-not (Test-MicrosoftSignature -Path $unsigned)) 'unsigned file rejected'
}

Test-Case 'ODT download: hash mismatch is rejected before extraction (exit code 5)' {
    . $ImportInstaller
    function Start-Process { throw 'extraction must not run' }
    function Invoke-FileDownload { param($Uri, $OutFile) Set-Content -LiteralPath $OutFile -Value 'tampered' }
    $plan = @{ Url = 'https://download.microsoft.com/x.exe'; Sha256 = ('0' * 64); Description = 'test' }
    try { $null = Get-OfficeDeploymentTool -Plan $plan -Directory (New-TestDirectory) 6>$null; throw 'not rejected' }
    catch { Assert-Equal 5 (Get-InstallerExitCode $_) "exit code ($($_.Exception.Message))" }
}

Test-Case 'ODT download: unsigned executable is rejected before extraction (exit code 5)' {
    . $ImportInstaller
    function Start-Process { throw 'extraction must not run' }
    function Invoke-FileDownload { param($Uri, $OutFile) Set-Content -LiteralPath $OutFile -Value 'unsigned' }
    $plan = @{ Url = 'https://download.microsoft.com/x.exe'; Sha256 = $null; Description = 'test' }
    try { $null = Get-OfficeDeploymentTool -Plan $plan -Directory (New-TestDirectory) 6>$null; throw 'not rejected' }
    catch {
        Assert-Equal 5 (Get-InstallerExitCode $_) "exit code ($($_.Exception.Message))"
        Assert-True ($_.Exception.Message -match 'signed') 'signature message'
    }
}

Test-Case 'ODT download: network failure maps to exit code 5' {
    . $ImportInstaller
    function Invoke-FileDownload { throw 'connection reset' }
    $plan = @{ Url = 'https://download.microsoft.com/x.exe'; Sha256 = $null; Description = 'test' }
    try { $null = Get-OfficeDeploymentTool -Plan $plan -Directory (New-TestDirectory) 6>$null; throw 'not rejected' }
    catch { Assert-Equal 5 (Get-InstallerExitCode $_) "exit code ($($_.Exception.Message))" }
}

Test-Case 'Legacy "newurl" arguments are still accepted' {
    . $ImportInstaller
    Assert-Equal 'https://x/odt.exe' (Resolve-LegacyArguments @('newurl', 'ignored', 'https://x/odt.exe')) 'three-argument form'
    Assert-Equal 'https://x/odt.exe' (Resolve-LegacyArguments @('newurl', 'https://x/odt.exe')) 'two-argument form'
    Assert-Equal $null (Resolve-LegacyArguments @()) 'no arguments'
    try { $null = Resolve-LegacyArguments @('-Bogus'); throw 'not rejected' }
    catch { Assert-Equal 1 (Get-InstallerExitCode $_) 'unknown argument exit code' }
}

# ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'Unattended mode (stubbed setup.exe; nothing is installed)'
# ---------------------------------------------------------------------------

Test-Case 'Unattended ProPlus: download then configure, silent XML, cleanup, exit 0' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ Unattended = $true; Edition = 'ProPlus'; AcceptEula = $true; RemovePreviousVersions = $true; WorkingDirectory = $work }
    Assert-Equal 0 $result.ExitCode 'exit code'
    Assert-Equal 'plan:Auto|odt|setup:download|setup:configure' ($result.Calls -join '|') 'call sequence'
    Assert-True (-not ($result.Calls -match '^prompt:')) 'no prompts'
    $xml = New-Object System.Xml.XmlDocument
    $xml.LoadXml($result.Configs[1])
    Assert-Equal 'ProPlus2024Volume' $xml.SelectSingleNode('//Product').GetAttribute('ID') 'product'
    Assert-Equal 'PerpetualVL2024' $xml.SelectSingleNode('//Add').GetAttribute('Channel') 'channel'
    Assert-Equal 'None' $xml.SelectSingleNode('//Display').GetAttribute('Level') 'Display Level'
    Assert-Equal 'TRUE' $xml.SelectSingleNode('//Display').GetAttribute('AcceptEULA') 'AcceptEULA'
    Assert-Equal 1 $xml.SelectNodes('//RemoveMSI').Count 'RemoveMSI'
    Assert-Equal 0 @(Get-OdtFolder $work).Count 'working folder cleaned up'
}

Test-Case 'Unattended Standard: no RemoveMSI by default' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ Unattended = $true; Edition = 'Standard'; WorkingDirectory = $work }
    Assert-Equal 0 $result.ExitCode 'exit code'
    $xml = New-Object System.Xml.XmlDocument
    $xml.LoadXml($result.Configs[0])
    Assert-Equal 'Standard2024Volume' $xml.SelectSingleNode('//Product').GetAttribute('ID') 'product'
    Assert-Equal 0 $xml.SelectNodes('//RemoveMSI').Count 'RemoveMSI'
    Assert-Equal 'FALSE' $xml.SelectSingleNode('//Display').GetAttribute('AcceptEULA') 'AcceptEULA'
}

Test-Case 'Unattended custom configuration (-ConfigPath implies Custom)' {
    $work = New-TestDirectory
    $config = New-TestFile '<Configuration><Add OfficeClientEdition="64" Channel="PerpetualVL2024"><Product ID="ProPlus2024Volume"><Language ID="en-us" /></Product><Product ID="VisioPro2024Volume"><Language ID="en-us" /></Product></Add></Configuration>'
    $result = Invoke-StubbedInstaller -Parameters @{ Unattended = $true; ConfigPath = $config; WorkingDirectory = $work }
    Assert-Equal 0 $result.ExitCode 'exit code'
    Assert-True ($result.Configs[1].Contains('VisioPro2024Volume')) 'custom product passed to setup'
}

Test-Case 'setup.exe /configure failure returns exit code 8 and keeps files' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ Unattended = $true; Edition = 'ProPlus'; WorkingDirectory = $work } -SetupExitCodes @{ download = 0; configure = 17002 }
    Assert-Equal 8 $result.ExitCode 'exit code'
    Assert-Equal 1 @(Get-OdtFolder $work).Count 'working folder retained'
}

Test-Case 'setup.exe /download failure returns exit code 7 and skips configure' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ Unattended = $true; Edition = 'Standard'; WorkingDirectory = $work } -SetupExitCodes @{ download = 30088; configure = 0 }
    Assert-Equal 7 $result.ExitCode 'exit code'
    Assert-True (-not ($result.Calls -contains 'setup:configure')) 'configure not run'
    Assert-Equal 1 @(Get-OdtFolder $work).Count 'working folder retained'
}

Test-Case '-DownloadOnly runs only /download, keeps files, needs no elevation' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ Unattended = $true; Edition = 'Standard'; DownloadOnly = $true; WorkingDirectory = $work } -IsAdmin $false
    Assert-Equal 0 $result.ExitCode 'exit code'
    Assert-Equal 'plan:Auto|odt|setup:download' ($result.Calls -join '|') 'call sequence'
    Assert-Equal 1 @(Get-OdtFolder $work).Count 'working folder retained'
}

Test-Case '-SkipDownload with -SourcePath runs only /configure from the source' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ Unattended = $true; Edition = 'ProPlus'; SkipDownload = $true; SourcePath = '\\server\share\Office2024'; WorkingDirectory = $work }
    Assert-Equal 0 $result.ExitCode 'exit code'
    Assert-Equal 'plan:Auto|odt|setup:configure' ($result.Calls -join '|') 'call sequence'
    Assert-True ($result.Configs[0].Contains('SourcePath="\\server\share\Office2024"')) 'SourcePath in configuration'
}

Test-Case 'Not elevated: exit code 4 before anything is downloaded' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ Unattended = $true; Edition = 'ProPlus'; WorkingDirectory = $work } -IsAdmin $false
    Assert-Equal 4 $result.ExitCode 'exit code'
    Assert-Equal 0 $result.Calls.Count 'no actions'
}

Test-Case 'Insufficient disk space: exit code 3 before anything is downloaded' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ Unattended = $true; Edition = 'ProPlus'; WorkingDirectory = $work; MinimumFreeSpaceGB = 1024 }
    Assert-Equal 3 $result.ExitCode 'exit code'
    Assert-Equal 0 $result.Calls.Count 'no actions'
}

Test-Case '-Cleanup No keeps the files after success' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ Unattended = $true; Edition = 'ProPlus'; Cleanup = 'No'; WorkingDirectory = $work }
    Assert-Equal 0 $result.ExitCode 'exit code'
    Assert-Equal 1 @(Get-OdtFolder $work).Count 'working folder retained'
}

Test-Case '-OdtSource and -OdtUrl are passed to the ODT step' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ Unattended = $true; Edition = 'ProPlus'; OdtSource = 'Pinned'; WorkingDirectory = $work }
    Assert-Equal 'plan:Pinned' $result.Calls[0] 'OdtSource forwarded'
}

# ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'Interactive mode (stubbed prompts and setup.exe)'
# ---------------------------------------------------------------------------

Test-Case 'Interactive ProPlus: menu, remove previous = Y, proceed = Y, cleanup = N' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ WorkingDirectory = $work } -Answers @('2', 'Y', 'Y', 'N')
    Assert-Equal 0 $result.ExitCode 'exit code'
    Assert-Equal 4 @($result.Calls | Where-Object { $_ -like 'prompt:*' }).Count 'four prompts'
    $xml = New-Object System.Xml.XmlDocument
    $xml.LoadXml($result.Configs[1])
    Assert-Equal 'ProPlus2024Volume' $xml.SelectSingleNode('//Product').GetAttribute('ID') 'ProPlus selected'
    Assert-Equal 1 $xml.SelectNodes('/Configuration/RemoveMSI').Count 'RemoveMSI added'
    Assert-Equal 0 $xml.SelectNodes('/Configuration/Display').Count 'interactive install keeps the Office UI'
    Assert-Equal 1 @(Get-OdtFolder $work).Count 'files retained'
}

Test-Case 'Interactive Standard: remove previous = N, proceed = Y, cleanup = Y' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ WorkingDirectory = $work } -Answers @('1', 'N', 'Y', 'Y')
    Assert-Equal 0 $result.ExitCode 'exit code'
    $xml = New-Object System.Xml.XmlDocument
    $xml.LoadXml($result.Configs[1])
    Assert-Equal 'Standard2024Volume' $xml.SelectSingleNode('//Product').GetAttribute('ID') 'Standard selected'
    Assert-Equal 0 $xml.SelectNodes('/Configuration/RemoveMSI').Count 'RemoveMSI not added'
    Assert-Equal 0 @(Get-OdtFolder $work).Count 'files deleted'
}

Test-Case 'Interactive cancel at confirmation: exit code 10, nothing downloaded, folder removed' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ WorkingDirectory = $work } -Answers @('1', 'N', 'N')
    Assert-Equal 10 $result.ExitCode 'exit code'
    Assert-True (-not ($result.Calls -contains 'odt')) 'ODT not downloaded'
    Assert-Equal 0 @(Get-OdtFolder $work).Count 'working folder removed'
}

Test-Case 'Interactive invalid menu selection: exit code 1' {
    $result = Invoke-StubbedInstaller -Parameters @{ WorkingDirectory = (New-TestDirectory) } -Answers @('7')
    Assert-Equal 1 $result.ExitCode 'exit code'
}

Test-Case 'Interactive custom without Configuration.xml: exit code 2' {
    $result = Invoke-StubbedInstaller -Parameters @{ WorkingDirectory = (New-TestDirectory); ConfigPath = (Join-Path $TestRoot 'missing\Configuration.xml') }
    Assert-Equal 2 $result.ExitCode 'exit code'
}

Test-Case '-RemovePreviousVersions:$false skips the removal question' {
    $work = New-TestDirectory
    $result = Invoke-StubbedInstaller -Parameters @{ WorkingDirectory = $work; Edition = 'Standard'; RemovePreviousVersions = $false; Cleanup = 'Yes' } -Answers @('Y')
    Assert-Equal 0 $result.ExitCode 'exit code'
    Assert-Equal 1 @($result.Calls | Where-Object { $_ -like 'prompt:*' }).Count 'only the confirmation prompt'
}

# ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'Command line (separate non-interactive powershell.exe processes, -DryRun)'
# ---------------------------------------------------------------------------

$common = @('-LogDirectory', $LogDir, '-MinimumFreeSpaceGB', '0')

Test-Case 'Unattended dry run, ProPlus: exit 0 and silent configuration written' {
    $work = New-TestDirectory
    $run = Invoke-ScriptProcess $InstallerPath (@('-Unattended', '-Edition', 'ProPlus', '-AcceptEula', '-DryRun', '-WorkingDirectory', $work) + $common)
    Assert-Equal 0 $run.ExitCode "exit code. Output: $($run.Output)"
    $folders = @(Get-OdtFolder $work)
    Assert-Equal 1 $folders.Count 'dry-run folder'
    $xml = Import-Xml (Join-Path $folders[0].FullName 'configuration.xml')
    Assert-Equal 'ProPlus2024Volume' $xml.SelectSingleNode('//Product').GetAttribute('ID') 'product'
    Assert-Equal 'None' $xml.SelectSingleNode('//Display').GetAttribute('Level') 'Display Level'
    Assert-Equal @(Get-ChildItem -LiteralPath $folders[0].FullName).Count 1 'only configuration.xml (nothing downloaded)'
}

Test-Case 'Unattended dry run, Standard with languages as one comma-separated string' {
    $work = New-TestDirectory
    $run = Invoke-ScriptProcess $InstallerPath (@('-Unattended', '-Edition', 'Standard', '-Language', 'en-us,fr-fr', '-Architecture', '32', '-DryRun', '-WorkingDirectory', $work) + $common)
    Assert-Equal 0 $run.ExitCode "exit code. Output: $($run.Output)"
    $xml = Import-Xml (Join-Path (@(Get-OdtFolder $work)[0].FullName) 'configuration.xml')
    Assert-Equal 'en-us,fr-fr' ((@($xml.SelectNodes('//Language') | ForEach-Object { $_.GetAttribute('ID') })) -join ',') 'languages'
    Assert-Equal '32' $xml.SelectSingleNode('//Add').GetAttribute('OfficeClientEdition') 'architecture'
}

Test-Case 'Unattended dry run with a custom file containing a key: key never printed' {
    $config = New-TestFile ('<Configuration><Add OfficeClientEdition="64" Channel="PerpetualVL2024"><Product ID="ProPlus2024Volume" PIDKEY="{0}"><Language ID="en-us" /></Product></Add></Configuration>' -f $FakeKey)
    $run = Invoke-ScriptProcess $InstallerPath (@('-Unattended', '-ConfigPath', $config, '-DryRun', '-WorkingDirectory', (New-TestDirectory)) + $common)
    Assert-Equal 0 $run.ExitCode "exit code. Output: $($run.Output)"
    Assert-True (-not $run.Output.Contains($FakeKey)) 'key not printed'
    Assert-True ($run.Output.Contains('FAKE9')) 'masked suffix printed'
}

Test-Case 'Unattended without -Edition: exit code 1' {
    $run = Invoke-ScriptProcess $InstallerPath (@('-Unattended', '-DryRun', '-WorkingDirectory', (New-TestDirectory)) + $common)
    Assert-Equal 1 $run.ExitCode "exit code. Output: $($run.Output)"
}

Test-Case 'Interactive mode in a non-interactive session fails cleanly: exit code 1' {
    $run = Invoke-ScriptProcess $InstallerPath (@('-DryRun', '-WorkingDirectory', (New-TestDirectory)) + $common)
    Assert-Equal 1 $run.ExitCode "exit code. Output: $($run.Output)"
    Assert-True ($run.Output -match 'Unattended') 'message suggests -Unattended'
}

Test-Case 'Invalid language: exit code 1' {
    $run = Invoke-ScriptProcess $InstallerPath (@('-Unattended', '-Edition', 'Standard', '-Language', 'english', '-DryRun', '-WorkingDirectory', (New-TestDirectory)) + $common)
    Assert-Equal 1 $run.ExitCode "exit code. Output: $($run.Output)"
}

Test-Case '-DownloadOnly with -SkipDownload: exit code 1' {
    $run = Invoke-ScriptProcess $InstallerPath (@('-Unattended', '-Edition', 'Standard', '-DownloadOnly', '-SkipDownload', '-DryRun', '-WorkingDirectory', (New-TestDirectory)) + $common)
    Assert-Equal 1 $run.ExitCode "exit code. Output: $($run.Output)"
}

Test-Case 'Missing custom configuration: exit code 2' {
    $run = Invoke-ScriptProcess $InstallerPath (@('-Unattended', '-ConfigPath', (Join-Path $TestRoot 'nope.xml'), '-DryRun', '-WorkingDirectory', (New-TestDirectory)) + $common)
    Assert-Equal 2 $run.ExitCode "exit code. Output: $($run.Output)"
}

Test-Case 'Invalid custom configuration: exit code 2 and working folder kept for troubleshooting' {
    $work = New-TestDirectory
    $run = Invoke-ScriptProcess $InstallerPath (@('-Unattended', '-ConfigPath', (New-TestFile '<Configuration><Add>'), '-DryRun', '-WorkingDirectory', $work) + $common)
    Assert-Equal 2 $run.ExitCode "exit code. Output: $($run.Output)"
    Assert-Equal 1 @(Get-OdtFolder $work).Count 'working folder retained'
}

Test-Case 'Insufficient disk space: exit code 3' {
    $run = Invoke-ScriptProcess $InstallerPath @('-Unattended', '-Edition', 'ProPlus', '-DryRun', '-WorkingDirectory', (New-TestDirectory), '-LogDirectory', $LogDir, '-MinimumFreeSpaceGB', '1024')
    Assert-Equal 3 $run.ExitCode "exit code. Output: $($run.Output)"
}

Test-Case 'Legacy "newurl" form and https enforcement' {
    $ok = Invoke-ScriptProcess $InstallerPath (@('newurl', 'x', 'https://download.microsoft.com/custom/odt.exe', '-Unattended', '-Edition', 'Standard', '-DryRun', '-WorkingDirectory', (New-TestDirectory)) + $common)
    Assert-Equal 0 $ok.ExitCode "exit code. Output: $($ok.Output)"
    Assert-True ($ok.Output.Contains('https://download.microsoft.com/custom/odt.exe')) 'custom URL used'
    $bad = Invoke-ScriptProcess $InstallerPath (@('-OdtUrl', 'http://example.com/odt.exe', '-Unattended', '-Edition', 'Standard', '-DryRun', '-WorkingDirectory', (New-TestDirectory)) + $common)
    Assert-Equal 1 $bad.ExitCode 'http URL rejected'
}

Test-Case 'Unknown argument: exit code 1' {
    $run = Invoke-ScriptProcess $InstallerPath (@('-Unattended', '-Edition', 'Standard', '-Bogus', '-DryRun', '-WorkingDirectory', (New-TestDirectory)) + $common)
    Assert-Equal 1 $run.ExitCode "exit code. Output: $($run.Output)"
}

Test-Case 'Transcript written to -LogDirectory' {
    Assert-True (@(Get-ChildItem -LiteralPath $LogDir -Filter 'Install_Office_*.log').Count -ge 1) 'transcript exists'
}

Test-Case 'install_office_nodefault.PS1 forwards arguments with -OdtSource Latest' {
    $run = Invoke-ScriptProcess $NoDefaultPath (@('-Unattended', '-Edition', 'Standard', '-DryRun', '-WorkingDirectory', (New-TestDirectory)) + $common)
    Assert-Equal 0 $run.ExitCode "exit code. Output: $($run.Output)"
    Assert-True ($run.Output.Contains('OdtSource=Latest')) 'Latest ODT source used'
    $fail = Invoke-ScriptProcess $NoDefaultPath (@('-Unattended', '-DryRun') + $common)
    Assert-Equal 1 $fail.ExitCode 'exit code propagated'
}

# ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'Activation helper (dry run only; licensing is never changed)'
# ---------------------------------------------------------------------------

Test-Case 'Default action is a status check' {
    $run = Invoke-ScriptProcess $ActivatorPath @('-DryRun')
    Assert-Equal 0 $run.ExitCode "exit code. Output: $($run.Output)"
    Assert-True ($run.Output.Contains('/dstatus')) 'dstatus'
    Assert-True (-not $run.Output.Contains('/act')) 'no activation'
}

Test-Case 'KMS host: /sethst, /setprt, /act' {
    $run = Invoke-ScriptProcess $ActivatorPath @('-KmsHost', 'kms.contoso.com', '-KmsPort', '1688', '-DryRun')
    Assert-Equal 0 $run.ExitCode "exit code. Output: $($run.Output)"
    Assert-True ($run.Output -match '/sethst:kms\.contoso\.com[\s\S]*/setprt:1688[\s\S]*/act') 'command order'
}

Test-Case 'MAK: key masked in output' {
    $run = Invoke-ScriptProcess $ActivatorPath @('-ProductKey', $FakeKey, '-Status', '-DryRun')
    Assert-Equal 0 $run.ExitCode "exit code. Output: $($run.Output)"
    Assert-True (-not $run.Output.Contains($FakeKey)) 'key not printed'
    Assert-True ($run.Output.Contains('/inpkey:*****-*****-*****-*****-FAKE9')) 'masked key'
    Assert-True ($run.Output -match '/act[\s\S]*/dstatus') 'activate then status'
}

Test-Case 'Invalid arguments: exit code 1' {
    Assert-Equal 1 (Invoke-ScriptProcess $ActivatorPath @('-ProductKey', 'not-a-key', '-DryRun')).ExitCode 'malformed key'
    Assert-Equal 1 (Invoke-ScriptProcess $ActivatorPath @('-KmsHost', 'kms', '-ClearKmsHost', '-DryRun')).ExitCode 'KmsHost with ClearKmsHost'
    Assert-Equal 1 (Invoke-ScriptProcess $ActivatorPath @('-KmsPort', '1688', '-DryRun')).ExitCode 'KmsPort without KmsHost'
    Assert-Equal 1 (Invoke-ScriptProcess $ActivatorPath @('-KmsHost', 'bad host;rm', '-DryRun')).ExitCode 'invalid host name'
}

Test-Case 'OSPP.VBS not found: exit code 2 (nothing run)' {
    $run = Invoke-ScriptProcess $ActivatorPath @('-OsppPath', (Join-Path $TestRoot 'missing\OSPP.VBS'), '-Activate')
    Assert-Equal 2 $run.ExitCode "exit code. Output: $($run.Output)"
}

# ---------------------------------------------------------------------------
if ($Online) {
    Write-Host ''
    Write-Host 'Online: current Office Deployment Tool from Microsoft (setup.exe is never run)'

    Test-Case 'Resolve, download, verify and extract the latest ODT' {
        . $ImportInstaller
        $plan = Get-OdtDownloadPlan -OdtSource Latest 6>$null
        Assert-True ($plan.Url -like 'https://download.microsoft.com/*officedeploymenttool_*.exe') "latest URL: $($plan.Url)"
        $setup = Get-OfficeDeploymentTool -Plan $plan -Directory (New-TestDirectory) 6>$null
        Assert-True (Test-MicrosoftSignature -Path $setup) 'setup.exe signed by Microsoft'
        Write-Host "        latest ODT: $($plan.Url) (setup.exe $((Get-Item -LiteralPath $setup).VersionInfo.FileVersion))" -ForegroundColor DarkGray
    }

    Test-Case 'Pinned ODT still downloads and matches its SHA-256' {
        . $ImportInstaller
        $setup = Get-OfficeDeploymentTool -Plan (Get-OdtDownloadPlan -OdtSource Pinned) -Directory (New-TestDirectory) 6>$null
        Assert-True (Test-Path -LiteralPath $setup) 'setup.exe extracted'
    }
}

} finally {
    Remove-Item -LiteralPath $TestRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------------
$failed = @($Results | Where-Object { -not $_.Passed })
Write-Host ''
Write-Host ("{0} tests, {1} passed, {2} failed" -f $Results.Count, ($Results.Count - $failed.Count), $failed.Count)
if ($failed.Count -gt 0) {
    $failed | ForEach-Object { Write-Host " - $($_.Name): $($_.Message)" -ForegroundColor Red }
    exit 1
}
exit 0
