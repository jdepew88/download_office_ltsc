# Office LTSC 2024 Deployment Script

This PowerShell script automates deploying Microsoft Office LTSC 2024 (Long-Term Servicing Channel) with Microsoft's Office Deployment Tool (ODT), following [Microsoft's official deployment documentation](https://learn.microsoft.com/en-us/office/ltsc/2024/deploy). It supports Office LTSC Standard 2024, Office LTSC Professional Plus 2024, or a custom configuration, either interactively or fully unattended.

## Features

- Installs volume licensed **Office LTSC Standard 2024** (`Standard2024Volume`) or **Office LTSC Professional Plus 2024** (`ProPlus2024Volume`) from the `PerpetualVL2024` channel.
- **Custom configuration** support: use any `Configuration.xml` (for example, one created with the [Office Customization Tool](https://config.office.com/deploymentsettings)).
- Downloads the **current Office Deployment Tool** from the Microsoft Download Center, falling back to a pinned, hash-verified build. Every ODT download is checked for a valid **Microsoft Authenticode signature** before it is run.
- **Interactive mode** (menu and prompts) and **unattended mode** (no prompts, documented exit codes) for automated deployment, such as Windows Deployment Builder.
- Optional removal of previous MSI-based Office versions (`<RemoveMSI />`).
- Offline staging (`-DownloadOnly`) and install-from-share (`-SkipDownload -SourcePath`).
- Validates configuration XML before anything is downloaded, checks free disk space and elevation, and logs every run.
- Optional cleanup of the downloaded files.
- **Activation is separate and optional.** No product keys are included anywhere in this repository (see [Activation](#activation)).

## Prerequisites

- Windows 10/11 or Windows Server with Windows PowerShell 5.1 (or PowerShell 7)
- Administrator rights (to install Office)
- Internet access to download the Office Deployment Tool and Office files (or a prepared `-SourcePath`)
- At least 10 GB of free space (configurable with `-MinimumFreeSpaceGB`)

## Files

| File | Purpose |
| --- | --- |
| `Install_Office.ps1` | Main installer (interactive or unattended). |
| `install_office_nodefault.PS1` | Same installer, but always uses the latest ODT from the Download Center with no pinned fallback (`-OdtSource Latest`). Requires `Install_Office.ps1` in the same folder. |
| `config_standard.xml` | ODT configuration for Office LTSC Standard 2024 (64-bit, English + OS language). |
| `config_ProPlus.xml` | ODT configuration for Office LTSC Professional Plus 2024 (64-bit, English + OS language). |
| `Activate_Office.ps1` | Optional activation helper (KMS host, MAK, status) using Microsoft's `OSPP.VBS`. |
| `tests/Test-OfficeInstaller.ps1` | Test suite. It never installs, uninstalls or activates Office. |

## How to Use

### 1. Clone or download the repository

```sh
git clone https://github.com/jdepew88/download_office_ltsc.git
```

Keep the `.xml` files next to `Install_Office.ps1`. If they are missing, the script downloads them from this repository's `main` branch and verifies that they are Office LTSC 2024 configurations.

### 2. Run the script interactively

Open PowerShell **as Administrator**, go to the repository folder and run:

```powershell
cd path\to\download_office_ltsc
.\Install_Office.ps1
```

You will be asked to choose an edition:

```text
Select the Office LTSC 2024 edition to install:
1. Office LTSC Standard 2024
2. Office LTSC Professional Plus 2024
3. Custom (place your Configuration.xml file in the same directory as this script)
Enter the number of your choice (1, 2, or 3):
```

The script then asks whether to remove previous Office versions, asks for confirmation, downloads and installs Office, and finally offers to delete the downloaded files.

- **Custom configuration:** for option 3, place `Configuration.xml` next to the script (or pass `-ConfigPath`). If it is missing or not valid XML, the script stops before downloading anything.

### 3. Run unattended (Windows Deployment Builder, task sequences, scripts)

```powershell
# Professional Plus, remove MSI-based Office, accept the license terms
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install_Office.ps1 -Unattended -Edition ProPlus -AcceptEula -RemovePreviousVersions

# Standard, 32-bit, English and French
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install_Office.ps1 -Unattended -Edition Standard -Architecture 32 -Language en-us,fr-fr -AcceptEula

# Custom configuration and log folder
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install_Office.ps1 -Unattended -ConfigPath D:\Deploy\Office2024.xml -LogDirectory D:\Logs

# Preview what would happen, without downloading or installing anything
.\Install_Office.ps1 -Unattended -Edition ProPlus -DryRun
```

In unattended mode the script never prompts, sets `<Display Level="None" />`, deletes downloaded files after a successful install (override with `-Cleanup No`), and returns an exit code.

#### Parameters

| Parameter | Description |
| --- | --- |
| `-Edition Standard\|ProPlus\|Custom` | Edition to install. Required with `-Unattended` unless `-ConfigPath` is given. |
| `-ConfigPath <file>` | Custom ODT configuration XML (implies `-Edition Custom`). |
| `-Unattended` | No prompts; silent Office setup. |
| `-AcceptEula` | Sets `AcceptEULA="TRUE"`. Use only if your organization has accepted the Microsoft license terms. Without it, users see the license terms when they first open Office. |
| `-RemovePreviousVersions` | Adds `<RemoveMSI />` to remove MSI-based Office (2016 and earlier). Interactively you are asked unless this is specified (`-RemovePreviousVersions:$false` skips the question). |
| `-Architecture 64\|32` | Overrides `OfficeClientEdition`. |
| `-Language <ids>` | Replaces the language list, for example `en-us,fr-fr` or `MatchOS`. |
| `-SourcePath <path>` | Where `/download` stores files and `/configure` installs from (folder or `\\server\share`). |
| `-DownloadOnly` | Only stage the installation files (`setup.exe /download`); does not need elevation. |
| `-SkipDownload` | Install directly from `-SourcePath` or the Office CDN. |
| `-OdtSource Auto\|Latest\|Pinned` | `Auto` (default): latest ODT from Microsoft, pinned build as fallback. `Latest`: fail if the latest cannot be found. `Pinned`: hash-verified pinned build only. |
| `-OdtUrl <https url>` | Explicit ODT download URL (its Microsoft signature is still verified). The original `newurl <x> <url>` argument form also still works. |
| `-WorkingDirectory <path>` | Parent folder for `OfficeDeploymentTool_<timestamp>` (default: script folder). |
| `-LogDirectory <path>` | Transcript and ODT logs (default: `%ProgramData%\OfficeLTSC2024Installer\Logs`). |
| `-Cleanup Ask\|Yes\|No` | Delete downloaded files after a successful install. Files are always kept after a failure. |
| `-MinimumFreeSpaceGB <n>` | Required free space on the system and download drives (default 10). |
| `-DryRun` | Validate everything and write the final `configuration.xml`, but make no changes. |

Run `Get-Help .\Install_Office.ps1 -Detailed` for the full help.

#### Exit codes

| Code | Meaning |
| --- | --- |
| 0 | Success (or dry run completed) |
| 1 | Invalid arguments or input |
| 2 | Configuration file missing or invalid |
| 3 | Insufficient disk space |
| 4 | Administrator rights required |
| 5 | Office Deployment Tool download or signature/hash verification failed |
| 6 | Office Deployment Tool extraction failed |
| 7 | Office download (`setup.exe /download`) failed |
| 8 | Office installation (`setup.exe /configure`) failed |
| 10 | Cancelled by the user |
| 99 | Unexpected error |

#### Offline / network-share deployment

```powershell
# 1. On a machine with internet access: stage the files
.\Install_Office.ps1 -Unattended -Edition ProPlus -DownloadOnly -SourcePath \\server\share\Office2024

# 2. On each target machine: install from the share
.\Install_Office.ps1 -Unattended -Edition ProPlus -AcceptEula -SkipDownload -SourcePath \\server\share\Office2024
```

The ODT itself is still downloaded from Microsoft in step 2. To fully air-gap a deployment, run `setup.exe /configure` from the staged folder directly.

## Using the Office Customization Tool

- To build your own configuration file, use the [Office Customization Tool](https://config.office.com/deploymentsettings) and select an Office LTSC 2024 volume licensed product.
- This web-based tool lets you choose products, apps, update channel, languages and more.
- Download the resulting `Configuration.xml` and place it next to the script (or pass `-ConfigPath`).

## Removing previous versions

- `-RemovePreviousVersions` (or answering **Y** interactively) adds `<RemoveMSI />`, which removes **Windows Installer (MSI)** versions of Office (Office 2016 and earlier). See [Remove existing versions of Office](https://learn.microsoft.com/en-us/office/ltsc/2024/deploy#remove-existing-versions-of-office-before-installing-office-ltsc-2024).
- Click-to-Run versions such as Office 2019 or Office LTSC 2021 are **not** removed by `<RemoveMSI />`. Microsoft recommends removing them with the ODT `Remove` element, for example with a custom configuration file passed to `-ConfigPath ... -SkipDownload`:

  ```xml
  <Configuration>
    <Remove>
      <Product ID="ProPlus2021Volume" />
    </Remove>
    <Display Level="None" AcceptEULA="TRUE" />
  </Configuration>
  ```

## Activation

Activation is deliberately **separate and optional**, and only Microsoft-authorized methods are used. No product keys are stored in this repository.

- **KMS:** volume licensed Office LTSC 2024 is installed with Microsoft's [Generic Volume License Key (GVLK)](https://learn.microsoft.com/en-us/office/volume-license-activation/gvlks) by default and activates automatically against a KMS host published in DNS.
- **Active Directory-based activation:** automatic; no action required.
- **MAK:** provide your organization's Multiple Activation Key at deployment time.

`Activate_Office.ps1` wraps Microsoft's `OSPP.VBS`:

```powershell
.\Activate_Office.ps1                                  # show license status (/dstatus)
.\Activate_Office.ps1 -KmsHost kms.contoso.com         # set a KMS host and activate
.\Activate_Office.ps1 -ProductKey $env:OFFICE_MAK      # install a MAK and activate
.\Activate_Office.ps1 -ClearKmsHost -Activate          # return to DNS auto-discovery
.\Activate_Office.ps1 -KmsHost kms.contoso.com -DryRun # show the commands only
```

Product keys are masked in all output. Pass MAKs from a secure source (for example a deployment secret or environment variable) rather than saving them in scripts. The [Volume Activation Management Tool (VAMT)](https://learn.microsoft.com/en-us/windows/deployment/volume-activation/volume-activation-management-tool) can also manage activation centrally.

## Testing

The test suite validates PowerShell syntax, the XML configurations, error handling and exit codes, and both interactive and unattended flows. `setup.exe`, prompts and network calls are stubbed, so it never installs, uninstalls or activates Office:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-OfficeInstaller.ps1
# Also download, verify and extract (but not run) the current ODT from Microsoft:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-OfficeInstaller.ps1 -Online
```

PSScriptAnalyzer is run automatically when the module is installed.

## Troubleshooting

- **Script execution policy:** if scripts are blocked, allow them for the current session:

  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope Process
  ```

- **Logs:** each run writes a transcript and the ODT logs to `%ProgramData%\OfficeLTSC2024Installer\Logs` (or `-LogDirectory`).
- **Internet connectivity:** the script downloads the ODT and Office files from Microsoft; ensure a stable connection or use `-SourcePath`.
- **Custom configuration file:** if using a custom configuration, check that `Configuration.xml` is next to the script (or use `-ConfigPath`). Use `-DryRun` to validate it without installing.
- **Running Office apps:** close Office applications before installing; the configurations use `FORCEAPPSHUTDOWN=FALSE`.

## License

This project is licensed under the MIT License. See the [LICENSE](https://github.com/jdepew88/download_office_ltsc/blob/main/LICENSE) file for more details.
