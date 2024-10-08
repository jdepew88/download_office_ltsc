# URL of the ODT download page
$odtDownloadPageUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=49117"

# Fetch the content of the download page
try {
    $pageContent = Invoke-WebRequest -Uri $odtDownloadPageUrl -ErrorAction Stop
    Write-Output "Fetched the ODT download page successfully."
} catch {
    Write-Error "Failed to fetch the ODT download page: $_"
    exit 1
}

# Extract the download link from the page content using regex
$odtDownloadLink = $pageContent.Links | Where-Object { $_.href -match "/download/.*.exe" } | Select-Object -ExpandProperty href | Select-Object -First 1

# Check if a valid link was found
if ($odtDownloadLink) {
    # Ensure the URL is properly constructed
    if ($odtDownloadLink.StartsWith("/")) {
        $odtUrl = "https://www.microsoft.com$odtDownloadLink"
    } else {
        $odtUrl = "https://www.microsoft.com/$odtDownloadLink"
    }
    Write-Output "Latest ODT URL found: $odtUrl"
} else {
    Write-Error "Failed to find the latest ODT URL."
    exit 1
}

# Define configuration file paths with the updated URLs
$configStandardUrl = "https://raw.githubusercontent.com/jdepew88/download_office_ltsc/main/config_standard.xml"
$configProPlusUrl = "https://raw.githubusercontent.com/jdepew88/download_office_ltsc/main/config_ProPlus.xml"
$customConfigFile = "$PSScriptRoot\Configuration.xml"

# Prompt the user to select a configuration file
Write-Output "Select the Office 2021 version to install:"
Write-Output "1. Office 2021 Standard"
Write-Output "2. Office 2021 ProPlus"
Write-Output "3. Custom (Place your Configuration.xml file in the same directory as this script)"
$selection = Read-Host "Enter the number of your choice (1, 2, or 3)"

# Set the configuration file based on the user's selection
switch ($selection) {
    1 {
        $configUrl = $configStandardUrl
        Write-Output "You selected Office 2021 Standard."
    }
    2 {
        $configUrl = $configProPlusUrl
        Write-Output "You selected Office 2021 ProPlus."
    }
    3 {
        if (Test-Path $customConfigFile) {
            Write-Output "You selected Custom configuration."
            $configUrl = $customConfigFile
        } else {
            Write-Output "Custom configuration file not found. Please place your Configuration.xml file in the same directory as this Script."
            Read-Host "Press Enter to exit."
            exit 1
        }
    }
    default {
        Write-Output "Invalid selection. Exiting script."
        exit 1
    }
}

# Check for available disk space before proceeding
$minRequiredSpaceGB = 10  # Minimum required space in GB
$freeSpaceGB = [math]::Round((Get-PSDrive -PSProvider FileSystem -Name C).Free / 1GB, 2)

if ($freeSpaceGB -lt $minRequiredSpaceGB) {
    Write-Error "Insufficient disk space. At least $minRequiredSpaceGB GB of free space is required."
    exit 1
} else {
    Write-Output "Sufficient disk space detected: $freeSpaceGB GB available."
}

# Create a unique directory name using a timestamp
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$outputDir = "$PSScriptRoot\OfficeDeploymentTool_$timestamp"
$odtInstaller = "$outputDir\ODTInstaller.exe"

# Create the unique OfficeDeploymentTool directory
try {
    New-Item -ItemType Directory -Path $outputDir -ErrorAction Stop | Out-Null
    Write-Output "Created output directory: $outputDir"
} catch {
    Write-Error "Failed to create output directory: $_"
    exit 1
}

# Download Office Deployment Tool from Microsoft
try {
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtInstaller -ErrorAction Stop
    if (Test-Path $odtInstaller) {
        Write-Output "Downloaded Office Deployment Tool successfully."
    } else {
        Write-Error "Office Deployment Tool download failed. File not found."
        exit 1
    }
} catch {
    Write-Error "Failed to download Office Deployment Tool: $_"
    exit 1
}

# Silently extract the Office Deployment Tool
try {
    Start-Process -FilePath $odtInstaller -ArgumentList "/quiet /extract:$outputDir" -Wait
    if (Test-Path "$outputDir\setup.exe") {
        Write-Output "Extracted Office Deployment Tool successfully."
    } else {
        Write-Error "Extraction failed. setup.exe not found."
        exit 1
    }
} catch {
    Write-Error "Failed to extract Office Deployment Tool: $_"
    exit 1
}

# Download or use the selected configuration file
$configFile = "$outputDir\configuration.xml"
if ($selection -ne 3) {
    try {
        Invoke-WebRequest -Uri $configUrl -OutFile $configFile -ErrorAction Stop
        if (Test-Path $configFile) {
            Write-Output "Downloaded configuration.xml successfully."
        } else {
            Write-Error "Configuration.xml download failed. File not found."
            exit 1
        }
    } catch {
        Write-Error "Failed to download configuration.xml: $_"
        exit 1
    }
} else {
    Copy-Item $customConfigFile -Destination $configFile -Force
}

# Prompt the user if they want to remove previous Office versions
Write-Output "NOTE: Removing previous versions of Office is optional."
Write-Output "You may want to retain the existing installation if:"
Write-Output "1. You are performing an update and want to ensure settings and customizations are preserved."
Write-Output "2. You are installing a different edition or version side-by-side with an existing Office installation."
Write-Output "3. You are troubleshooting or comparing features between different Office versions."
$removePrevious = Read-Host "Do you want to remove previous versions of Office on this computer? (Y/N)"

if ($removePrevious -eq "Y") {
    # Insert <RemoveMSI /> into the configuration.xml file
    try {
        $configContent = Get-Content $configFile
        $configContent -replace '</Configuration>', '<RemoveMSI /></Configuration>' | Set-Content $configFile
        Write-Output "The option to remove previous versions of Office has been added to the configuration.xml."
    } catch {
        Write-Error "Failed to modify configuration.xml to remove previous Office versions: $_"
        exit 1
    }
} else {
    Write-Output "Previous Office versions will not be removed."
}

# Confirm with the user before starting the installation
$confirmation = Read-Host "Ready to install Office 2021 LTSC. Do you want to proceed? (Y/N)"
if ($confirmation -ne "Y") {
    Write-Output "Installation canceled by the user."
    exit 1
}

# Change directory to where Office Deployment Tool is extracted
Set-Location $outputDir

# Run setup to download Office
Write-Output "Running setup /download configuration.xml..."
try {
    Start-Process -FilePath ".\setup.exe" -ArgumentList "/download configuration.xml" -Wait
    Write-Output "Office 2021 LTSC files have been downloaded successfully."
} catch {
    Write-Error "Failed to download Office 2021 LTSC files: $_"
    exit 1
}

# Inform the user about the installation process
Write-Output "Running setup /configure configuration.xml..."
try {
    Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure configuration.xml" -Wait
    Write-Output "Office installation completed successfully."
} catch {
    Write-Error "Failed to install Office 2021 LTSC: $_"
    exit 1
}

# Cleanup option with warning
Write-Output "IMPORTANT: Please ensure that the Office setup window is closed before proceeding."
$cleanup = Read-Host "Do you want to delete the downloaded files to free up space? (Y/N)"
if ($cleanup -eq "Y") {
    Remove-Item -Recurse -Force $outputDir
    Write-Output "Downloaded files have been deleted."
} else {
    Write-Output "Downloaded files have been retained."
}
