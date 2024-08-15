# Define variables for download URLs and file names
$odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17830-20162.exe"
$outputDir = "$PSScriptRoot\OfficeDeploymentTool"  # Output directory for extraction
$odtInstaller = "$outputDir\ODTInstaller.exe"

# Create OfficeDeploymentTool directory if it doesn't exist
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
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

# Download configuration.xml using Invoke-WebRequest
$configUrl = "https://raw.githubusercontent.com/jdepew88/download_office_ltsc/main/Configuration.xml"
$configFile = "$outputDir\configuration.xml"
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

# Inform the user about the download process
Write-Output "Office 2021 LTSC installation files have been downloaded to: $outputDir"

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
