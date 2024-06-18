# Define variables for download URLs and file names
$setupUrl = "https://github.com/jdepew88/download_office_ltsc/raw/main/setup.exe"
$configUrl = "https://github.com/jdepew88/download_office_ltsc/blob/main/Configuration.xml?raw=true"  # Updated URL for raw file download
$outputDir = "$PSScriptRoot\OfficeDeploymentTool"  # Output directory for extraction

# Create OfficeDeploymentTool directory if it doesn't exist
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Download setup.exe using wget
Invoke-Expression "wget -OutFile '$outputDir\setup.exe' '$setupUrl'"

# Download configuration.xml using wget
Invoke-Expression "wget -OutFile '$outputDir\configuration.xml' '$configUrl'"

# Inform the user about the download process
Write-Output ""
Write-Output "Office 2021 LTSC installation files have been downloaded to:"
Write-Output "$outputDir"

# Change directory to where Office Deployment Tool is extracted
Set-Location $outputDir

# Run setup to download Office
Write-Output "Running setup /download configuration.xml..."
Start-Process -FilePath ".\setup.exe" -ArgumentList "/download configuration.xml" -Wait

# Inform the user about the installation process
Write-Output ""
Write-Output "Office 2021 LTSC installation files have been downloaded."
Write-Output "A window will now open to show the installation progress."

# Run setup to install Office
Write-Output "Running setup /configure configuration.xml..."
Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure configuration.xml" -Wait

Write-Output "Office installation completed."
