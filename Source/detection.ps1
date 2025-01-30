# ==============================================================
# Intune Email Signature Detection Script
# Author: Lukas "Luc3as" Porubcan
# GitHub: https://github.com/luc3as
# Version: 2.8
# Created: 2025
# License: MIT
#
# Description:
# This script checks the installed email signature version 
# against the latest version stored in a shared Excel file.
# If the installed version is outdated or missing, it returns Exit Code 1
# to trigger an update.
#
# Features:
# - Downloads the latest version data from Google Drive.
# - Saves the data to a temporary Excel file.
# - Optimized for Intune compliance detection. 
# ==============================================================

# Define variables
$signatureFolder = "$($env:APPDATA)\Microsoft\Signatures"
$signaturePrefix = "BelusaFoods"
$googleDriveFileId = "REDACTED"
$googleDriveFileUrl = "https://drive.google.com/uc?export=download&id=$googleDriveFileId"
$tempExcelPath = "$env:TEMP\Signatures_data.xlsx"
$Debug = $true  # Set to $true for debugging (Prevents exiting)

# Function to handle safe exit
function SafeExit {
    param ([int]$code = 0)
    if ($Debug) {
        Write-Host "Debug Mode: Suppressed Exit ($code)"
        Exit $code
    }
    Exit $code
}

# Step 1: **Locate the installed version file**
$versionFile = Get-ChildItem -Path $signatureFolder -Filter "$signaturePrefix * version.txt" -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $versionFile) {
    Write-Host "Signature version file not found. Installation required."
    SafeExit 1
}

# Step 2: **Read the installed version from the file**
try {
    $installedVersion = Get-Content -Path $versionFile.FullName -Raw | ForEach-Object { $_.Trim() }
}
catch {
    Write-Host "Error: Could not read installed version file. $_"
    SafeExit 1
}

# Step 3: **Ensure ImportExcel module is installed**
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Error: ImportExcel module is missing. Ensure it's pre-installed."
    SafeExit 1
}

# Step 4: **Download Excel file to a temporary location**
try {
    Invoke-WebRequest -Uri $googleDriveFileUrl -OutFile $tempExcelPath -UseBasicParsing
}
catch {
    Write-Host "Error: Failed to download Excel file. $_"
    SafeExit 1
}

# Step 5: **Extract latest version from A1 (first cell in Excel)**
$latestVersion = $null
try {
    $versionData = Import-Excel -Path $tempExcelPath -WorksheetName "signatures" -StartRow 1 -EndRow 1 -NoHeader

    if ($versionData -and $versionData.PSObject.Properties.Count -gt 0) {
        $rawVersion = $versionData.PSObject.Properties.Value[0]
        if ($rawVersion -match "^Version:") {
            $latestVersion = $rawVersion -replace "Version:", "" -replace "^\s+|\s+$", ""
        }
        else {
            Write-Host "Error: Version data is missing or invalid."
            SafeExit 1
        }
    }
    else {
        Write-Host "Error: No version data found in Excel."
        SafeExit 1
    }
}
catch {
    Write-Host "Error: Could not retrieve version from Excel file. $_"
    SafeExit 1
}

# Step 6: **Compare installed vs latest version**
if ($installedVersion -eq $latestVersion) {
    Write-Host "Signature is up to date. Installed: $installedVersion, Expected: $latestVersion"
    SafeExit 0
}
else {
    Write-Host "Signature is outdated. Installed: $installedVersion, Expected: $latestVersion"
    SafeExit 1
}

# Cleanup: Remove the temporary Excel file
Remove-Item -Path $tempExcelPath -Force -ErrorAction SilentlyContinue
