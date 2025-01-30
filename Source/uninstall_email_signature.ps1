# ==============================================================
# Intune Email Signature Uninstallation Script
# Author: Lukas "Luc3as" Porubcan
# GitHub: https://github.com/luc3as
# Version: 2.8
# Created: 2025
# License: MIT
#
# Description:
# This script removes the installed email signature files and 
# corresponding Outlook registry entries for new and reply/forward emails.
#
# Features:
# - Detects and removes installed signatures dynamically.
# - Removes the corresponding version file.
# - Searches and deletes registry keys linked to the user's signature.
# - Works with Microsoft Intune for automated deployment.
#
# Exit Codes:
# - Exit 0: Uninstallation successful.
# - Exit 1: Failed to remove some or all components.
# ==============================================================  


# Ensure the script is running as administrator
if (-Not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Exit
}

# Start logging
$logPath = "$env:ProgramData\Microsoft\Intune\Logs"
if (-not (Test-Path $logPath)) {
    try {
        New-Item -ItemType Directory -Force -Path $logPath | Out-Null
    }
    catch {
        Write-Host "Failed to create log directory: $_"
        Exit 1
    }
}

Start-Transcript -Path "$logPath\Email_signature_Uninstallation.log" -Force

# Define logging function with timestamps
function Log() {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)] [String] $message
    )
    $ts = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    Write-Host "$ts $message"
}

# Define the signature folder path
$signatureFolder = "$($env:APPDATA)\Microsoft\Signatures"

# Define the signature prefix used during installation
$signaturePrefix = "BelusaFoods"

# Get all files and folders in the Signatures directory that match the installed pattern
$installedSignatures = Get-ChildItem -Path $signatureFolder -Filter "$signaturePrefix*" -Recurse

foreach ($signatureItem in $installedSignatures) {
    if ($signatureItem.PSIsContainer) {
        # Check if it's a _files directory (Prefix <date> (<email>)_files)
        if ($signatureItem.Name -match "^$signaturePrefix\s\d{1,2}-\d{1,2}-\d{4}\s\(.+\)_files$") {
            Log "Removing directory: $($signatureItem.FullName)"
            Remove-Item -Path $signatureItem.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    else {
        # Check for .htm, .rtf, .txt signature files (Prefix <date> (<email>).extension)
        if ($signatureItem.Name -match "^$signaturePrefix\s\d{1,2}-\d{1,2}-\d{4}\s\(.+\)\.(htm|rtf|txt)$") {
            Write-Host "Removing file: $($signatureItem.FullName)"
            Remove-Item -Path $signatureItem.FullName -Force -ErrorAction SilentlyContinue
        }
        # Check for version files (Prefix <date> version.txt)
        elseif ($signatureItem.Name -match "^$signaturePrefix\s\d{1,2}-\d{1,2}-\d{4}\sversion\.txt$") {
            Write-Host "Removing version file: $($signatureItem.FullName)"
            Remove-Item -Path $signatureItem.FullName -Force -ErrorAction SilentlyContinue
        }
    }
}

Log "Signature files successfully removed."

# Remove the New and Reply Signature from Outlook registry
try {
    # Define the base Outlook registry path
    $outlookRegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook"
    if (Test-Path $outlookRegPath) {
        $profilename = (Get-ItemProperty -Path $outlookRegPath -Name DefaultProfile -ErrorAction SilentlyContinue).DefaultProfile
        $profileBasePath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\$profilename\9375CFF0413111d3B88A00104B2A6676"

        # Get the current user's UPN
        $userPrincipalName = whoami /upn
        Log "Current User UPN: $userPrincipalName"

        # Search for registry entries matching "Prefix" and UPN
        $profileKeys = Get-ChildItem -Path $profileBasePath -ErrorAction SilentlyContinue
        if (-not $profileKeys) {
            Log "No Outlook profile keys found under: $profileBasePath"
            Exit 1
        }

        $signaturePattern = "$signaturePrefix .* \($userPrincipalName\)"  # Match "Prefix ... (user@domain)"
        $foundMatch = $false

        foreach ($key in $profileKeys) {
            $newSignatureValue = (Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue)."New Signature"
            $replySignatureValue = (Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue)."Reply-Forward Signature"

            if ($newSignatureValue -match $signaturePattern) {
                Remove-ItemProperty -Path $key.PSPath -Name "New Signature" -Force -ErrorAction SilentlyContinue
                Log "Successfully removed 'New Signature' registry key: $newSignatureValue"
                $foundMatch = $true
            }

            if ($replySignatureValue -match $signaturePattern) {
                Remove-ItemProperty -Path $key.PSPath -Name "Reply-Forward Signature" -Force -ErrorAction SilentlyContinue
                Log "Successfully removed 'Reply-Forward Signature' registry key: $replySignatureValue"
                $foundMatch = $true
            }
        }

        if (-not $foundMatch) {
            Log "No matching signature registry entries found for user."
        }
    }
    else {
        Log "Outlook profile registry path not found."
    }
}
catch {
    Log "Failed to remove signature registry keys. Error: $_"
}




Log "Signature uninstallation completed."

# Stop logging and exit
Stop-Transcript
Exit 0
