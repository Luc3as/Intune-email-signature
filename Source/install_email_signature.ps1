# ==============================================================
# Intune Email Signature Uninstallation Script
# 
# Author: Lukas "Luc3as" Porubcan
# GitHub: https://github.com/luc3as
# Version: 2.8
# Created: 2025
# License: MIT 
# 
# Description:
# This script installs Outlook email signatures for users based on 
# data retrieved from an Excel file. It downloads the latest version 
# of the signature data, processes placeholders, and updates the 
# default signatures for new and reply emails.
# 
# Features:
# - Retrieves user details from an Excel file
# - Dynamically generates email signatures
# - Supports mobile number formatting 
# - Automatically sets default signatures for new and reply emails
# - Downloads the latest signature version from Google Drive
# - Ensures proper UPN matching for accurate signature assignment
# 
# ============================================================== 


# Ensure script runs as administrator
if (-Not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Exit
} 

# Define paths
$logPath = "$env:ProgramData\Microsoft\Intune\Logs"
$signatureFolder = "$($env:APPDATA)\Microsoft\Signatures"
$signaturePrefix = "BelusaFoods"
$excelFilePath = "$logPath\signatures_data.xlsx"
$googleDriveFileId = "REDACTED"
$googleDriveFileUrl = "https://drive.google.com/uc?export=download&id=$googleDriveFileId"
$Debug = $false  # Enable Debug mode


# Create logs directory if needed
if (-not (Test-Path $logPath)) {
    New-Item -ItemType Directory -Force -Path $logPath | Out-Null
}

Start-Transcript -Path "$logPath\Email_signature_Installation.log" -Force

# Logging function
function Log() {
    param ([string]$message)
    $ts = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    Write-Host "$ts $message"
}

# Stop logging function gracefully
function StopLoggingAndExit {
    param ([int]$exitCode)
    Stop-Transcript -ErrorAction SilentlyContinue
    Exit $exitCode
}

# Function to download Excel file from Google Drive
function Download-ExcelFromGoogleDrive {
    param (
        [string]$Url,
        [string]$OutputPath
    )

    try {
        Log "Downloading Excel file from Google Drive..."
        Invoke-WebRequest -Uri $Url -OutFile $OutputPath -UseBasicParsing
        Log "Download completed successfully: $OutputPath"
    }
    catch {
        Log "Error: Failed to download the Excel file. $_"
        StopLoggingAndExit 1
    }
}

# Download the Excel file
Download-ExcelFromGoogleDrive -Url $googleDriveFileUrl -OutputPath $excelFilePath

# Verify the downloaded file
if (-not (Test-Path $excelFilePath)) {
    Log "Error: Excel file not found at $excelFilePath"
    StopLoggingAndExit 1
}

# Install ImportExcel module if missing
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

# Get currently logged-in user's UPN
$whoamiUPN = whoami /upn
$whoamiUPN = $whoamiUPN.Trim()
if (-not $whoamiUPN) {
    Log "Error: Unable to determine UPN using 'whoami /upn'"
    StopLoggingAndExit 1
}
Log "User UPN: $whoamiUPN"

# Extract version from A1 (first cell)
try {
    # Read only the first row without headers
    $versionData = Import-Excel -Path $excelFilePath -WorksheetName "signatures" -StartRow 1 -EndRow 1 -NoHeader

    # Debug: Print raw version data for verification
    if ($Debug) {
        Log "Raw Version Data: $(ConvertTo-Json $versionData -Depth 1)"
    }

    # Check if versionData contains any values before accessing
    if ($versionData -and $versionData.PSObject.Properties.Count -gt 0) {
        # Extract version from first cell (A1)
        $rawVersion = $versionData.PSObject.Properties.Value[0]
        
        # Ensure version is not null before applying string operations
        if ($rawVersion -ne $null -and $rawVersion -match "^Version:") {
            $version = $rawVersion -replace "Version:", "" -replace "^\s+|\s+$", ""  # Remove prefix and trim spaces
            Log "Extracted Excel version: $version"
        }
        else {
            Log "Error: Version data is empty or does not match expected format."
            StopLoggingAndExit 1
        }
    }
    else {
        Log "Error: No version data found in Excel A1."
        StopLoggingAndExit 1
    }
}
catch {
    Log "Error: Could not retrieve version from Excel file. $_"
    StopLoggingAndExit 1
}


# Read the actual user data (starting from row 2)
$rawData = Import-Excel -Path $excelFilePath -WorksheetName "signatures" -StartRow 2

# Ensure data is not empty
if (-not $rawData -or $rawData.Count -eq 0) {
    Log "Error: No data found in Excel. Check file format."
    StopLoggingAndExit 1
}

# Manually assign headers to ensure correctness
$headers = @(
    "userPrincipalName", "displayName", "givenName", "surname", "mail", "jobTitle",
    "department", "usageLocation", "streetAddress", "country", "officeLocation",
    "city", "postalCode", "telephoneNumber", "mobilePhone", "companyName",
    "setNewEmail", "setReplyEmail"
)

# Convert data into structured objects with proper headers
$excelData = @()
foreach ($row in $rawData) {
    $obj = @{}
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $obj[$headers[$i]] = $row.PSObject.Properties.Value[$i]
    }
    $excelData += New-Object PSObject -Property $obj
}

# Debug: Print the first few rows to check if data was read correctly
if ($Debug) {
    Log "Excel Data (First 5 Rows):"
    $excelData | Select-Object -First 5 | Format-Table | Out-String | Write-Host
}

# Find user details in the Excel file
$userRow = $excelData | Where-Object { $_.userPrincipalName -eq $whoamiUPN }

if (-not $userRow) {
    Log "Error: User $whoamiUPN not found in Excel file."
    StopLoggingAndExit 1
}

Log "User found in Excel: $($userRow.displayName)"

# Extract user details using normalized headers
$userDetails = @{
    DisplayName     = [string]$userRow.displayName.Trim()
    GivenName       = [string]$userRow.givenName.Trim()
    Surname         = [string]$userRow.surname.Trim()
    Mail            = [string]$userRow.mail.Trim()
    JobTitle        = [string]$userRow.jobTitle.Trim()
    Department      = [string]$userRow.department.Trim()
    City            = [string]$userRow.city.Trim()
    Country         = if ($userRow.country -eq $null -or $userRow.country -eq "") { "Slovakia" } else { [string]$userRow.country.Trim() }
    StreetAddress   = [string]$userRow.streetAddress.Trim()
    PostalCode      = [string]$userRow.postalCode.Trim()
    TelephoneNumber = [string]$userRow.telephoneNumber.Trim()
    MobilePhone     = [string]$userRow.mobilePhone.Trim()
    CompanyName     = [string]$userRow.companyName.Trim() -replace " s r o", ", s.r.o."
    SetNewEmail     = [string]$userRow.setNewEmail.Trim()
    SetReplyEmail   = [string]$userRow.setReplyEmail.Trim()
}

# Function to clean invisible characters (LTR, RTL, etc.)
function Remove-InvisibleCharacters {
    param ([string]$inputString)
    return [regex]::Replace($inputString, '[^\u0020-\u007E]', '') # Keep only visible ASCII characters
}

# Function to convert special characters into RTF escape sequences (Windows-1250 / Central European)
function ConvertTo-RtfEncoding {
    param (
        [string]$inputString
    )

    # Convert characters into valid RTF escape sequences
    $rtfEncoded = $inputString -creplace "á", "\'e1"
    $rtfEncoded = $rtfEncoded -creplace "č", "\'e8"
    $rtfEncoded = $rtfEncoded -creplace "ď", "\'ef"
    $rtfEncoded = $rtfEncoded -creplace "é", "\'e9"
    $rtfEncoded = $rtfEncoded -creplace "ě", "\'ec"
    $rtfEncoded = $rtfEncoded -creplace "í", "\'ed"
    $rtfEncoded = $rtfEncoded -creplace "ľ", "\'be"
    $rtfEncoded = $rtfEncoded -creplace "ň", "\'f2"
    $rtfEncoded = $rtfEncoded -creplace "ó", "\'f3"
    $rtfEncoded = $rtfEncoded -creplace "ô", "\'f4"
    $rtfEncoded = $rtfEncoded -creplace "ř", "\'f8"
    $rtfEncoded = $rtfEncoded -creplace "š", "\'e6"
    $rtfEncoded = $rtfEncoded -creplace "ť", "\'bb"
    $rtfEncoded = $rtfEncoded -creplace "ú", "\'fa"
    $rtfEncoded = $rtfEncoded -creplace "ů", "\'f9"
    $rtfEncoded = $rtfEncoded -creplace "ý", "\'fd"
    $rtfEncoded = $rtfEncoded -creplace "ž", "\'e6"
    $rtfEncoded = $rtfEncoded -creplace "Á", "\'c1"
    $rtfEncoded = $rtfEncoded -creplace "Č", "\'c8"
    $rtfEncoded = $rtfEncoded -creplace "Ď", "\'cf"
    $rtfEncoded = $rtfEncoded -creplace "É", "\'c9"
    $rtfEncoded = $rtfEncoded -creplace "Ě", "\'cc"
    $rtfEncoded = $rtfEncoded -creplace "Í", "\'cd"
    $rtfEncoded = $rtfEncoded -creplace "Ľ", "\'a5"
    $rtfEncoded = $rtfEncoded -creplace "Ň", "\'d2"
    $rtfEncoded = $rtfEncoded -creplace "Ó", "\'d3"
    $rtfEncoded = $rtfEncoded -creplace "Ô", "\'d4"
    $rtfEncoded = $rtfEncoded -creplace "Ř", "\'d8"
    $rtfEncoded = $rtfEncoded -creplace "Š", "\'a6"
    $rtfEncoded = $rtfEncoded -creplace "Ť", "\'de"
    $rtfEncoded = $rtfEncoded -creplace "Ú", "\'db"
    $rtfEncoded = $rtfEncoded -creplace "Ů", "\'d9"
    $rtfEncoded = $rtfEncoded -creplace "Ý", "\'dd"
    $rtfEncoded = $rtfEncoded -creplace "Ž", "\'c6"

    return $rtfEncoded
}



# Normalize Mobile Number
$userDetails.MobilePhone = Remove-InvisibleCharacters($userDetails.MobilePhone.ToString().Trim())

# Apply padding to MobilePhone (General international format)
if ($userDetails.MobilePhone -match '^\+(\d{1,3})(\d{6,12})$') {
    # Extract country code and main number
    $countryCode = $matches[1]
    $mainNumber = $matches[2]

    # Determine formatting based on the length of the main number
    if ($mainNumber.Length -eq 9) {
        # Format for 9-digit main numbers (e.g., Slovakia +421 XXX XXX XXX)
        $formattedNumber = $mainNumber -replace '(\d{3})(\d{3})(\d{3})', '$1 $2 $3'
    }
    elseif ($mainNumber.Length -eq 10) {
        # Format for 10-digit main numbers (e.g., Czech Republic +420 XXX XXX XXXX)
        $formattedNumber = $mainNumber -replace '(\d{3})(\d{3})(\d{4})', '$1 $2 $3'
    }
    elseif ($mainNumber.Length -eq 7) {
        # Format for 7-digit numbers (landline numbers)
        $formattedNumber = $mainNumber -replace '(\d{3})(\d{4})', '$1 $2'
    }
    else {
        # Default: Keep the main number unchanged
        $formattedNumber = $mainNumber
    }

    # Combine country code with formatted number
    $userDetails.Mobile_padded = "+$countryCode $formattedNumber"
}
else {
    $userDetails.Mobile_padded = $userDetails.MobilePhone  # Keep as is if format is different
}


# Determine if we need to set new and reply signatures
$setNewSignature = $false
$setReplySignature = $false

if ($userDetails.SetNewEmail -eq "Yes") {
    $setNewSignature = $true
}

if ($userDetails.SetReplyEmail -eq "Yes") {
    $setReplySignature = $true
}

# Debug: Print extracted user details
if ($Debug) {
    Log "Extracted User Data: $(ConvertTo-Json $userDetails -Depth 1)"
}



# Remove old signatures before copying new ones
Get-ChildItem -Path $signatureFolder -Filter "$signaturePrefix*" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# Define signature names
$signatureBaseName = "$signaturePrefix $version ($($whoamiUPN))"
$signatureSubfolderName = "$signatureBaseName`_files"
$signatureSubfolderPath = Join-Path -Path $signatureFolder -ChildPath $signatureSubfolderName

# Create subfolder for signature files
if (-not (Test-Path $signatureSubfolderPath)) {
    New-Item -Path $signatureSubfolderPath -ItemType Directory -Force
}

# Create version file
$versionFileName = "$signaturePrefix $version version.txt"
$versionFilePath = Join-Path -Path $signatureFolder -ChildPath $versionFileName
$version | Out-File -FilePath $versionFilePath -Encoding UTF8 -Force
Log "Version file created: $versionFilePath with content: $version"
# Define paths
$sourceFilesPath = "$PSScriptRoot\Files"
$sourceSubfolderPath = "$sourceFilesPath\signature_files"

# Get all signature files from the source folder
$signatureFiles = Get-ChildItem -Path $sourceFilesPath -Recurse

# Copy and process .htm, .rtf, .txt files to the main signature folder
foreach ($signatureFile in $signatureFiles) {
    if ($signatureFile.Name -match "\.(htm|rtf|txt)$") {
        # Read file content
        $signatureFileContent = Get-Content -Path $signatureFile.FullName -Raw

        # Replace placeholders with user details
        foreach ($key in $userDetails.Keys) {
            $signatureFileContent = $signatureFileContent -replace "%$key%", $userDetails[$key]
        }

        # Replace the placeholder %source_files_dir% with the correct folder name for images
        $signatureFileContent = $signatureFileContent -replace "%source_files_dir%", $signatureSubfolderName

        # If the file is RTF, convert diacritics to RTF escape sequences
        if ($signatureFile.Extension -match "\.rtf") {
            $signatureFileContent = ConvertTo-RtfEncoding $signatureFileContent
        }

        # Save the modified content in the Signatures directory
        $newFileName = "$signatureBaseName$($signatureFile.Extension)"
        $newFilePath = Join-Path -Path $signatureFolder -ChildPath $newFileName
        
        # Use UTF-8 for .htm and .txt, and RTF-encoded content for .rtf
        if ($signatureFile.Extension -match "\.htm|\.txt") {
            Set-Content -Path $newFilePath -Value $signatureFileContent -Encoding UTF8 -Force
        }
        elseif ($signatureFile.Extension -match "\.rtf") {
            Set-Content -Path $newFilePath -Value $signatureFileContent -Encoding ASCII -Force
        }

        Log "Copied and modified: $newFileName"
    }
}

# Copy all contents of "signature_files" to the renamed _files subdirectory
if (Test-Path $sourceSubfolderPath) {
    Copy-Item -Path "$sourceSubfolderPath\*" -Destination $signatureSubfolderPath -Recurse -Force
    Log "Copied contents of 'signature_files' to: $signatureSubfolderPath"
}
else {
    Log "Warning: Source folder 'signature_files' not found!"
}

Log "Signature files successfully installed with directory name: $signatureSubfolderName"


# Set Outlook registry keys for default signature
try {
    $profilename = (Get-ItemProperty -Path "hkcu:\SOFTWARE\Microsoft\Office\16.0\Outlook" -Name DefaultProfile -ErrorAction Stop).DefaultProfile
    $profilepath = Get-ItemProperty -Path "hkcu:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\$profilename\9375CFF0413111d3B88A00104B2A6676\*" -ErrorAction Stop |
    Where-Object { $_."Account Name" -eq $userDetails.Mail } |
    Select-Object -ExpandProperty PSPath

    if ($setNewSignature) {
        New-ItemProperty -Path $profilepath -Name "New Signature" -Value $signatureBaseName -Force -ErrorAction Stop
    }

    if ($setReplySignature) {
        New-ItemProperty -Path $profilepath -Name "Reply-Forward Signature" -Value $signatureBaseName -Force -ErrorAction Stop
    }
}
catch {
    Write-Host "Warning: Failed to set Outlook registry keys. $_"
}

# Cleanup: Remove the temporary Excel file
Remove-Item -Path $excelFilePath -Force -ErrorAction SilentlyContinue

StopLoggingAndExit 0

