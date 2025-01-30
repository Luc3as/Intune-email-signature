# 📧 **Intune Email Signature Installer**

## 🚀 **Overview**

Intune Email Signature Installer is a **PowerShell-based automated solution** for deploying and managing email signatures across an organization. The script retrieves user details from a **centralized Excel sheet stored on Google Drive**, dynamically generates signatures, and updates Outlook settings.

## 🎯 **Features**

✅ **Automated Installation** – Signatures are deployed without user intervention.  
✅ **Excel-Driven Data** – Signatures are filled with user-specific data from a shared Google Sheet.  
✅ **Dynamic Signature Formatting** – Custom placeholders replaced with user details.  
✅ **Supports Intune Deployment** – Easily deploy via **Microsoft Intune**.  
✅ **Automatic Updates** – Ensures signatures remain up to date.  
✅ **Hidden Execution** – Runs silently in the background.

---

## 📌 **Requirements**

- Windows 10/11 with **Microsoft Outlook** installed.
- PowerShell 5.1 or later.
- Access to a **shared Excel sheet** on Google Drive.
- **Admin rights** (required for registry modifications).

---

## 🛠 **Configuration Guide**

### **1️⃣ Setting Up the Shared Excel Sheet**

The script pulls data from an Excel sheet stored on **Google Drive**. Follow these steps to configure it:

1. **Create a Google Sheet** and populate it with the following columns:
   ```plaintext
   userPrincipalName, displayName, givenName, surname, mail, jobTitle,
   department, usageLocation, streetAddress, country, officeLocation,
   city, postalCode, telephoneNumber, mobilePhone, companyName,
   setNewEmail, setReplyEmail
   ```
2. Click **File > Download > Microsoft Excel (.xlsx)** and save the file.
3. Upload the Excel file to **Google Drive**.
4. **Get the Google Drive File ID:**

   - Right-click the file and select **Get link**.
   - Copy the URL, which looks like:
     ```plaintext
     https://drive.google.com/file/d/superSecretId/view?usp=sharing
     ```
   - The **File ID** is the part after `/d/` and before `/view`.
   - **Example File ID:** `superSecretId`

5. Replace the `googleDriveFileId` variable in the script:
   ```powershell
   $googleDriveFileId = "superSecretId"
   ```

---

### **2️⃣ Deploying via Microsoft Intune**

📷 _Refer to the screenshot for Intune setup_

#### **Create a Win32 App in Intune**

1. **Package the script:**
   - Use **Microsoft Win32 Content Prep Tool** (`IntuneWinAppUtil.exe`).
   - Convert the script into a `.intunewin` package.
2. **Package the script using `makeapp.cmd`:**
3. **Upload to Intune:**
   - Go to **Microsoft Endpoint Manager Admin Center**.
   - Navigate to **Apps > Windows > Add**.
   - Select **Win32 app** and upload the `.intunewin` file.
4. **Configure the App Settings:**
   - **Install command:**
     ```plaintext
     powershell.exe -noprofile -WindowStyle Hidden -executionpolicy bypass -file .\install_email_signature.ps1
     ```
   - **Uninstall command:**
     ```plaintext
     powershell.exe -noprofile -WindowStyle Hidden -executionpolicy bypass -file .\uninstall_email_signature.ps1
     ```
   - **Return codes:**
     - `0` - Success
     - `1707` - Success
     - `3010` - Soft reboot
     - `1641` - Hard reboot
     - `1618` - Retry

#### **Setting Up Detection Script**

- Use attached **detection.ps1** script from source folder

---

## 🚀 **Usage**

### **Manual Execution (for testing/debugging)**

```powershell
powershell.exe -noprofile -executionpolicy bypass -file install_email_signature.ps1
```

### **Uninstall Signature Manually**

```powershell
powershell.exe -noprofile -executionpolicy bypass -file uninstall_email_signature.ps1
```

---

## ❓ **Troubleshooting**

- **Script not executing?** Run as **Administrator**.
- **Google Drive file not found?** Ensure the **File ID** is correct and accessible.
- **Signature not updating?** Check the **Excel sheet formatting** and confirm data accuracy.
- **Installation failing in Intune?** Review Intune logs under `C:\ProgramData\Microsoft\Intune\Logs`.

---

🚀 _Enjoy seamless email signature deployment!_ 🎯

### If You liked my work, You can buy me a coffee :)

<a class="" target="_blank" href="https://www.buymeacoffee.com/luc3as"><img src="https://lukasporubcan.sk/images/buymeacoffee.png" alt="Buy Me A Coffee" style="max-width: 217px !important;"></a>

### Or send some crypto

<a class="" target="_blank" href="https://lukasporubcan.sk/donate"><img src="https://lukasporubcan.sk/images/donatebitcoin.png" alt="Donate Bitcoin" style="max-width: 217px !important;"></a>
