Automates the process of retrieving audit files from a Windows XP Virtual Machine and uploading them to an FTP server.

---

## 🚀 Overview

This script performs the following tasks:

* Connects to a shared folder from a Windows XP VM
* Reads audit files from the shared directory
* Uploads files to a remote FTP server
* Logs all operations for monitoring and troubleshooting

---

## 🛠️ Features

* 🔄 Automated file transfer
* 📁 Supports uploading newest file only (optional)
* 🌐 FTP upload with encoded paths
* 🧾 Daily log file generation
* ⏰ Designed for scheduled execution (Task Scheduler)

---

## 📋 Requirements

* Windows OS
* PowerShell 5.1 or later
* Oracle VirtualBox installed
* Network access to the Windows XP VM
* FTP server access

---

## ⚙️ Configuration

Update the following variables inside the script:

```powershell
$VBoxManage      = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
$VMName          = "Windows XP"

$XP_VM_IP        = "192.168.86.128"
$SharedFolder    = "\\$XP_VM_IP\Audit"
$XP_Username     = "Automation"
$XP_Password     = "password"

$FTPHost         = "ftp.yourserver.com"
$FTPUser         = "your_username"
$FTPPassword     = "your_password"
$FTPRemotePath   = "/your/remote/path/"
$UploadNewestOnly = $true

$LogFile         = "C:\Scripts\Logs\AuditTransfer_$(Get-Date -Format 'yyyy-MM-dd').log"
```

---

## ▶️ Usage

Run the script manually:

```powershell
powershell -ExecutionPolicy Bypass -File TransferAuditFiles.ps1
```

---

## ⏰ Scheduling (Task Scheduler)

1. Open **Task Scheduler**
2. Create a new task
3. Set trigger:

   * Daily at **10:00 AM**
4. Action:

   ```plaintext
   Program/script: powershell
   Arguments: -ExecutionPolicy Bypass -File "C:\Path\To\TransferAuditFiles.ps1"
   ```

---

## 🧾 Logging

Logs are stored in:

```
C:\Scripts\Logs\
```

Each run generates a file like:

```
AuditTransfer_YYYY-MM-DD.log
```

---

## 🔐 Security Notes

⚠️ **Important:**

* Do NOT store plain-text passwords in production.
* Consider:

  * Using Windows Credential Manager
  * Encrypting credentials
  * Using secure vault solutions

---

## 📌 Notes

* Ensure the XP VM is running and accessible before execution
* Verify shared folder permissions
* Confirm FTP credentials and remote path are correct

---

## 🤝 Contributing

Feel free to fork this repository and submit pull requests for improvements.

---

## 📄 License

This project is provided as-is for internal automation use.
