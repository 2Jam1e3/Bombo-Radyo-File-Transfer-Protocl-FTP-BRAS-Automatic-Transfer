# ============================================================
#  TransferAuditFiles.ps1
#  Reads from XP VM shared folder → Uploads to FTP
#  Schedule: Daily 10:00 AM via Task Scheduler
# ============================================================

# ─── CONFIGURATION ───────────────────────────────────────────
$VBoxManage      = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
$VMName          = "Windows XP"              # <-- Your exact VM name in VirtualBox

$XP_VM_IP        = "192.168.86.128"            # <-- Replace with your XP VM's actual IP
$SharedFolder    = "\\$XP_VM_IP\Audit"
$XP_Username     = "Automation"
$XP_Password     = "password"               # <-- Your XP login password

$FTPHost         = "ftp.bomboradyo.com"
$FTPUser         = "BOMBOMALAYBALAY"
$FTPPassword     = "yk3933rX2b"   # <-- Put your NEW password here after changing it
$FTPRemotePath   = "/Automation Log File/Bombo/Malaybalay/"
$UploadNewestOnly = $true          # true = upload only the newest generated file each run

$LogFile         = "C:\Scripts\Logs\AuditTransfer_$(Get-Date -Format 'yyyy-MM-dd').log"
# ─────────────────────────────────────────────────────────────

# --- LOGGING FUNCTION ---
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] $Message"
    Write-Host $entry
    Add-Content -Path $LogFile -Value $entry
}

function Get-EncodedFtpPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $segments = $Path.Trim("/") -split "/"
    $encodedSegments = $segments | ForEach-Object { [System.Uri]::EscapeDataString($_) }
    return ($encodedSegments -join "/")
}

function Write-FtpExceptionDetail {
    param([Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $exception = $ErrorRecord.Exception
    while ($exception.InnerException) {
        $exception = $exception.InnerException
    }

    $message = $exception.Message
    if ($exception -is [System.Net.WebException] -and $exception.Response) {
        $ftpResponse = [System.Net.FtpWebResponse]$exception.Response
        $message = "$message | FTP Status: $($ftpResponse.StatusCode) - $($ftpResponse.StatusDescription.Trim())"
        $ftpResponse.Dispose()
    }

    return $message
}

function Test-FtpLogin {
    param(
        [Parameter(Mandatory = $true)][string]$FtpHostName,
        [Parameter(Mandatory = $true)][string]$EncodedPath,
        [Parameter(Mandatory = $true)][System.Net.NetworkCredential]$Credential,
        [Parameter(Mandatory = $true)][bool]$UseSsl
    )

    $testUrl = "ftp://${FtpHostName}/${EncodedPath}/"
    $request = [System.Net.FtpWebRequest]::Create($testUrl)
    $request.Method = [System.Net.WebRequestMethods+Ftp]::PrintWorkingDirectory
    $request.Credentials = $Credential
    $request.EnableSsl = $UseSsl
    $request.UseBinary = $true
    $request.UsePassive = $true
    $request.KeepAlive = $false
    $request.Proxy = $null

    $response = [System.Net.FtpWebResponse]$request.GetResponse()
    $status = $response.StatusDescription.Trim()
    $response.Dispose()
    return $status
}

# --- ENSURE LOG DIRECTORY EXISTS ---
if (!(Test-Path "C:\Scripts\Logs")) {
    New-Item -ItemType Directory -Path "C:\Scripts\Logs" | Out-Null
}

Write-Log "=========================================="
Write-Log "  Bombo Malaybalay - Audit File Transfer  "
Write-Log "=========================================="

# --- AUTO-START XP VM IF NOT RUNNING ---
Write-Log "Checking XP VM status..."
$VMState = & "$VBoxManage" showvminfo "$VMName" --machinereadable 2>&1 | Select-String "VMState="
if ($VMState -notmatch "running") {
    Write-Log "XP VM is not running. Starting it now..."
    & "$VBoxManage" startvm "$VMName" --type headless
    Write-Log "Waiting 60 seconds for VM to boot..."
    Start-Sleep -Seconds 60
} else {
    Write-Log "XP VM is already running."
}

# --- CONNECT TO XP SHARED FOLDER ---
Write-Log "Connecting to: $SharedFolder"
net use $SharedFolder /delete /yes 2>$null
$ConnectResult = net use $SharedFolder /user:$XP_Username $XP_Password 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Log "ERROR: Could not connect to XP shared folder."
    Write-Log "Detail: $ConnectResult"
    exit 1
}
Write-Log "SUCCESS: Connected to XP shared folder."

# --- GET FILES ---
try {
    $Files = Get-ChildItem -Path $SharedFolder -File -ErrorAction Stop
} catch {
    Write-Log "ERROR: Could not read files -- $($_.Exception.Message)"
    net use $SharedFolder /delete /yes 2>$null
    exit 1
}

if ($Files.Count -eq 0) {
    Write-Log "No files found. Nothing to transfer."
    net use $SharedFolder /delete /yes 2>$null
    exit 0
}

if ($UploadNewestOnly) {
    $NewestFile = $Files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $Files = @($NewestFile)
    Write-Log "Found newest file to upload: $($NewestFile.Name) (LastWrite: $($NewestFile.LastWriteTime))"
} else {
    Write-Log "Found $($Files.Count) file(s) to upload."
}

# --- UPLOAD FILES TO FTP ---
$SuccessCount = 0
$FailCount    = 0
$FtpCredential = New-Object System.Net.NetworkCredential($FTPUser.Trim(), $FTPPassword.Trim())
$EncodedRemoteBasePath = Get-EncodedFtpPath -Path $FTPRemotePath
$UseFtps = $false
$ProtocolLabel = "FTP"

Write-Log "Checking FTP login for user '$($FTPUser.Trim())'..."
try {
    $loginStatus = Test-FtpLogin -FtpHostName $FTPHost -EncodedPath $EncodedRemoteBasePath -Credential $FtpCredential -UseSsl:$false
    Write-Log "FTP login OK (plain FTP): $loginStatus"
} catch {
    $plainErrorDetail = Write-FtpExceptionDetail -ErrorRecord $_
    Write-Log "Plain FTP login failed | $plainErrorDetail"
    Write-Log "Trying FTPS (Explicit TLS)..."

    try {
        $loginStatus = Test-FtpLogin -FtpHostName $FTPHost -EncodedPath $EncodedRemoteBasePath -Credential $FtpCredential -UseSsl:$true
        $UseFtps = $true
        $ProtocolLabel = "FTPS"
        Write-Log "FTP login OK (FTPS): $loginStatus"
    } catch {
        $ftpsErrorDetail = Write-FtpExceptionDetail -ErrorRecord $_
        Write-Log "ERROR: FTP login failed on both FTP and FTPS."
        Write-Log "Detail (FTP): $plainErrorDetail"
        Write-Log "Detail (FTPS): $ftpsErrorDetail"
        net use $SharedFolder /delete /yes 2>$null
        exit 1
    }
}

foreach ($File in $Files) {
    $EncodedFileName = [System.Uri]::EscapeDataString($File.Name)
    $RemoteURL = "ftp://${FTPHost}/${EncodedRemoteBasePath}/${EncodedFileName}"
    Write-Log "Uploading ($ProtocolLabel): $($File.Name) --> $RemoteURL"

    try {
        $FTPRequest               = [System.Net.FtpWebRequest]::Create($RemoteURL)
        $FTPRequest.Method        = [System.Net.WebRequestMethods+Ftp]::UploadFile
        $FTPRequest.Credentials   = $FtpCredential
        $FTPRequest.EnableSsl     = $UseFtps
        $FTPRequest.UseBinary     = $true
        $FTPRequest.UsePassive    = $true
        $FTPRequest.KeepAlive     = $false
        $FTPRequest.Proxy         = $null

        $FileBytes                = [System.IO.File]::ReadAllBytes($File.FullName)
        $FTPRequest.ContentLength = $FileBytes.Length

        $Stream = $FTPRequest.GetRequestStream()
        $Stream.Write($FileBytes, 0, $FileBytes.Length)
        $Stream.Close()

        $Response = $FTPRequest.GetResponse()
        Write-Log "  SUCCESS: $($File.Name) | $($Response.StatusDescription.Trim())"
        $Response.Dispose()
        $SuccessCount++

    } catch {
        $errorDetail = Write-FtpExceptionDetail -ErrorRecord $_
        Write-Log "  FAILED: $($File.Name) | $errorDetail"
        $FailCount++
    }
}

# --- DISCONNECT ---
net use $SharedFolder /delete /yes 2>$null
Write-Log "Disconnected from shared folder."
Write-Log "------------------------------------------"
Write-Log "Summary: $SuccessCount uploaded, $FailCount failed."
Write-Log "=========================================="