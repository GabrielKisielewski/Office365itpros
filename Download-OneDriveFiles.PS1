# Download-OneDriveFiles.PS1
# A script to download all files from a user's OneDrive for Business account to a local folder
# V1.0 12-Dec-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Download-OneDriveFiles.PS1

<#
.SYNOPSIS
    Downloads all files from a user's OneDrive for Business account to a local folder.

.DESCRIPTION
    This script connects to Microsoft Graph and downloads all files from a specified user's 
    OneDrive for Business account. The script maintains the folder structure and provides 
    progress information during the download process.

.PARAMETER UserPrincipalName
    The User Principal Name (email address) of the account whose OneDrive files to download.
    If not specified, the script will download files from the signed-in user's OneDrive.

.PARAMETER DownloadPath
    The local path where files should be downloaded. If not specified, files will be 
    downloaded to a 'OneDrive-Download' folder in the user's Downloads directory.

.PARAMETER LogFilePath
    The path to the log file. If not specified, a log file will be created in the 
    download directory with the name 'Download-Log-<timestamp>.txt'.

.EXAMPLE
    .\Download-OneDriveFiles.PS1
    Downloads files from the signed-in user's OneDrive to the default location.

.EXAMPLE
    .\Download-OneDriveFiles.PS1 -UserPrincipalName user@contoso.com -DownloadPath "C:\Backup\OneDrive"
    Downloads files from user@contoso.com's OneDrive to C:\Backup\OneDrive.

.NOTES
    Requires the Microsoft.Graph PowerShell module and appropriate permissions.
    The signed-in user must have access to read the target OneDrive account.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory = $false)]
    [string]$DownloadPath,
    
    [Parameter(Mandatory = $false)]
    [string]$LogFilePath
)

function Write-Log {
    <#
    .SYNOPSIS
        Writes a log message to both console and log file.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogMessage = "[$Timestamp] [$Level] $Message"
    
    # Write to console with color
    switch ($Level) {
        'Info'    { Write-Host $Message -ForegroundColor Cyan }
        'Warning' { Write-Host $Message -ForegroundColor Yellow }
        'Error'   { Write-Host $Message -ForegroundColor Red }
        'Success' { Write-Host $Message -ForegroundColor Green }
        default   { Write-Host $Message }
    }
    
    # Write to log file if path is set
    if ($Script:LogFilePath) {
        try {
            Add-Content -Path $Script:LogFilePath -Value $LogMessage -ErrorAction Stop
        } catch {
            Write-Host "Warning: Failed to write to log file: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function Sanitize-FileName {
    <#
    .SYNOPSIS
        Sanitizes a file or folder name to prevent directory traversal and invalid characters.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    
    # Remove path separators and invalid characters
    $SanitizedName = $Name -replace '[\\/:*?\"<>|]', '_'
    # Remove any leading or trailing dots or spaces
    $SanitizedName = $SanitizedName.Trim('. ')
    # Prevent empty names
    if ([string]::IsNullOrWhiteSpace($SanitizedName)) {
        $SanitizedName = "unnamed"
    }
    return $SanitizedName
}

function Download-DriveItems {
    <#
    .SYNOPSIS
        Recursively downloads files from a OneDrive folder.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Drive,
        
        [Parameter(Mandatory = $true)]
        $FolderId,
        
        [Parameter(Mandatory = $true)]
        [string]$LocalPath,
        
        [Parameter(Mandatory = $false)]
        [int]$Depth = 0,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxDepth = 100
    )
    
    # Prevent excessive recursion depth
    if ($Depth -ge $MaxDepth) {
        Write-Log -Message ("Warning: Maximum folder depth ({0}) reached. Skipping deeper folders." -f $MaxDepth) -Level Warning
        return
    }
    
    try {
        # Get items in the current folder
        $Items = Get-MgDriveItemChild -DriveId $Drive -DriveItemId $FolderId -All -ErrorAction Stop
        
        # Split into files and folders
        $Folders = $Items | Where-Object { $_.Folder }
        $Files = $Items | Where-Object { $_.File }
        
        # Download files in the current folder
        foreach ($File in $Files) {
            $Script:TotalFiles++
            $SafeFileName = Sanitize-FileName -Name $File.Name
            $FilePath = Join-Path -Path $LocalPath -ChildPath $SafeFileName
            
            Write-Log -Message ("Downloading file: {0} ({1})" -f $File.Name, (Format-FileSize $File.Size)) -Level Info
            
            try {
                # Get the download URL for the file
                $DownloadUrl = "https://graph.microsoft.com/v1.0/drives/{0}/items/{1}/content" -f $Drive, $File.Id
                
                # Download the file
                Invoke-MgGraphRequest -Uri $DownloadUrl -Method GET -OutputFilePath $FilePath -ErrorAction Stop
                
                $Script:DownloadedFiles++
                $Script:TotalBytesDownloaded += $File.Size
                
                # Set file timestamps to match OneDrive
                if ($File.CreatedDateTime -or $File.LastModifiedDateTime) {
                    $FileItem = Get-Item $FilePath
                    if ($File.CreatedDateTime) {
                        $FileItem.CreationTime = $File.CreatedDateTime
                    }
                    if ($File.LastModifiedDateTime) {
                        $FileItem.LastWriteTime = $File.LastModifiedDateTime
                    }
                }
                
                $DownloadInfo = [PSCustomObject]@{
                    FileName         = $File.Name
                    LocalPath        = $FilePath
                    Size             = $File.Size
                    FormattedSize    = Format-FileSize $File.Size
                    Created          = $File.CreatedDateTime
                    Modified         = $File.LastModifiedDateTime
                    DownloadStatus   = "Success"
                }
                $Script:DownloadReport.Add($DownloadInfo)
                
            } catch {
                Write-Log -Message ("  Error downloading file {0}: {1}" -f $File.Name, $_.Exception.Message) -Level Error
                $Script:FailedFiles++
                
                $DownloadInfo = [PSCustomObject]@{
                    FileName         = $File.Name
                    LocalPath        = $FilePath
                    Size             = $File.Size
                    FormattedSize    = Format-FileSize $File.Size
                    Created          = $File.CreatedDateTime
                    Modified         = $File.LastModifiedDateTime
                    DownloadStatus   = "Failed: $($_.Exception.Message)"
                }
                $Script:DownloadReport.Add($DownloadInfo)
            }
        }
        
        # Process subfolders recursively
        foreach ($Folder in $Folders) {
            $Script:TotalFolders++
            $SafeFolderName = Sanitize-FileName -Name $Folder.Name
            $SubfolderPath = Join-Path -Path $LocalPath -ChildPath $SafeFolderName
            
            # Create subfolder if it doesn't exist
            if (-not (Test-Path -Path $SubfolderPath)) {
                New-Item -Path $SubfolderPath -ItemType Directory -Force | Out-Null
                Write-Log -Message ("Created folder: {0}" -f $SubfolderPath) -Level Warning
            }
            
            Write-Log -Message ("Processing folder: {0} ({1} items)" -f $Folder.Name, $Folder.Folder.ChildCount) -Level Success
            
            # Recursively download from this folder
            Download-DriveItems -Drive $Drive -FolderId $Folder.Id -LocalPath $SubfolderPath -Depth ($Depth + 1) -MaxDepth $MaxDepth
        }
        
    } catch {
        Write-Log -Message ("Error processing folder: {0}" -f $_.Exception.Message) -Level Error
    }
}

function Format-FileSize {
    <#
    .SYNOPSIS
        Formats a file size in bytes to a human-readable string.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [long]$InFileSize
    )
    
    if ($InFileSize -lt 1KB) {
        return "{0} B" -f $InFileSize
    } elseif ($InFileSize -lt 1MB) {
        return "{0:n2} KB" -f ($InFileSize / 1KB)
    } elseif ($InFileSize -lt 1GB) {
        return "{0:n2} MB" -f ($InFileSize / 1MB)
    } else {
        return "{0:n2} GB" -f ($InFileSize / 1GB)
    }
}

# Main script execution starts here

# Initialize log file path (will be set later after determining download path)
$Script:LogFilePath = $null

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "OneDrive for Business File Download Script" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

# Initialize counters and report list
$Script:TotalFiles = 0
$Script:DownloadedFiles = 0
$Script:FailedFiles = 0
$Script:TotalFolders = 0
$Script:TotalBytesDownloaded = 0
$Script:DownloadReport = [System.Collections.Generic.List[Object]]::new()

# Connect to Microsoft Graph with required permissions
Write-Log -Message "Connecting to Microsoft Graph..." -Level Warning
try {
    Connect-MgGraph -Scopes "Files.Read.All", "User.Read.All" -NoWelcome -ErrorAction Stop
} catch {
    Write-Log -Message "Error: Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level Error
    exit 1
}

# Determine which user's OneDrive to access
if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
    $UserPrincipalName = (Get-MgContext).Account
    Write-Log -Message ("Using signed-in user's account: {0}" -f $UserPrincipalName) -Level Success
} else {
    Write-Log -Message ("Accessing OneDrive for user: {0}" -f $UserPrincipalName) -Level Success
}

# Get the user's OneDrive account
Write-Log -Message "Retrieving OneDrive account information..." -Level Warning
try {
    $OneDriveInfo = Get-MgUserDefaultDrive -UserId $UserPrincipalName -ErrorAction Stop
} catch {
    Write-Log -Message ("Error: Could not retrieve OneDrive account for {0}: {1}" -f $UserPrincipalName, $_.Exception.Message) -Level Error
    exit 1
}

if (-not $OneDriveInfo) {
    Write-Log -Message ("Error: No OneDrive for Business account found for {0}" -f $UserPrincipalName) -Level Error
    exit 1
}

Write-Log -Message ("Found OneDrive account: {0}" -f $OneDriveInfo.Owner.User.DisplayName) -Level Success
Write-Host ("  WebUrl: {0}" -f $OneDriveInfo.WebUrl) -ForegroundColor Gray
Write-Host ("  Total Quota: {0}" -f (Format-FileSize $OneDriveInfo.Quota.Total)) -ForegroundColor Gray
Write-Host ("  Used Quota: {0}" -f (Format-FileSize $OneDriveInfo.Quota.Used)) -ForegroundColor Gray
Write-Host ""

# Determine download path
if ([string]::IsNullOrWhiteSpace($DownloadPath)) {
    $DownloadPath = Join-Path -Path ([Environment]::GetFolderPath("Downloads")) -ChildPath "OneDrive-Download"
} else {
    # Validate user-provided path exists or can be created
    try {
        $DownloadPath = [System.IO.Path]::GetFullPath($DownloadPath)
    } catch {
        Write-Log -Message ("Error: Invalid download path specified: {0}" -f $_.Exception.Message) -Level Error
        exit 1
    }
}

# Sanitize the download path by adding the user's display name
$SafeUserName = Sanitize-FileName -Name $OneDriveInfo.Owner.User.DisplayName
$DownloadPath = Join-Path -Path $DownloadPath -ChildPath $SafeUserName

# Create the download directory if it doesn't exist
if (-not (Test-Path -Path $DownloadPath)) {
    Write-Log -Message ("Creating download directory: {0}" -f $DownloadPath) -Level Warning
    New-Item -Path $DownloadPath -ItemType Directory -Force | Out-Null
}

# Initialize log file if not provided
if ([string]::IsNullOrWhiteSpace($LogFilePath)) {
    $LogFileName = "Download-Log-{0}.txt" -f (Get-Date -Format 'yyyyMMdd-HHmmss')
    $Script:LogFilePath = Join-Path -Path $DownloadPath -ChildPath $LogFileName
} else {
    $Script:LogFilePath = $LogFilePath
}

# Create log file and write header
try {
    $LogHeader = @"
==================================================
OneDrive for Business File Download Log
==================================================
Script Version: 1.0
Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
User Account: $UserPrincipalName
OneDrive Owner: $($OneDriveInfo.Owner.User.DisplayName)
Download Path: $DownloadPath
==================================================

"@
    Set-Content -Path $Script:LogFilePath -Value $LogHeader -ErrorAction Stop
    Write-Log -Message ("Log file created: {0}" -f $Script:LogFilePath) -Level Success
} catch {
    Write-Host ("Warning: Failed to create log file: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    $Script:LogFilePath = $null
}

Write-Log -Message ("Download location: {0}" -f $DownloadPath) -Level Success
Write-Host ""
Write-Log -Message "Starting download process..." -Level Warning
Write-Host ""

# Record start time
$StartTime = Get-Date

# Start downloading from the root folder
try {
    Download-DriveItems -Drive $OneDriveInfo.Id -FolderId "root" -LocalPath $DownloadPath
} catch {
    Write-Log -Message ("Error during download process: {0}" -f $_.Exception.Message) -Level Error
}

# Record end time and calculate duration
$EndTime = Get-Date
$Duration = $EndTime - $StartTime

# Generate summary report
Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "Download Complete - Summary" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ("OneDrive Account: {0}" -f $OneDriveInfo.Owner.User.DisplayName) -ForegroundColor White
Write-Host ("Download Location: {0}" -f $DownloadPath) -ForegroundColor White
Write-Host ""
Write-Host ("Total Files Found: {0}" -f $Script:TotalFiles) -ForegroundColor White
Write-Host ("Files Downloaded Successfully: {0}" -f $Script:DownloadedFiles) -ForegroundColor Green
Write-Host ("Files Failed to Download: {0}" -f $Script:FailedFiles) -ForegroundColor $(if ($Script:FailedFiles -gt 0) { "Red" } else { "White" })
Write-Host ("Total Folders Processed: {0}" -f $Script:TotalFolders) -ForegroundColor White
Write-Host ("Total Data Downloaded: {0}" -f (Format-FileSize $Script:TotalBytesDownloaded)) -ForegroundColor White
Write-Host ("Duration: {0}:{1:D2}:{2:D2}" -f $Duration.Hours, $Duration.Minutes, $Duration.Seconds) -ForegroundColor White
Write-Host ""

# Write summary to log file
if ($Script:LogFilePath) {
    $LogSummary = @"

==================================================
Download Summary
==================================================
End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Total Files Found: $($Script:TotalFiles)
Files Downloaded Successfully: $($Script:DownloadedFiles)
Files Failed to Download: $($Script:FailedFiles)
Total Folders Processed: $($Script:TotalFolders)
Total Data Downloaded: $(Format-FileSize $Script:TotalBytesDownloaded)
Duration: $($Duration.Hours):$($Duration.Minutes.ToString('00')):$($Duration.Seconds.ToString('00'))
==================================================
"@
    Add-Content -Path $Script:LogFilePath -Value $LogSummary
}

# Export detailed report to CSV
if ($Script:DownloadReport.Count -gt 0) {
    $ReportFile = Join-Path -Path $DownloadPath -ChildPath "Download-Report.csv"
    $Script:DownloadReport | Export-Csv -Path $ReportFile -NoTypeInformation -Encoding UTF8
    Write-Log -Message ("Detailed download report saved to: {0}" -f $ReportFile) -Level Success
}

# Show failed downloads if any
if ($Script:FailedFiles -gt 0) {
    Write-Host ""
    Write-Host "Failed Downloads:" -ForegroundColor Red
    $Script:DownloadReport | Where-Object { $_.DownloadStatus -ne "Success" } | Format-Table FileName, DownloadStatus -AutoSize
}

# Log completion
if ($Script:LogFilePath) {
    Write-Log -Message ("Log file saved to: {0}" -f $Script:LogFilePath) -Level Success
}

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
