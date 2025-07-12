# Function to write to log files
function Write-Log {
    param (
        [string]$LogFile,
        [string]$Message,
        [string]$Status
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - [$Status] - $Message" | Out-File -FilePath $LogFile -Append
}