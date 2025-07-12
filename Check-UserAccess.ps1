#Requires -RunAsAdministrator

function Check-UserAccess {
    param (
        [string]$LogFile
    )
    try {
        . "$PSScriptRoot\Write-Log.ps1"
        $results = @()
        $users = Get-LocalUser -ErrorAction Stop
        foreach ($user in $users) {
            $status = if ($user.Enabled) { "Enabled" } else { "Disabled" }
            Write-Log -LogFile $LogFile -Message "User: $($user.Name) - Status: $status" -Status $status
            $results += [PSCustomObject]@{Name=$user.Name; Status=$status; Details="SID: $($user.SID)"}
        }
        return $results
    } catch {
        Write-Log -LogFile $LogFile -Message "Error checking users: $_" -Status "Error"
        return $null
    }
}