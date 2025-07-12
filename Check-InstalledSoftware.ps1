#Requires -RunAsAdministrator

function Check-InstalledSoftware {
    param (
        [string]$LogFile
    )
    try {
        . "$PSScriptRoot\Write-Log.ps1"
        $results = @()
        $software = Get-WmiObject -Class Win32_Product -ErrorAction Stop
        foreach ($app in $software) {
            Write-Log -LogFile $LogFile -Message "Software: $($app.Name) - Version: $($app.Version)" -Status "Info"
            $results += [PSCustomObject]@{Name=$app.Name; Version=$app.Version; Status="Installed"; Details="Vendor: $($app.Vendor)"}
        }
        return $results
    } catch {
        Write-Log -LogFile $LogFile -Message "Error checking software: $_" -Status "Error"
        return $null
    }
}