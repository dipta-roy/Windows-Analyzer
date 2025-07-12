#Requires -RunAsAdministrator

function Check-Certificates {
    param (
        [string]$LogFile
    )
    try {
        . "$PSScriptRoot\Write-Log.ps1"
        $results = @()
        $certs = Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction Stop
        foreach ($cert in $certs) {
            Write-Log -LogFile $LogFile -Message "Certificate: $($cert.Subject) - Expires: $($cert.NotAfter)" -Status "Info"
            $results += [PSCustomObject]@{Subject=$cert.Subject; Expires=$cert.NotAfter; Status="Valid"; Details="Thumbprint: $($cert.Thumbprint)"}
        }
        return $results
    } catch {
        Write-Log -LogFile $LogFile -Message "Error checking certificates: $_" -Status "Error"
        return $null
    }
}