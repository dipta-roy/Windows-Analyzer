#Requires -RunAsAdministrator

function Check-Firewall {
    param (
        [string]$LogFile
    )
    try {
        . "$PSScriptRoot\Write-Log.ps1"
        $results = @()
        $rules = Get-NetFirewallRule | Where-Object { $_.Enabled -eq $true } -ErrorAction Stop
        foreach ($rule in $rules) {
            $ports = $rule | Get-NetFirewallPortFilter
            $portList = if ($ports.LocalPort) { $ports.LocalPort -join ", " } else { "Any" }
            Write-Log -LogFile $LogFile -Message "Firewall Rule: $($rule.DisplayName) - Ports: $portList" -Status "Info"
            $results += [PSCustomObject]@{Name=$rule.DisplayName; Ports=$portList; Status="Enabled"; Details="Protocol: $($ports.Protocol)"}
        }
        $portsToCheck = @(80, 443, 3389)
        $openConnections = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue
        if (-not $openConnections) {
            Write-Log -LogFile $LogFile -Message "No listening TCP connections found." -Status "Info"
        }
        foreach ($port in $portsToCheck) {
            if ($openConnections | Where-Object { $_.LocalPort -eq $port }) {
                Write-Log -LogFile $LogFile -Message "Open Port: $port" -Status "Warning"
                $results += [PSCustomObject]@{Name="Port Check"; Ports=$port; Status="Open"; Details="Port is listening"}
            } else {
                Write-Log -LogFile $LogFile -Message "Closed Port: $port" -Status "Info"
                $results += [PSCustomObject]@{Name="Port Check"; Ports=$port; Status="Closed"; Details="Port is not listening"}
            }
        }
        return $results
    } catch {
        Write-Log -LogFile $LogFile -Message "Error checking firewall: $_" -Status "Error"
        return $null
    }
}