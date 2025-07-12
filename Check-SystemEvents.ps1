#Requires -RunAsAdministrator

function Check-SystemEvents {
    param (
        [string]$LogFile
    )
    try {
        . "$PSScriptRoot\Write-Log.ps1"
        $results = @()
        $timeFrame = (Get-Date).AddDays(-7)
        
        # Check Security log and auditing
        if (Get-WinEvent -ListLog "Security" -ErrorAction SilentlyContinue) {
            # Verify auditing is enabled for Logon category
            $auditPol = auditpol /get /category:Logon 2>&1
            if ($auditPol -match "Success\s+Enable" -or $auditPol -match "Failure\s+Enable") {
                $securityEventIds = @(4625, 4740, 4672)
                foreach ($eventId in $securityEventIds) {
                    try {
                        $securityEvents = Get-WinEvent -LogName "Security" -FilterHashtable @{
                            StartTime = $timeFrame
                            ID = $eventId
                        } -ErrorAction SilentlyContinue
                        foreach ($event in $securityEvents) {
                            $status = "Normal"
                            $details = ""
                            switch ($event.Id) {
                                4625 {
                                    $details = "Account: $($event.Properties[5].Value) from $($event.Properties[19].Value)"
                                    $ip = $event.Properties[19].Value
                                    $count = ($securityEvents | Where-Object { $_.Id -eq 4625 -and $_.Properties[19].Value -eq $ip }).Count
                                    if ($count -gt 10) { $status = "Suspicious" }
                                }
                                4740 { $details = "Account: $($event.Properties[0].Value)"; $status = "Suspicious" }
                                4672 { $details = "Account: $($event.Properties[1].Value) assigned $($event.Properties[3].Value)"; $status = "Suspicious" }
                            }
                            Write-Log -LogFile $LogFile -Message "Event ID $($event.Id): $details" -Status $status
                            $results += [PSCustomObject]@{EventID=$event.Id; Time=$event.TimeCreated; Description=$event.Message.Split("`n")[0]; Status=$status; Details=$details}
                        }
                    } catch {
                        Write-Log -LogFile $LogFile -Message "Error retrieving Security event ID ${eventId}: $_" -Status "Warning"
                    }
                }
            } else {
                Write-Log -LogFile $LogFile -Message "Security log auditing is disabled. Enable with: auditpol /set /category:'Logon' /success:enable /failure:enable" -Status "Warning"
            }
        } else {
            Write-Log -LogFile $LogFile -Message "Security log not accessible." -Status "Warning"
        }

        # Check System log
        if (Get-WinEvent -ListLog "System" -ErrorAction SilentlyContinue) {
            $systemEventIds = @(4688, 7045)
            foreach ($eventId in $systemEventIds) {
                try {
                    $systemEvents = Get-WinEvent -LogName "System" -FilterHashtable @{
                        StartTime = $timeFrame
                        ID = $eventId
                    } -ErrorAction SilentlyContinue
                    foreach ($event in $systemEvents) {
                        $status = "Normal"
                        $details = ""
                        switch ($event.Id) {
                            4688 {
                                $process = $event.Properties[5].Value
                                $details = "Process: $process"
                                if ($process -match "cmd.exe|powershell.exe" -and $event.Properties[8].Value -match "Temp|AppData") {
                                    $status = "Suspicious"
                                }
                            }
                            7045 { $details = "Service: $($event.Properties[0].Value) - Path: $($event.Properties[1].Value)"; $status = "Suspicious" }
                        }
                        Write-Log -LogFile $LogFile -Message "Event ID $($event.Id): $details" -Status $status
                        $results += [PSCustomObject]@{EventID=$event.Id; Time=$event.TimeCreated; Description=$event.Message.Split("`n")[0]; Status=$status; Details=$details}
                    }
                } catch {
                    Write-Log -LogFile $LogFile -Message "Error retrieving System event ID ${eventId}: $_" -Status "Warning"
                }
            }
        } else {
            Write-Log -LogFile $LogFile -Message "System log not accessible." -Status "Warning"
        }

        if (-not $results) {
            Write-Log -LogFile $LogFile -Message "No suspicious events found in the last 7 days." -Status "Info"
            $results += [PSCustomObject]@{EventID="None"; Time=(Get-Date); Description="No events detected"; Status="Normal"; Details="No suspicious activity found"}
        }
        return $results
    } catch {
        Write-Log -LogFile $LogFile -Message "Error checking system events: $_" -Status "Error"
        return $null
    }
}