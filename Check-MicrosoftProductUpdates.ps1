#Requires -RunAsAdministrator

function Check-MicrosoftProductUpdates {
    param (
        [string]$LogFile
    )
    try {
        . "$PSScriptRoot\Write-Log.ps1"
        $results = @()
        $processedUpdates = @()

        # Check Windows Update service status
        $wuService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
        if ($wuService) {
            if ($wuService.Status -ne 'Running') {
                Write-Log -LogFile $LogFile -Message "Windows Update service is not running. Attempting to restart." -Status "Warning"
                try {
                    Restart-Service -Name wuauserv -Force -ErrorAction Stop
                    Write-Log -LogFile $LogFile -Message "Windows Update service restarted." -Status "Info"
                } catch {
                    Write-Log -LogFile $LogFile -Message "Failed to restart Windows Update service: $_" -Status "Error"
                    return $null
                }
            } else {
                Write-Log -LogFile $LogFile -Message "Windows Update service is running." -Status "Info"
            }
        } else {
            Write-Log -LogFile $LogFile -Message "Windows Update service not found." -Status "Error"
            return $null
        }

        # Initialize COM object
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager
        $updateSearcher = $updateSession.CreateUpdateSearcher()

        # Register Microsoft Update service
        $muServiceId = "7971f918-a847-4430-9279-4a52d1efe18d"
        $serviceRegistered = $updateServiceManager.Services | Where-Object { $_.ServiceID -eq $muServiceId }
        if (-not $serviceRegistered) {
            Write-Log -LogFile $LogFile -Message "Registering Microsoft Update service." -Status "Info"
            try {
                $updateServiceManager.AddService2($muServiceId, 7, "") | Out-Null
                Write-Log -LogFile $LogFile -Message "Microsoft Update service registered successfully." -Status "Info"
            } catch {
                Write-Log -LogFile $LogFile -Message "Failed to register Microsoft Update service: $_" -Status "Error"
                return $null
            }
        } else {
            Write-Log -LogFile $LogFile -Message "Microsoft Update service already registered." -Status "Info"
        }

        # Verify Microsoft Update service connectivity
        Write-Log -LogFile $LogFile -Message "Verifying Microsoft Update service connectivity." -Status "Info"
        try {
            $updateSearcher.ServerSelection = 1
            $updateSearcher.ServiceID = $muServiceId
            $testSearch = $updateSearcher.Search("IsInstalled=0 and Type='Software'")
            Write-Log -LogFile $LogFile -Message "Microsoft Update service is accessible." -Status "Info"
        } catch {
            Write-Log -LogFile $LogFile -Message "Microsoft Update service is not accessible: $_" -Status "Error"
        }

        # Check missing updates
        Write-Log -LogFile $LogFile -Message "Checking for missing Microsoft product updates." -Status "Info"
        $missingUpdates = $null
        $retryCount = 0
        $maxRetries = 3
        while (-not $missingUpdates -and $retryCount -lt $maxRetries) {
            try {
                $missingUpdates = $updateSearcher.Search("IsInstalled=0")
                Write-Log -LogFile $LogFile -Message "Missing updates query succeeded on attempt $($retryCount + 1). Found $($missingUpdates.Updates.Count) updates." -Status "Info"
            } catch {
                $retryCount++
                Write-Log -LogFile $LogFile -Message "Error checking missing updates (attempt $retryCount): $_" -Status "Warning"
                if ($retryCount -eq $maxRetries) {
                    Write-Log -LogFile $LogFile -Message "Max retries reached for missing updates query." -Status "Error"
                    break
                }
                Start-Sleep -Seconds 10
            }
        }

        if ($missingUpdates) {
            Write-Log -LogFile $LogFile -Message "Raw missing updates: $($missingUpdates.Updates | ForEach-Object { $_.Title } | Out-String)" -Status "Info"
            foreach ($update in $missingUpdates.Updates) {
                $kbIds = $update.KBArticleIDs -join ','
                $updateKey = "$($update.Title)_$kbIds_Missing"
                if ($updateKey -notin $processedUpdates) {
                    Write-Log -LogFile $LogFile -Message "Update: $($update.Title) - Status: Missing - KB: $kbIds" -Status "Warning"
                    $results += [PSCustomObject]@{
                        Title = $update.Title
                        Status = "Missing"
                        Details = "KB: $kbIds"
                    }
                    $processedUpdates += $updateKey
                }
            }
        }

        # Check installed updates via COM API
        Write-Log -LogFile $LogFile -Message "Checking for installed Microsoft product updates via COM API." -Status "Info"
        $installedUpdates = $null
        $retryCount = 0
        while (-not $installedUpdates -and $retryCount -lt $maxRetries) {
            try {
                $installedUpdates = $updateSearcher.Search("IsInstalled=1")
                Write-Log -LogFile $LogFile -Message "Installed updates query succeeded on attempt $($retryCount + 1). Found $($installedUpdates.Updates.Count) updates." -Status "Info"
            } catch {
                $retryCount++
                Write-Log -LogFile $LogFile -Message "Error checking installed updates via COM API (attempt $retryCount): $_" -Status "Warning"
                if ($retryCount -eq $maxRetries) {
                    Write-Log -LogFile $LogFile -Message "Max retries reached for installed updates query." -Status "Error"
                    break
                }
                Start-Sleep -Seconds 10
            }
        }

        if ($installedUpdates) {
            Write-Log -LogFile $LogFile -Message "Raw installed updates: $($installedUpdates.Updates | ForEach-Object { $_.Title } | Out-String)" -Status "Info"
            foreach ($update in $installedUpdates.Updates) {
                $kbIds = $update.KBArticleIDs -join ','
                $updateKey = "$($update.Title)_$kbIds_Installed"
                if ($updateKey -notin $processedUpdates) {
                    Write-Log -LogFile $LogFile -Message "Update: $($update.Title) - Status: Installed (COM API) - KB: $kbIds" -Status "Info"
                    $results += [PSCustomObject]@{
                        Title = $update.Title
                        Status = "Installed"
                        Details = "KB: $kbIds"
                    }
                    $processedUpdates += $updateKey
                }
            }
        }

        # Fallback 1: Get-CimInstance
        Write-Log -LogFile $LogFile -Message "Checking installed updates via Get-CimInstance." -Status "Info"
        try {
            $hotfixes = Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction Stop
            Write-Log -LogFile $LogFile -Message "Get-CimInstance returned $($hotfixes.Count) hotfixes." -Status "Info"
            Write-Log -LogFile $LogFile -Message "Raw hotfixes: $($hotfixes | ForEach-Object { "$($_.HotFixID): $($_.Description)" } | Out-String)" -Status "Info"
            foreach ($hotfix in $hotfixes) {
                $kbId = $hotfix.HotFixID
                $title = if ($hotfix.Description) { $hotfix.Description } else { "Microsoft Product Update $kbId" }
                $updateKey = "$title_$kbId_Installed"
                if ($updateKey -notin $processedUpdates) {
                    Write-Log -LogFile $LogFile -Message "Update: $title - Status: Installed (Get-CimInstance) - KB: $kbId" -Status "Info"
                    $results += [PSCustomObject]@{
                        Title = $title
                        Status = "Installed"
                        Details = "KB: $kbId"
                    }
                    $processedUpdates += $updateKey
                }
            }
        } catch {
            Write-Log -LogFile $LogFile -Message "Error checking installed updates via Get-CimInstance: $_" -Status "Error"
        }

        # Fallback 2: Query update history
        Write-Log -LogFile $LogFile -Message "Checking installed updates via update history." -Status "Info"
        try {
            $history = $updateSession.QueryHistory("", 0, 1000)
            Write-Log -LogFile $LogFile -Message "Update history returned $($history.Count) entries." -Status "Info"
            foreach ($entry in $history) {
                if ($entry.ResultCode -eq 2) {
                    $kbId = if ($entry.Title -match "KB\d+") { $matches[0] } else { "Unknown" }
                    $title = $entry.Title
                    $updateKey = "$title_$kbId_Installed"
                    if ($updateKey -notin $processedUpdates) {
                        Write-Log -LogFile $LogFile -Message "Update: $title - Status: Installed (Update History) - KB: $kbId" -Status "Info"
                        $results += [PSCustomObject]@{
                            Title = $title
                            Status = "Installed"
                            Details = "KB: $kbId"
                        }
                        $processedUpdates += $updateKey
                    }
                }
            }
        } catch {
            Write-Log -LogFile $LogFile -Message "Error checking installed updates via update history: $_" -Status "Error"
        }

        # Fallback 3: Registry check
        Write-Log -LogFile $LogFile -Message "Checking installed updates via registry." -Status "Info"
        try {
            $regPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
            )
            $regUpdates = @()
            foreach ($path in $regPaths) {
                $regUpdates += Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName }
            }
            Write-Log -LogFile $LogFile -Message "Registry returned $($regUpdates.Count) entries." -Status "Info"
            Write-Log -LogFile $LogFile -Message "Raw registry updates: $($regUpdates | ForEach-Object { "$($_.DisplayName): $($_.PSChildName)" } | Out-String)" -Status "Info"
            foreach ($update in $regUpdates) {
                if ($update.DisplayName) {
                    $kbId = if ($update.PSChildName -match "KB\d+") { $matches[0] } elseif ($update.DisplayName -match "KB\d+") { $matches[0] } else { $update.PSChildName }
                    $title = $update.DisplayName
                    $updateKey = "$title_$kbId_Installed"
                    if ($updateKey -notin $processedUpdates) {
                        Write-Log -LogFile $LogFile -Message "Update: $title - Status: Installed (Registry) - KB: $kbId" -Status "Info"
                        $results += [PSCustomObject]@{
                            Title = $title
                            Status = "Installed"
                            Details = "KB: $kbId"
                        }
                        $processedUpdates += $updateKey
                    }
                }
            }
        } catch {
            Write-Log -LogFile $LogFile -Message "Error checking installed updates via registry: $_" -Status "Error"
        }

        # Finalize results
        if (-not $results) {
            Write-Log -LogFile $LogFile -Message "No Microsoft product updates found (neither installed nor missing)." -Status "Info"
            $results += [PSCustomObject]@{
                Title = "None"
                Status = "None"
                Details = "No Microsoft product updates detected"
            }
        } else {
            Write-Log -LogFile $LogFile -Message "Found $($results.Count) Microsoft product updates (installed and missing)." -Status "Info"
        }

        return $results
    } catch {
        Write-Log -LogFile $LogFile -Message "General error in Check-MicrosoftProductUpdates: $_" -Status "Error"
        return $null
    }
}