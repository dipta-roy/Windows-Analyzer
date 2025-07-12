#Requires -RunAsAdministrator

function Check-WindowsUpdates {
    param (
        [string]$CabFile,
        [string]$LogFile
    )
    try {
        . "$PSScriptRoot\Write-Log.ps1"
        $results = @()

        # Check Windows Update service status
        $wuService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
        if ($wuService) {
            if ($wuService.Status -ne 'Running') {
                Write-Log -LogFile $LogFile -Message "Windows Update service is not running. Attempting to start." -Status "Warning"
                try {
                    Start-Service -Name wuauserv -ErrorAction Stop
                    Write-Log -LogFile $LogFile -Message "Windows Update service started." -Status "Info"
                } catch {
                    Write-Log -LogFile $LogFile -Message "Failed to start Windows Update service: $_" -Status "Error"
                }
            } else {
                Write-Log -LogFile $LogFile -Message "Windows Update service is running." -Status "Info"
            }
        } else {
            Write-Log -LogFile $LogFile -Message "Windows Update service not found." -Status "Warning"
        }

        if ($CabFile -and (Test-Path $CabFile)) {
            Write-Log -LogFile $LogFile -Message "CAB file provided but offline scanning is not supported. Using online Windows Update scanning." -Status "Warning"
        }

        # Initialize COM object
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateUpdateSearcher()
        $updateSearcher.ServerSelection = 1 # ssWindowsUpdate

        # Check missing updates
        Write-Log -LogFile $LogFile -Message "Checking for missing updates using online Windows Update." -Status "Info"
        try {
            $missingUpdates = $updateSearcher.Search("IsInstalled=0")
            foreach ($update in $missingUpdates.Updates) {
                Write-Log -LogFile $LogFile -Message "Update: $($update.Title) - Status: Missing" -Status "Warning"
                $results += [PSCustomObject]@{
                    Title=$update.Title
                    Status="Missing"
                    Details="KB: $($update.KBArticleIDs -join ',')"
                }
            }
        } catch {
            Write-Log -LogFile $LogFile -Message "Error checking missing updates: $_" -Status "Error"
        }

        # Check installed updates via COM API
        Write-Log -LogFile $LogFile -Message "Checking for installed updates using online Windows Update." -Status "Info"
        $installedKBs = @()
        try {
            $installedUpdates = $updateSearcher.Search("IsInstalled=1")
            foreach ($update in $installedUpdates.Updates) {
                Write-Log -LogFile $LogFile -Message "Update: $($update.Title) - Status: Installed (COM API)" -Status "Info"
                $kbIds = $update.KBArticleIDs -join ','
                $results += [PSCustomObject]@{
                    Title=$update.Title
                    Status="Installed"
                    Details="KB: $kbIds"
                }
                $installedKBs += $update.KBArticleIDs
            }
        } catch {
            Write-Log -LogFile $LogFile -Message "Error checking installed updates via COM API: $_" -Status "Error"
        }

        # Fallback to Get-Hotfix if no installed updates found via COM API
        if (-not $installedKBs) {
            Write-Log -LogFile $LogFile -Message "No installed updates found via COM API. Falling back to Get-Hotfix." -Status "Warning"
            try {
                $hotfixes = Get-Hotfix -ErrorAction Stop
                foreach ($hotfix in $hotfixes) {
                    $kbId = $hotfix.HotFixID
                    if ($kbId -notin $installedKBs) {
                        Write-Log -LogFile $LogFile -Message "Update: KB$kbId - Status: Installed (Get-Hotfix)" -Status "Info"
                        $results += [PSCustomObject]@{
                            Title="Windows Update KB$kbId"
                            Status="Installed"
                            Details="KB: $kbId"
                        }
                        $installedKBs += $kbId
                    }
                }
            } catch {
                Write-Log -LogFile $LogFile -Message "Error checking installed updates via Get-Hotfix: $_" -Status "Error"
            }
        }

        if (-not $results) {
            Write-Log -LogFile $LogFile -Message "No updates found (neither installed nor missing)." -Status "Info"
            $results += [PSCustomObject]@{
                Title="None"
                Status="None"
                Details="No updates detected"
            }
        } else {
            Write-Log -LogFile $LogFile -Message "Found $($results.Count) updates (installed and missing)." -Status "Info"
        }

        return $results
    } catch {
        Write-Log -LogFile $LogFile -Message "General error in Check-WindowsUpdates: $_" -Status "Error"
        return $null
    }
}