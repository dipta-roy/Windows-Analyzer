#Requires -RunAsAdministrator

function Check-GroupPolicy {
    param (
        [string]$LogFile
    )
    try {
        . "$PSScriptRoot\Write-Log.ps1"
        $results = @()
        if (Get-Module -ListAvailable -Name GroupPolicy) {
            Import-Module GroupPolicy -ErrorAction Stop
            $gpoReport = Get-GPResultantSetOfPolicy -ReportType Xml -ErrorAction Stop
            [xml]$gpoXml = $gpoReport
            $policies = $gpoXml.RSOP.ComputerResults.GPO
            if ($policies) {
                foreach ($policy in $policies) {
                    Write-Log -LogFile $LogFile -Message "GPO Applied: $($policy.Name)" -Status "Info"
                    $results += [PSCustomObject]@{Name=$policy.Name; Status="Applied"; Details="GPO is active"}
                }
            } else {
                Write-Log -LogFile $LogFile -Message "No Group Policies applied to this computer." -Status "Info"
                $results += [PSCustomObject]@{Name="None"; Status="Not Applied"; Details="No GPOs found"}
            }
        } else {
            Write-Log -LogFile $LogFile -Message "GroupPolicy module not available. Falling back to gpresult. To enable full GPO analysis, install RSAT: https://www.microsoft.com/en-us/download/details.aspx?id=45520 or enable via Optional Features." -Status "Warning"
            $gpresult = gpresult /r /scope computer 2>&1
            $gpoSection = $false
            $gpoNames = @()
            foreach ($line in ($gpresult -split "`n")) {
                if ($line -match "Applied Group Policy Objects") {
                    $gpoSection = $true
                    continue
                }
                if ($gpoSection -and $line -match "^\s+(.+)$") {
                    $gpoName = $matches[1].Trim()
                    if ($gpoName -and $gpoName -notmatch "^-+$") {
                        $gpoNames += $gpoName
                        Write-Log -LogFile $LogFile -Message "GPO Applied: $gpoName" -Status "Info"
                        $results += [PSCustomObject]@{Name=$gpoName; Status="Applied"; Details="GPO is active (via gpresult)"}
                    }
                }
                if ($gpoSection -and $line -match "^\s*$") {
                    $gpoSection = $false
                }
            }
            if (-not $gpoNames) {
                Write-Log -LogFile $LogFile -Message "No Group Policies applied to this computer (via gpresult)." -Status "Info"
                $results += [PSCustomObject]@{Name="None"; Status="Not Applied"; Details="No GPOs found (via gpresult)"}
            }
        }
        return $results
    } catch {
        Write-Log -LogFile $LogFile -Message "Error retrieving GPO: $_" -Status "Error"
        return $null
    }
}