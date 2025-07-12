#Requires -RunAsAdministrator
Add-Type -AssemblyName System.Xml

function Check-STIGCompliance {
    param (
        [string]$StigFile,
        [string]$LogFile
    )
    try {
        . "$PSScriptRoot\Write-Log.ps1"
        if (-not (Test-Path $StigFile)) {
            Write-Log -LogFile $LogFile -Message "STIG file not found: $StigFile" -Status "Error"
            return $null
        }
        $xml = [xml](Get-Content $StigFile -ErrorAction Stop)
        $benchmark = $xml.Benchmark
        if (-not $benchmark) {
            Write-Log -LogFile $LogFile -Message "Invalid XCCDF file: No Benchmark element found" -Status "Error"
            return $null
        }
        $rules = $benchmark.Group.Rule
        if (-not $rules) {
            Write-Log -LogFile $LogFile -Message "No rules found in XCCDF file" -Status "Warning"
            return @([PSCustomObject]@{ID="None"; Title="No Rules"; Status="Not Processed"; Details="No STIG rules found"})
        }
        $results = @()
        foreach ($rule in $rules) {
            $id = $rule.id
            $title = $rule.title
            if ($id -match "V-\d+") {
                # Try to get MinimumPasswordLength from registry
                $minLength = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Name MinimumPasswordLength -ErrorAction SilentlyContinue).MinimumPasswordLength
                if ($null -eq $minLength) {
                    # Fallback to local security policy using secedit
                    Write-Log -LogFile $LogFile -Message "Registry key for MinimumPasswordLength not found. Checking local security policy." -Status "Info"
                    $tempFile = [System.IO.Path]::GetTempFileName()
                    secedit /export /cfg $tempFile /quiet
                    $secPol = Get-Content $tempFile
                    Remove-Item $tempFile
                    $minLengthLine = $secPol | Where-Object { $_ -match "^MinimumPasswordLength\s*=" }
                    if ($minLengthLine -match "^MinimumPasswordLength\s*=\s*(\d+)") {
                        $minLength = [int]$matches[1]
                    } else {
                        $minLength = $null
                    }
                }
                $displayLength = if ($null -eq $minLength) { "Unknown" } else { $minLength }
                if ($null -ne $minLength -and $minLength -ge 14) {
                    Write-Log -LogFile $LogFile -Message "STIG $id ($title): Password length is $minLength" -Status "Pass"
                    $results += [PSCustomObject]@{ID=$id; Title=$title; Status="Pass"; Details="Password length is $minLength"}
                } else {
                    Write-Log -LogFile $LogFile -Message "STIG $id ($title): Password length is $displayLength (Required: 14)" -Status "Fail"
                    $results += [PSCustomObject]@{ID=$id; Title=$title; Status="Fail"; Details="Password length is $displayLength (Required: 14)"}
                }
            } else {
                Write-Log -LogFile $LogFile -Message "STIG $id ($title): Not checked (no specific rule implemented)" -Status "Info"
                $results += [PSCustomObject]@{ID=$id; Title=$title; Status="Not Checked"; Details="No check implemented for this rule"}
            }
        }
        return $results
    } catch {
        Write-Log -LogFile $LogFile -Message "Error parsing STIG file: $_" -Status "Error"
        return $null
    }
}