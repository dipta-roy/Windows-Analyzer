function Generate-Report {
    param (
        [string]$OutputDir,
        [array]$StigResults,
        [array]$GpoResults,
        [array]$FirewallResults,
        [array]$UserResults,
        [array]$CertResults,
        [array]$UpdateResults,
        [array]$SoftwareResults,
        [array]$EventResults,
        [array]$MsProductUpdateResults
    )
    try {
        . "$PSScriptRoot\Write-Log.ps1"
        $logFile = Join-Path $OutputDir "general.log"

        # Check if any results exist
        $hasResults = ($StigResults -or $GpoResults -or $FirewallResults -or $UserResults -or $CertResults -or $UpdateResults -or $SoftwareResults -or $EventResults -or $MsProductUpdateResults)
        if (-not $hasResults) {
            $html = @"
<html>
<head>
    <title>Windows Security Analysis Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #2e6da4; }
        p { color: #555; }
    </style>
</head>
<body>
    <h1>Windows Security Analysis Report</h1>
    <p>No checks were selected for analysis.</p>
</body>
</html>
"@
            $htmlFile = Join-Path $OutputDir "SecurityAnalysisReport.html"
            $html | Out-File -FilePath $htmlFile -Encoding UTF8
            Write-Log -LogFile $logFile -Message "Empty report generated at $htmlFile (no checks selected)" -Status "Info"
            return $htmlFile
        }

        $html = @"
<html>
<head>
    <title>Windows Security Analysis Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #2e6da4; }
        h2 { color: #4CAF50; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #4CAF50; color: white; }
        .pass { background-color: #dff0d8; }
        .fail { background-color: #f2dede; }
        .warning { background-color: #fcf8e3; }
        .info { background-color: #d9edf7; }
        .suspicious { background-color: #ffcccc; }
    </style>
</head>
<body>
    <h1>Windows Security Analysis Report</h1>
    <p><strong>Generated on:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    <p><strong>Hostname:</strong> $env:COMPUTERNAME</p>
    <p><strong>Output Directory:</strong> $OutputDir</p>
"@

        if ($StigResults) {
            $html += @"
    <h2>STIG Compliance</h2>
    <table>
        <tr><th>ID</th><th>Title</th><th>Status</th><th>Details</th></tr>
"@
            foreach ($result in $(if ($null -eq $StigResults) { @() } else { $StigResults })) {
                $html += "<tr class='$($result.Status.ToLower())'><td>$($result.ID)</td><td>$($result.Title)</td><td>$($result.Status)</td><td>$($result.Details)</td></tr>"
            }
            $html += "</table>"
        }

        if ($GpoResults) {
            $html += @"
    <h2>Group Policy</h2>
    <table>
        <tr><th>Name</th><th>Status</th><th>Details</th></tr>
"@
            foreach ($result in $(if ($null -eq $GpoResults) { @() } else { $GpoResults })) {
                $html += "<tr class='$($result.Status.ToLower())'><td>$($result.Name)</td><td>$($result.Status)</td><td>$($result.Details)</td></tr>"
            }
            $html += "</table>"
        }

        if ($FirewallResults) {
            $html += @"
    <h2>Firewall Rules</h2>
    <table>
        <tr><th>Name</th><th>Ports</th><th>Status</th><th>Details</th></tr>
"@
            foreach ($result in $(if ($null -eq $FirewallResults) { @() } else { $FirewallResults })) {
                $html += "<tr class='$($result.Status.ToLower())'><td>$($result.Name)</td><td>$($result.Ports)</td><td>$($result.Status)</td><td>$($result.Details)</td></tr>"
            }
            $html += "</table>"
        }

        if ($UserResults) {
            $html += @"
    <h2>User Access</h2>
    <table>
        <tr><th>Name</th><th>Status</th><th>Details</th></tr>
"@
            foreach ($result in $(if ($null -eq $UserResults) { @() } else { $UserResults })) {
                $html += "<tr class='$($result.Status.ToLower())'><td>$($result.Name)</td><td>$($result.Status)</td><td>$($result.Details)</td></tr>"
            }
            $html += "</table>"
        }

        if ($CertResults) {
            $html += @"
    <h2>Certificates</h2>
    <table>
        <tr><th>Subject</th><th>Expires</th><th>Status</th><th>Details</th></tr>
"@
            foreach ($result in $(if ($null -eq $CertResults) { @() } else { $CertResults })) {
                $html += "<tr class='$($result.Status.ToLower())'><td>$($result.Subject)</td><td>$($result.Expires)</td><td>$($result.Status)</td><td>$($result.Details)</td></tr>"
            }
            $html += "</table>"
        }

        if ($UpdateResults) {
            $html += @"
    <h2>Windows Updates</h2>
    <table>
        <tr><th>Title</th><th>Status</th><th>Details</th></tr>
"@
            $missingUpdates = $false
            foreach ($result in $(if ($null -eq $UpdateResults) { @() } else { $UpdateResults })) {
                $rowClass = if ($result.Status -eq "Missing") { "fail" } else { "info" }
                if ($result.Status -eq "Missing") { $missingUpdates = $true }
                $html += "<tr class='$rowClass'><td>$($result.Title)</td><td>$($result.Status)</td><td>$($result.Details)</td></tr>"
            }
            if (-not $missingUpdates -and $UpdateResults) {
                $html += "<tr><td colspan='3'>No missing updates detected</td></tr>"
            }
            $html += "</table>"
        }

        if ($SoftwareResults) {
            $html += @"
    <h2>Installed Software</h2>
    <table>
        <tr><th>Name</th><th>Version</th><th>Status</th><th>Details</th></tr>
"@
            foreach ($result in $(if ($null -eq $SoftwareResults) { @() } else { $SoftwareResults })) {
                $html += "<tr class='$($result.Status.ToLower())'><td>$($result.Name)</td><td>$($result.Version)</td><td>$($result.Status)</td><td>$($result.Details)</td></tr>"
            }
            $html += "</table>"
        }

        if ($EventResults) {
            $html += @"
    <h2>System Events</h2>
    <table>
        <tr><th>Event ID</th><th>Time</th><th>Description</th><th>Status</th><th>Details</th></tr>
"@
            foreach ($result in $(if ($null -eq $EventResults) { @() } else { $EventResults })) {
                $html += "<tr class='$($result.Status.ToLower())'><td>$($result.EventID)</td><td>$($result.Time)</td><td>$($result.Description)</td><td>$($result.Status)</td><td>$($result.Details)</td></tr>"
            }
            $html += "</table>"
        }

        if ($MsProductUpdateResults) {
            $html += @"
    <h2>Microsoft Product Updates</h2>
    <table>
        <tr><th>Title</th><th>Status</th><th>Details</th></tr>
"@
            $missingUpdates = $false
            foreach ($result in $(if ($null -eq $MsProductUpdateResults) { @() } else { $MsProductUpdateResults })) {
                $rowClass = if ($result.Status -eq "Missing") { "fail" } else { "info" }
                if ($result.Status -eq "Missing") { $missingUpdates = $true }
                $html += "<tr class='$rowClass'><td>$($result.Title)</td><td>$($result.Status)</td><td>$($result.Details)</td></tr>"
            }
            if (-not $missingUpdates -and $MsProductUpdateResults) {
                $html += "<tr><td colspan='3'>No missing Microsoft product updates detected</td></tr>"
            }
            $html += "</table>"
        }

        $html += @"
</body>
</html>
"@
        $htmlFile = Join-Path $OutputDir "SecurityAnalysisReport.html"
        $html | Out-File -FilePath $htmlFile -Encoding UTF8
        Write-Log -LogFile $logFile -Message "Report generated at $htmlFile" -Status "Info"
        return $htmlFile
    } catch {
        Write-Log -LogFile $logFile -Message "Error generating report: $_" -Status "Error"
        return $null
    }
}