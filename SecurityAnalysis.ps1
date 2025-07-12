###################################################
#
#   Title   : Microsoft Security Analyzer
#   Author  : Dipta Roy
#   Version : 1.0
#   Date    : 13-07-2025
#   Usage   : Run as Administrator
#             PowerShell -ExecutionPolicy Bypass .\SecurityAnalysis.ps1
#   Note    : Download latest wsusscn2.cab from:
#             https://catalog.s.download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab
#
###################################################

#Requires -RunAsAdministrator

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create Windows Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows Security Analysis Tool"
$form.Size = New-Object System.Drawing.Size(600,600)
$form.StartPosition = "CenterScreen"

# STIG File Path
$labelStig = New-Object System.Windows.Forms.Label
$labelStig.Location = New-Object System.Drawing.Point(10,20)
$labelStig.Size = New-Object System.Drawing.Size(100,20)
$labelStig.Text = "STIG File Path:"
$form.Controls.Add($labelStig)

$textBoxStig = New-Object System.Windows.Forms.TextBox
$textBoxStig.Location = New-Object System.Drawing.Point(120,20)
$textBoxStig.Size = New-Object System.Drawing.Size(350,20)
$textBoxStig.Enabled = $false
$form.Controls.Add($textBoxStig)

$buttonBrowseStig = New-Object System.Windows.Forms.Button
$buttonBrowseStig.Location = New-Object System.Drawing.Point(480,20)
$buttonBrowseStig.Size = New-Object System.Drawing.Size(75,23)
$buttonBrowseStig.Text = "Browse"
$buttonBrowseStig.Enabled = $false
$buttonBrowseStig.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "XCCDF Files (*.xml)|*.xml"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $textBoxStig.Text = $openFileDialog.FileName
    }
})
$form.Controls.Add($buttonBrowseStig)

# CAB File Path
$labelCab = New-Object System.Windows.Forms.Label
$labelCab.Location = New-Object System.Drawing.Point(10,50)
$labelCab.Size = New-Object System.Drawing.Size(100,20)
$labelCab.Text = "CAB File Path:"
$form.Controls.Add($labelCab)

$textBoxCab = New-Object System.Windows.Forms.TextBox
$textBoxCab.Location = New-Object System.Drawing.Point(120,50)
$textBoxCab.Size = New-Object System.Drawing.Size(350,20)
$textBoxCab.Enabled = $false
$form.Controls.Add($textBoxCab)

$buttonBrowseCab = New-Object System.Windows.Forms.Button
$buttonBrowseCab.Location = New-Object System.Drawing.Point(480,50)
$buttonBrowseCab.Size = New-Object System.Drawing.Size(75,23)
$buttonBrowseCab.Text = "Browse"
$buttonBrowseCab.Enabled = $false
$buttonBrowseCab.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CAB Files (*.cab)|*.cab"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $textBoxCab.Text = $openFileDialog.FileName
    }
})
$form.Controls.Add($buttonBrowseCab)

# Checkboxes for selecting checks
$checkStig = New-Object System.Windows.Forms.CheckBox
$checkStig.Location = New-Object System.Drawing.Point(10,80)
$checkStig.Size = New-Object System.Drawing.Size(200,20)
$checkStig.Text = "STIG Compliance"
$checkStig.Add_CheckedChanged({
    $textBoxStig.Enabled = $checkStig.Checked
    $buttonBrowseStig.Enabled = $checkStig.Checked
})
$form.Controls.Add($checkStig)

$checkGpo = New-Object System.Windows.Forms.CheckBox
$checkGpo.Location = New-Object System.Drawing.Point(10,100)
$checkGpo.Size = New-Object System.Drawing.Size(200,20)
$checkGpo.Text = "Group Policy"
$form.Controls.Add($checkGpo)

$checkFirewall = New-Object System.Windows.Forms.CheckBox
$checkFirewall.Location = New-Object System.Drawing.Point(10,120)
$checkFirewall.Size = New-Object System.Drawing.Size(200,20)
$checkFirewall.Text = "Firewall"
$form.Controls.Add($checkFirewall)

$checkUser = New-Object System.Windows.Forms.CheckBox
$checkUser.Location = New-Object System.Drawing.Point(10,140)
$checkUser.Size = New-Object System.Drawing.Size(200,20)
$checkUser.Text = "User Access"
$form.Controls.Add($checkUser)

$checkCerts = New-Object System.Windows.Forms.CheckBox
$checkCerts.Location = New-Object System.Drawing.Point(10,160)
$checkCerts.Size = New-Object System.Drawing.Size(200,20)
$checkCerts.Text = "Certificates"
$form.Controls.Add($checkCerts)

$checkUpdates = New-Object System.Windows.Forms.CheckBox
$checkUpdates.Location = New-Object System.Drawing.Point(10,180)
$checkUpdates.Size = New-Object System.Drawing.Size(200,20)
$checkUpdates.Text = "Windows Updates"
$checkUpdates.Add_CheckedChanged({
    $textBoxCab.Enabled = $checkUpdates.Checked
    $buttonBrowseCab.Enabled = $checkUpdates.Checked
})
$form.Controls.Add($checkUpdates)

$checkSoftware = New-Object System.Windows.Forms.CheckBox
$checkSoftware.Location = New-Object System.Drawing.Point(10,200)
$checkSoftware.Size = New-Object System.Drawing.Size(200,20)
$checkSoftware.Text = "Installed Software"
$form.Controls.Add($checkSoftware)

$checkEvents = New-Object System.Windows.Forms.CheckBox
$checkEvents.Location = New-Object System.Drawing.Point(10,220)
$checkEvents.Size = New-Object System.Drawing.Size(200,20)
$checkEvents.Text = "System Events"
$form.Controls.Add($checkEvents)

$checkMsProductUpdates = New-Object System.Windows.Forms.CheckBox
$checkMsProductUpdates.Location = New-Object System.Drawing.Point(10,240)
$checkMsProductUpdates.Size = New-Object System.Drawing.Size(200,20)
$checkMsProductUpdates.Text = "Microsoft Product Updates"
$form.Controls.Add($checkMsProductUpdates)

# Status Label
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(10,270)
$labelStatus.Size = New-Object System.Drawing.Size(550,80)
$labelStatus.Text = "Ready to start analysis. Select checks to perform."
$form.Controls.Add($labelStatus)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,350)
$progressBar.Size = New-Object System.Drawing.Size(550,20)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Analyze Button
$buttonAnalyze = New-Object System.Windows.Forms.Button
$buttonAnalyze.Location = New-Object System.Drawing.Point(150,380)
$buttonAnalyze.Size = New-Object System.Drawing.Size(100,30)
$buttonAnalyze.Text = "Analyze"
$buttonAnalyze.Add_Click({
    $labelStatus.Text = "Starting analysis..."
    $progressBar.Visible = $true
    $progressBar.Value = 0
    $form.Refresh()

    # Count selected checks (including report generation as a step)
    $selectedChecks = @($checkStig.Checked, $checkGpo.Checked, $checkFirewall.Checked, $checkUser.Checked, $checkCerts.Checked, $checkUpdates.Checked, $checkSoftware.Checked, $checkEvents.Checked, $checkMsProductUpdates.Checked).Where({$_ -eq $true}).Count
    $totalSteps = $selectedChecks + 1 # Include report generation
    $stepIncrement = if ($totalSteps -gt 0) { 100 / $totalSteps } else { 100 }
    
    # Validate inputs
    $stigFile = $textBoxStig.Text
    $cabFile = $textBoxCab.Text

    if ($checkStig.Checked -and -not (Test-Path $stigFile)) {
        $labelStatus.Text = "Error: STIG file not found."
        $progressBar.Visible = $false
        return
    }
    if ($checkUpdates.Checked -and -not (Test-Path $cabFile)) {
        $labelStatus.Text = "Warning: CAB file not found. Checking online updates."
        $form.Refresh()
    }

    # Set output directory to <ScriptDir>\Reports\<Hostname>-<Timestamp>
    $scriptDir = $PSScriptRoot
    $reportsDir = Join-Path $scriptDir "Reports"
    $timestamp = Get-Date -Format "yyyyMMddTHHmmss"
    $hostname = $env:COMPUTERNAME
    $outputDir = Join-Path $reportsDir "$hostname-$timestamp"
    New-Item -Path $outputDir -ItemType Directory -Force
    . "$PSScriptRoot\Write-Log.ps1"
    Write-Log -LogFile (Join-Path $outputDir "general.log") -Message "Created output directory: $outputDir" -Status "Info"

    # Initialize log files
    $generalLog = Join-Path $outputDir "general.log"
    $stigLog = Join-Path $outputDir "stig_compliance.log"
    $gpoLog = Join-Path $outputDir "group_policy.log"
    $firewallLog = Join-Path $outputDir "firewall.log"
    $userLog = Join-Path $outputDir "user_access.log"
    $certLog = Join-Path $outputDir "certificates.log"
    $updateLog = Join-Path $outputDir "windows_updates.log"
    $softwareLog = Join-Path $outputDir "installed_software.log"
    $eventLog = Join-Path $outputDir "events.log"
    $msProductUpdateLog = Join-Path $outputDir "microsoft_product_updates.log"

    # Initialize results
    $stigResults = @()
    $gpoResults = @()
    $firewallResults = @()
    $userResults = @()
    $certResults = @()
    $updateResults = @()
    $softwareResults = @()
    $eventResults = @()
    $msProductUpdateResults = @()

    # Perform selected analyses
    . "$PSScriptRoot\Write-Log.ps1"

    if ($checkStig.Checked) {
        try {
            $labelStatus.Text = "Analyzing STIG compliance..."
            $form.Refresh()
            if (Test-Path "$PSScriptRoot\Check-STIGCompliance.ps1") {
                . "$PSScriptRoot\Check-STIGCompliance.ps1"
                $stigResults = Check-STIGCompliance -StigFile $stigFile -LogFile $stigLog
                Write-Log -LogFile $generalLog -Message "STIG compliance analysis completed" -Status "Info"
            } else {
                Write-Log -LogFile $generalLog -Message "Check-STIGCompliance.ps1 not found in $PSScriptRoot" -Status "Error"
            }
            $labelStatus.Text = "STIG compliance analysis complete."
            $progressBar.Value = [math]::Min($progressBar.Value + $stepIncrement, 100)
            $form.Refresh()
        } catch {
            Write-Log -LogFile $generalLog -Message "Error in STIG compliance analysis: $_" -Status "Error"
        }
    }

    if ($checkGpo.Checked) {
        try {
            $labelStatus.Text = "Analyzing Group Policy..."
            $form.Refresh()
            if (Test-Path "$PSScriptRoot\Check-GroupPolicy.ps1") {
                . "$PSScriptRoot\Check-GroupPolicy.ps1"
                $gpoResults = Check-GroupPolicy -LogFile $gpoLog
                Write-Log -LogFile $generalLog -Message "Group Policy analysis completed" -Status "Info"
            } else {
                Write-Log -LogFile $generalLog -Message "Check-GroupPolicy.ps1 not found in $PSScriptRoot" -Status "Error"
            }
            $labelStatus.Text = "Group Policy analysis complete. Note: GroupPolicy module may be unavailable (using gpresult)."
            $progressBar.Value = [math]::Min($progressBar.Value + $stepIncrement, 100)
            $form.Refresh()
        } catch {
            Write-Log -LogFile $generalLog -Message "Error in Group Policy analysis: $_" -Status "Error"
        }
    }

    if ($checkFirewall.Checked) {
        try {
            $labelStatus.Text = "Analyzing Firewall..."
            $form.Refresh()
            if (Test-Path "$PSScriptRoot\Check-Firewall.ps1") {
                . "$PSScriptRoot\Check-Firewall.ps1"
                $firewallResults = Check-Firewall -LogFile $firewallLog
                Write-Log -LogFile $generalLog -Message "Firewall analysis completed" -Status "Info"
            } else {
                Write-Log -LogFile $generalLog -Message "Check-Firewall.ps1 not found in $PSScriptRoot" -Status "Error"
            }
            $labelStatus.Text = "Firewall analysis complete."
            $progressBar.Value = [math]::Min($progressBar.Value + $stepIncrement, 100)
            $form.Refresh()
        } catch {
            Write-Log -LogFile $generalLog -Message "Error in Firewall analysis: $_" -Status "Error"
        }
    }

    if ($checkUser.Checked) {
        try {
            $labelStatus.Text = "Analyzing User Access..."
            $form.Refresh()
            if (Test-Path "$PSScriptRoot\Check-UserAccess.ps1") {
                . "$PSScriptRoot\Check-UserAccess.ps1"
                $userResults = Check-UserAccess -LogFile $userLog
                Write-Log -LogFile $generalLog -Message "User access analysis completed" -Status "Info"
            } else {
                Write-Log -LogFile $generalLog -Message "Check-UserAccess.ps1 not found in $PSScriptRoot" -Status "Error"
            }
            $labelStatus.Text = "User access analysis complete."
            $progressBar.Value = [math]::Min($progressBar.Value + $stepIncrement, 100)
            $form.Refresh()
        } catch {
            Write-Log -LogFile $generalLog -Message "Error in User Access analysis: $_" -Status "Error"
        }
    }

    if ($checkCerts.Checked) {
        try {
            $labelStatus.Text = "Analyzing Certificates..."
            $form.Refresh()
            if (Test-Path "$PSScriptRoot\Check-Certificates.ps1") {
                . "$PSScriptRoot\Check-Certificates.ps1"
                $certResults = Check-Certificates -LogFile $certLog
                Write-Log -LogFile $generalLog -Message "Certificates analysis completed" -Status "Info"
            } else {
                Write-Log -LogFile $generalLog -Message "Check-Certificates.ps1 not found in $PSScriptRoot" -Status "Error"
            }
            $labelStatus.Text = "Certificates analysis complete."
            $progressBar.Value = [math]::Min($progressBar.Value + $stepIncrement, 100)
            $form.Refresh()
        } catch {
            Write-Log -LogFile $generalLog -Message "Error in Certificates analysis: $_" -Status "Error"
        }
    }

    if ($checkUpdates.Checked) {
        try {
            $labelStatus.Text = "Analyzing Windows Updates..."
            $form.Refresh()
            if (Test-Path "$PSScriptRoot\Check-WindowsUpdates.ps1") {
                . "$PSScriptRoot\Check-WindowsUpdates.ps1"
                $updateResults = Check-WindowsUpdates -CabFile $cabFile -LogFile $updateLog
                Write-Log -LogFile $generalLog -Message "Windows Updates analysis completed" -Status "Info"
            } else {
                Write-Log -LogFile $generalLog -Message "Check-WindowsUpdates.ps1 not found in $PSScriptRoot" -Status "Error"
            }
            $labelStatus.Text = "Windows Updates analysis complete."
            $progressBar.Value = [math]::Min($progressBar.Value + $stepIncrement, 100)
            $form.Refresh()
        } catch {
            Write-Log -LogFile $generalLog -Message "Error in Windows Updates analysis: $_" -Status "Error"
        }
    }

    if ($checkSoftware.Checked) {
        try {
            $labelStatus.Text = "Analyzing Installed Software..."
            $form.Refresh()
            if (Test-Path "$PSScriptRoot\Check-InstalledSoftware.ps1") {
                . "$PSScriptRoot\Check-InstalledSoftware.ps1"
                $softwareResults = Check-InstalledSoftware -LogFile $softwareLog
                Write-Log -LogFile $generalLog -Message "Installed Software analysis completed" -Status "Info"
            } else {
                Write-Log -LogFile $generalLog -Message "Check-InstalledSoftware.ps1 not found in $PSScriptRoot" -Status "Error"
            }
            $labelStatus.Text = "Installed Software analysis complete."
            $progressBar.Value = [math]::Min($progressBar.Value + $stepIncrement, 100)
            $form.Refresh()
        } catch {
            Write-Log -LogFile $generalLog -Message "Error in Installed Software analysis: $_" -Status "Error"
        }
    }

    if ($checkEvents.Checked) {
        try {
            $labelStatus.Text = "Analyzing System Events..."
            $form.Refresh()
            if (Test-Path "$PSScriptRoot\Check-SystemEvents.ps1") {
                . "$PSScriptRoot\Check-SystemEvents.ps1"
                $eventResults = Check-SystemEvents -LogFile $eventLog
                Write-Log -LogFile $generalLog -Message "System Events analysis completed" -Status "Info"
            } else {
                Write-Log -LogFile $generalLog -Message "Check-SystemEvents.ps1 not found in $PSScriptRoot" -Status "Error"
            }
            $labelStatus.Text = "System Events analysis complete."
            $progressBar.Value = [math]::Min($progressBar.Value + $stepIncrement, 100)
            $form.Refresh()
        } catch {
            Write-Log -LogFile $generalLog -Message "Error in System Events analysis: $_" -Status "Error"
        }
    }

    if ($checkMsProductUpdates.Checked) {
        try {
            $labelStatus.Text = "Analyzing Microsoft Product Updates..."
            $form.Refresh()
            if (Test-Path "$PSScriptRoot\Check-MicrosoftProductUpdates.ps1") {
                . "$PSScriptRoot\Check-MicrosoftProductUpdates.ps1"
                $msProductUpdateResults = Check-MicrosoftProductUpdates -LogFile $msProductUpdateLog
                Write-Log -LogFile $generalLog -Message "Microsoft Product Updates analysis completed" -Status "Info"
            } else {
                Write-Log -LogFile $generalLog -Message "Check-MicrosoftProductUpdates.ps1 not found in $PSScriptRoot" -Status "Error"
            }
            $labelStatus.Text = "Microsoft Product Updates analysis complete."
            $progressBar.Value = [math]::Min($progressBar.Value + $stepIncrement, 100)
            $form.Refresh()
        } catch {
            Write-Log -LogFile $generalLog -Message "Error in Microsoft Product Updates analysis: $_" -Status "Error"
        }
    }

    # Generate Report
    try {
        $labelStatus.Text = "Generating report..."
        $form.Refresh()
        if (Test-Path "$PSScriptRoot\Generate-Report.ps1") {
            . "$PSScriptRoot\Generate-Report.ps1"
            $reportPath = Generate-Report -OutputDir $outputDir -StigResults $stigResults -GpoResults $gpoResults -FirewallResults $firewallResults -UserResults $userResults -CertResults $certResults -UpdateResults $updateResults -SoftwareResults $softwareResults -EventResults $eventResults -MsProductUpdateResults $msProductUpdateResults
            Write-Log -LogFile $generalLog -Message "Report generation completed" -Status "Info"
            $labelStatus.Text = "Analysis complete. Report saved to $reportPath"
            $progressBar.Value = 100
            $form.Refresh()
        } else {
            Write-Log -LogFile $generalLog -Message "Generate-Report.ps1 not found in $PSScriptRoot" -Status "Error"
            $labelStatus.Text = "Error: Report generation script not found. Check logs in $outputDir"
        }
    } catch {
        Write-Log -LogFile $generalLog -Message "Error in report generation: $_" -Status "Error"
        $labelStatus.Text = "Error during report generation. Check logs in $outputDir"
    } finally {
        $progressBar.Value = 0
        $progressBar.Visible = $false
        $form.Refresh()
    }
})
$form.Controls.Add($buttonAnalyze)

# Open Report Folder Button
$buttonOpenFolder = New-Object System.Windows.Forms.Button
$buttonOpenFolder.Location = New-Object System.Drawing.Point(260,380)
$buttonOpenFolder.Size = New-Object System.Drawing.Size(120,30)
$buttonOpenFolder.Text = "Open Report Folder"
$buttonOpenFolder.Add_Click({
    $scriptDir = $PSScriptRoot
    $reportsDir = Join-Path $scriptDir "Reports"
    $timestamp = Get-Date -Format "yyyyMMddTHHmmss"
    $hostname = $env:COMPUTERNAME
    $outputDir = Join-Path $reportsDir "$hostname-$timestamp"
    if (Test-Path $outputDir) {
        Invoke-Item $outputDir
        $labelStatus.Text = "Opened report folder: $outputDir"
    } else {
        $labelStatus.Text = "Error: Report folder does not exist. Run analysis to create it."
    }
    $form.Refresh()
})
$form.Controls.Add($buttonOpenFolder)

# Reset Button
$buttonReset = New-Object System.Windows.Forms.Button
$buttonReset.Location = New-Object System.Drawing.Point(390,380)
$buttonReset.Size = New-Object System.Drawing.Size(100,30)
$buttonReset.Text = "Reset"
$buttonReset.Add_Click({
    $textBoxStig.Text = ""
    $textBoxCab.Text = ""
    $checkStig.Checked = $false
    $checkGpo.Checked = $false
    $checkFirewall.Checked = $false
    $checkUser.Checked = $false
    $checkCerts.Checked = $false
    $checkUpdates.Checked = $false
    $checkSoftware.Checked = $false
    $checkEvents.Checked = $false
    $checkMsProductUpdates.Checked = $false
    $textBoxStig.Enabled = $false
    $buttonBrowseStig.Enabled = $false
    $textBoxCab.Enabled = $false
    $buttonBrowseCab.Enabled = $false
    $labelStatus.Text = "Ready to start analysis. Select checks to perform."
    $progressBar.Value = 0
    $progressBar.Visible = $false
    $form.Refresh()
})
$form.Controls.Add($buttonReset)

# Footer Label
$labelFooter = New-Object System.Windows.Forms.Label
$labelFooter.Location = New-Object System.Drawing.Point(10,560)
$labelFooter.Size = New-Object System.Drawing.Size(550,20)
$labelFooter.Text = "Developed by Dipta"
$labelFooter.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$labelFooter.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($labelFooter)

# Show Form
$form.ShowDialog()