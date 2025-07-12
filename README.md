# Microsoft Security Analyzer
A Windows GUI tool built with PowerShell that helps administrators run system-level security checks like STIG compliance, Group Policy validation, certificate inspection, firewall status, update analysis, and more. Output is organized into detailed logs and reports, auto-generated per machine with timestamps.

## 🔧 Features
STIG Compliance (requires .xml XCCDF file)
Group Policy extraction
Firewall configuration check
User access permissions audit
Certificate inventory
Windows Update assessment (via CAB or online)
Installed software list
System event log scan
Microsoft product updates
Automated report generation

## 📁 Folder Structure (after execution)
```
/Reports/
  └── <HOSTNAME>-<TIMESTAMP>/
        ├── general.log
        ├── stig_compliance.log
        ├── group_policy.log
        ├── firewall.log
        ├── user_access.log
        ├── certificates.log
        ├── windows_updates.log
        ├── installed_software.log
        ├── events.log
        ├── microsoft_product_updates.log
        └── <Generated_Report>.html
```

## 🚀 How to Use

1. Run PowerShell as Administrator
	Execute the Script
```
PowerShell -ExecutionPolicy Bypass .\SecurityAnalysis.ps1
```
2. Select the Checks
	Choose the security checks you want from the GUI. Some options will enable file selectors:
	a. STIG Compliance: requires an `.xml` XCCDF file
	b. Windows Updates: optionally use the offline CAB file (`wsusscn2.cab`)

3. Click "Analyze"
	Logs and reports will be saved in a timestamped directory inside the Reports/ folder.

4. Click "Open Report Folder" to view results.

## 📥 CAB File for Windows Update Scan

To scan for Windows updates offline, download the latest WSUS catalog:

Download `wsusscn2.cab` (https://catalog.s.download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab)

Place the file locally and provide the path via the GUI.

## 🧩 External Scripts Required

The following helper scripts must be present in the same directory:

`Check-STIGCompliance.ps1`
`Check-GroupPolicy.ps1`
`Check-Firewall.ps1`
`Check-UserAccess.ps1`
`Check-Certificates.ps1`
`Check-WindowsUpdates.ps1`
`Check-InstalledSoftware.ps1`
`Check-SystemEvents.ps1`
`Check-MicrosoftProductUpdates.ps1`
`Generate-Report.ps1`
`Write-Log.ps1`

These scripts are modular and called based on user-selected options.

## 🖥 Compatibility

Windows 10/11, Windows Server 2016+
PowerShell 5.1 or later
GUI support (WinForms)

## 👨‍💻 Author
Dipta Roy
Version 1.0 | July 13, 2025