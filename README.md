<p align="center">
  <h1 align="center">🎯 VISIO ENTERPRISE AUDIT SUITE</h1>
  <p align="center">
    Comprehensive Domain-Wide Visio Installation & Usage Tracking
  </p>
  <p align="center">
    🔍 Audit • 📊 Analytics • 💰 Cost Control • 🛡️ Compliance
  </p>
  <p align="center">
    <strong>PowerShell-based enterprise auditing for Active Directory environments</strong>
  </p>
</p>

---

<p align="center">
  🚀 Scan 1000+ machines &nbsp;|&nbsp;
  📈 HTML & CSV Dashboards &nbsp;|&nbsp;
  ⚙️ Parallel Processing &nbsp;|&nbsp;
  🧠 Usage Intelligence
</p>

---
<p align="center">
  <img alt="PowerShell" src="https://img.shields.io/badge/PowerShell-5.1%2B-blue">
  <img alt="Platform" src="https://img.shields.io/badge/Platform-Windows%20Domain-lightgrey">
  <img alt="Scope" src="https://img.shields.io/badge/Scope-Enterprise-green">
  <img alt="Reports" src="https://img.shields.io/badge/Reports-HTML%20%7C%20CSV-orange">
  <img alt="Automation" src="https://img.shields.io/badge/Automation-Scheduled%20Tasks-purple">
</p>

---

# 🎯 Visio Enterprise Audit Suite
## Comprehensive Domain-Wide Visio Installation & Usage Tracking

---

## Scripts Overview

## Documentation Index

- Start here: `documentation/USER_GUIDE.md`
- Audit implementation details: `documentation/VISIO_AUDIT_GUIDE.md`
- Troubleshooting: `documentation/TROUBLESHOOTING.md`
- Deployment: `documentation/DEPLOYMENT.md`
- Download/package metadata: `documentation/DOWNLOAD_GUIDE.md`, `documentation/PACKAGE_INFO.md`

This suite contains PowerShell scripts for auditing Visio installations and detecting Office versions across your enterprise Active Directory environment.

### Office-Version-Detector.ps1

**Purpose:** Detects Microsoft Office installations and identifies if Office 365, Office 2019, or Office 2016 is installed.

**Description:**
This script performs version detection for Microsoft Office installations by checking registry keys for Windows Installer (MSI) deployments. It specifically identifies Office 365, Office 2019, and Office 2016 installations while rejecting older versions (Office 2013, 2010, etc.).

**Features:**
- Registry-based detection for MSI installations
- Supports both 32-bit and 64-bit system detection
- Detailed logging to console and file
- Comprehensive error handling with strict mode option
- Exit codes: 0 (success - supported version), 1 (unsupported version or error)

**Parameters:**
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-LogFilePath` | string | `.\Office-Version-Detection.log` | Path for the log file |
| `-StrictErrorHandling` | switch | $false | Enables strict error handling mode - terminates on non-critical errors |
| `-VerboseLogging` | switch | $false | Enables verbose logging output to console |

**Usage Examples:**
```powershell
# Basic detection
.\Office-Version-Detector.ps1

# With custom log file and verbose logging
.\Office-Version-Detector.ps1 -LogFilePath "C:\Logs\OfficeDetection.log" -VerboseLogging

# With strict error handling for production environments
.\Office-Version-Detector.ps1 -StrictErrorHandling
```

- `0` - Success: Supported Office version detected (Office 365, Office 2019, or Office 2016)
- `1` - Error: Unsupported version detected or detection failed

---

### Visio-Enterprise-Audit.ps1

**Purpose:** Enterprise Visio Installation Audit Script - Scans all domain computers for Visio installations and last usage.

**Description:**
- This script queries Active Directory for all computers, then uses WMI/Registry to check for Visio installations. Supports Visio Professional/Standard 2016, 2019, and Office 365/2021 (x64). Generates CSV and HTML reports.

**Features:**
- x64-only support (Office 365/2021/2019/2016)
- Dynamic script path detection ($PSScriptRoot)
- ComputerPrefix filtering (GOT* prefix by default)
- Targeted OU search with configurable SearchBase
- Parallel processing with configurable thread count
- CSV and HTML report generation
- Last access time tracking for Visio installations

**Parameters:**
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-OutputPath` | string | Script directory\Output\VisioAudit | Directory to save reports |
| `-ComputerFilter` | string | `*` | Filter for AD computer search |
| `-ThreadCount` | int | `10` | Number of parallel jobs (1-20) |
| `-IncludeOfflineComputers` | switch | $false | Include offline computers in scan |
| `-ComputerPrefix` | string | `GOT` | Computer name prefix filter (e.g., GOT*) |
| `-SearchBase` | string | `OU=Workstations,OU=NEOS CIB 64,OU=SE,OU=CRDF,DC=euro,DC=net,DC=intra` | LDAP path to the OU to search |
| `-ScanCredential` | PSCredential | _None_ | Optional credential that is local admin on target hosts; pass `(Get-Credential)` when you lack a domain admin account. |

**Usage Examples:**
```powershell
# Basic audit with default settings
.\Visio-Enterprise-Audit.ps1

# Audit with custom output path and thread count
.\Visio-Enterprise-Audit.ps1 -OutputPath "C:\Reports" -ThreadCount 20

# Scan computers with specific prefix
.\Visio-Enterprise-Audit.ps1 -ComputerPrefix "GOTM007"

# Scan specific OU with increased threads
.\Visio-Enterprise-Audit.ps1 -SearchBase "OU=Workstations,OU=NEOS CIB 64,OU=SE,OU=CRDF,DC=euro,DC=net,DC=intra" -ThreadCount 15
```

**Default SearchBase:**
```
OU=Workstations,OU=NEOS CIB 64,OU=SE,OU=CRDF,DC=euro,DC=net,DC=intra
```

---

## Requirements

### Prerequisites
- **Windows PowerShell 5.1+**
- **ActiveDirectory** module
- **Administrator privileges**
- Domain-joined computer with network access

### Windows 11 Setup
Run PowerShell as Administrator:
```powershell
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```

**Windows Server:** Prerequisites are pre-installed.

### Local admin scanning (Access error 397)

If you do not have a domain admin account, supply a local admin credential via `-ScanCredential`. The audit uses that credential for CIM/WinRM calls and, when WMI access is denied (access error 397), falls back to a remote helper script that runs on each host via `Invoke-Command`.

```powershell
$cred = Get-Credential            # prompt for the local admin account shared across the OU
.\Visio-Enterprise-Audit.ps1 -ScanCredential $cred -ThreadCount 20 -SearchBase "<Your OU>"

Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" `
    -Name "LocalAccountTokenFilterPolicy" -Value 1

Enable-NetFirewallRule -DisplayGroup "Windows Management Instrumentation (WMI)"
Enable-PSRemoting -Force          # ensures WinRM listens on the targets
```

Access error 397 indicates CIM/WMI communication is blocked; run `Visio-Helper-Utils.ps1` option 13 for the detailed remediation steps (credential caching, firewall rules, and helper guidance) and use `Office-Version-Detector.ps1` locally to confirm Microsoft 365 Apps 10.0.60910 plus Visio 2016 Professional/Standard installs before rerunning the audit.

Enabling `LocalAccountTokenFilterPolicy` stops Windows from stripping the remote token from the local account, while the firewall rule and PSRemoting ensure `Invoke-Command` can connect. Run these commands per host or push via Group Policy before the audit.

**New helpers:**  
Run `.\Visio-Helper-Utils.ps1` option 18 to remediate Access 397 in bulk (LocalAccountTokenFilterPolicy, WMI firewall, PSRemoting) and save a JSON report of the changes. Use option 19 to dump a lightweight JSON snapshot of the latest audit output and optionally post it to a webhook-ready API so downstream tooling can ingest the counts/errors/output paths.

---

## Quick Start

### 1. Install Prerequisites
Run PowerShell as Administrator:
```powershell
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```

### 2. Run the Audit
```powershell
cd C:\automation-package
.\Visio-Enterprise-Audit.ps1
```

### 3. View Reports
Reports automatically generated in the script's Output\VisioAudit directory:
- `VisioAudit_YYYYMMDD_HHMMSS.csv` - Data export
- `VisioAudit_YYYYMMDD_HHMMSS.html` - Beautiful dashboard

---

## 📊 What Gets Scanned

✓ Office 365 Visio installations  
✓ Visio 2016 & 2019 Professional/Standard  
✓ x64-only support (Office 365/2021/2019/2016)  
✓ Last used dates  
✓ Version information  
✓ Installation paths  
✓ Online/offline status  
✓ Office 365 subscription detection  
✓ Microsoft 365 Apps 10.0.60910 plus Visio 2016 Professional/Standard via MSI/uninstall detection  

---

## 📈 Report Examples

### CSV Output
```
ComputerName,IsOnline,VisioInstalled,VisioVersion,Office365,LastUsedDate,InstallPath
WS-001,Yes,Yes,16.0.14931,Yes,2024-01-15 14:30:22,C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE
WS-002,Yes,No,N/A,No,N/A,N/A
WS-003,No,Unknown,N/A,N/A,N/A,N/A
```

### HTML Report
- Dashboard with key metrics
- Installation summary table
- Office 365 vs Desktop breakdown
- Offline computer list
- Responsive mobile-friendly design

---

## 🔧 Common Commands

```powershell
# Basic audit
.\Visio-Enterprise-Audit.ps1

# Audit with custom output path
.\Visio-Enterprise-Audit.ps1 -OutputPath "C:\Reports\Visio"

# Scan specific department with prefix
.\Visio-Enterprise-Audit.ps1 -ComputerPrefix "GOTM007"

# Faster scanning (more threads)
.\Visio-Enterprise-Audit.ps1 -ThreadCount 20

# Office version detection
.\Office-Version-Detector.ps1

# Office detection with verbose logging
.\Office-Version-Detector.ps1 -VerboseLogging -LogFilePath "C:\Logs\Office.log"

# View latest report
Import-Csv ".\Output\VisioAudit\VisioAudit_*.csv" | Format-Table
```

---

## 🆘 Troubleshooting

### Error: "File cannot be loaded. The file is not digitally signed"

Run as Administrator:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```

Or use bypass:
```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\Visio-Enterprise-Audit.ps1"
```

### Error: "Active Directory Module is not loaded"

Install on Windows 11:
```powershell
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
```

### Slow Performance

Reduce thread count:
```powershell
.\Visio-Enterprise-Audit.ps1 -ThreadCount 5
```

Or filter by prefix:
```powershell
.\Visio-Enterprise-Audit.ps1 -ComputerPrefix "GOT"
```

---

## 📋 File Structure

```
visio-enterprise-audit-suite/
├── README.md                          (This file)
├── Office-Version-Detector.ps1        (Office version detection)
├── Visio-Enterprise-Audit.ps1        (Main scanner)
├── Visio-Usage-Analytics.ps1         (Usage tracking)
├── Visio-Helper-Utils.ps1            (Interactive menu)
├── documentation/
│   ├── VISIO_AUDIT_GUIDE.md          (Detailed documentation)
│   ├── DEPLOYMENT.md                 (Deployment guide)
│   ├── TROUBLESHOOTING.md            (Troubleshooting guide)
│   └── ...
└── CHANGELOG.md                      (Version history)
```

---

## 🎯 Use Cases

### Compliance Auditing
- Track Visio installations across domain
- Verify Office 365 license usage
- Generate audit reports for compliance teams

### Cost Analysis
- Calculate total Visio licenses in use
- Identify unused installations (can be removed)
- Estimate annual licensing costs

### Usage Monitoring
- Identify which departments use Visio
- Track last usage dates
- Monitor Visio document access patterns

### Office Version Validation
- Validate only Office 365/2019/2016 installations
- Reject unsupported Office versions
- Generate compliance reports

---

## ⚙️ Advanced Features

### Scheduled Automation
Create weekly automated scans with option 9:
```powershell
.\Visio-Helper-Utils.ps1
# Select option 9: Schedule Recurring Audit
```
The helper builds a scheduled task (runs as SYSTEM) and preserves the configured prefix, thread count, LDAP scope, and intended start time. Use option 14 to inspect the task’s next/last run and verify success.

### Email & Webhook Notifications
Option 8 now bundles the latest audit CSV/HTML into email/webhook notifications, optionally zipping attachments before sending, and logs totals/errors inline.
```powershell
.\Visio-Helper-Utils.ps1
# Select option 8: Send Report Notification (email + webhook)
```

### Credential Caching
Option 15 caches the local admin credential in `VisioScanCredential.txt` in the script root, encrypted with DPAPI so only the current user can read it. Clear it via the same option if you need to rotate credentials or share the suite with another operator.

### Report Retention
Option 16 runs `Remove-OldReports` and removes CSV/HTML pairs older than 30 days (default) or beyond a custom file count. Use it manually or schedule a reminder so the `Output\VisioAudit` folder stays under control.

### Health Check Dashboard
Option 17 runs `Invoke-VisioHealthCheck`, validating AD connectivity, WinRM reachability, scheduled task health, and recent report availability, then writes `Output\VisioAudit\VisioHealthStatus.html` for a quick compliance snapshot.

### Excel Export
Export to formatted Excel workbooks:
```powershell
.\Visio-Helper-Utils.ps1
# Select option 4: Export Latest Report to Excel
```

---

## 📊 Performance Benchmarks

| Scenario | Computers | Time | Threads |
|----------|-----------|------|---------|
| Small Business | 50 | 5-10 min | 5 |
| Medium Enterprise | 200 | 15-25 min | 10 |
| Large Enterprise | 500 | 30-45 min | 15 |
| Very Large | 1000+ | 60-90 min | 20 |

---

## 🔐 Security Notes

- Scripts require Administrator privileges
- No data is sent to external services
- Reports stored locally in script's Output\VisioAudit directory
- Requires domain admin/delegated permissions
- WMI/Registry access needed for detailed scanning

---

## 📞 Support & Documentation

**Full documentation available in:**
- `documentation/VISIO_AUDIT_GUIDE.md` - Complete reference guide
- `documentation/DEPLOYMENT.md` - Deployment instructions
- `documentation/TROUBLESHOOTING.md` - Common issues & solutions

**For issues:**
1. Check `documentation/TROUBLESHOOTING.md`
2. Review error messages in CSV reports
3. Verify prerequisites are installed
4. Check domain connectivity
5. Verify admin privileges

---

## 📝 Version

**Version:** 2.0.0  
**Release Date:** 2026  
**Tested On:** 
- Windows Server 2019, 2022
- Windows 11 (with RSAT tools)
- PowerShell 5.1+
- Active Directory 2008 R2+

---

## 📄 License

These scripts are provided for enterprise IT administration purposes.
Use freely within your organization.

---

## 🎉 Getting Started

1. **Extract the ZIP file**
2. **Read this README.md** (you are here!)
3. **Run the setup command** (for Windows 11 only):
   ```powershell
   Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
   ```
4. **Execute Visio-Enterprise-Audit.ps1** with administrator privileges
5. **View reports** in the script's Output\VisioAudit directory

That's it! Enjoy comprehensive Visio auditing! 🚀

---

**Need help?** See `documentation/VISIO_AUDIT_GUIDE.md` for detailed documentation.

Note: Office-Version-Detector now recognizes Microsoft 365 10.0.60910 (2408) Visio builds and Office 2016 installs as supported Professional/Standard installs so the suite matches the audit logic.
