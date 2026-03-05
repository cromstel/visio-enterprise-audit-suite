# Visio Enterprise Audit Suite User Guide

## 1. Purpose

This guide explains how to install, run, and operate the Visio Enterprise Audit Suite in a Windows Active Directory environment.

It covers:
- Prerequisites and setup
- Script usage patterns
- Report interpretation
- New tracking fields: `CurrentUser` and `LastUsageSource`
- Troubleshooting and operational best practices

## 2. Scripts in This Repository

### `Visio-Enterprise-Audit.ps1`
- Main domain audit script.
- Discovers domain computers from AD and scans for Visio installations.
- Produces CSV and HTML reports.
- Includes parallel scan execution.

### `Visio-Usage-Analytics.ps1`
- Detailed usage analytics for specified or discovered computers.
- Collects process/activity-oriented signals.
- Generates an HTML analytics report.

### `Visio-Helper-Utils.ps1`
- Interactive utility menu for common operational tasks:
- Run full audit
- Compare reports
- Generate summaries
- Cost analysis
- Scheduling and email helper actions

### `Office-Version-Detector.ps1`
- Local machine Office detector.
- Supports Office 365, Office 2019, and Office 2016 installations.
- Returns exit code `0` for supported versions, `1` otherwise.

## 3. Prerequisites

Required:
- Windows PowerShell 5.1+
- Active Directory module
- Domain connectivity to target computers
- Administrative privileges

Recommended:
- Run from a management server/jump box with stable network connectivity
- Execution policy allowing local scripts

Example setup:

```powershell
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
Import-Module ActiveDirectory
```

## 4. Quick Start

From the repository root:

```powershell
cd C:\PROJECTS\visio-enterprise-audit-suite
.\Visio-Enterprise-Audit.ps1
```

The script will:
- Determine the output directory (`.\Output\VisioAudit` by default)
- Query AD for computer objects by prefix (default `GOT*`)
- Scan computers in parallel
- Save CSV + HTML reports

## 5. Main Audit Script Reference

Script: `Visio-Enterprise-Audit.ps1`

Parameters:

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-OutputPath` | `string` | `.\Output\VisioAudit` | Report destination |
| `-ComputerFilter` | `string` | `*` | Maintained for compatibility |
| `-ThreadCount` | `int` | `10` | Valid range `1..64` |
| `-IncludeOfflineComputers` | `switch` | `false` | Compatibility flag |
| `-ComputerPrefix` | `string` | `GOT` | AD query uses `Name -like "<prefix>*"` |
| `-SearchBase` | `string` | `null` | If omitted, scans the full domain scope |

Examples:

```powershell
.\Visio-Enterprise-Audit.ps1 -ThreadCount 20
.\Visio-Enterprise-Audit.ps1 -ComputerPrefix "SALES"
.\Visio-Enterprise-Audit.ps1 -SearchBase "OU=Workstations,DC=contoso,DC=com" -ThreadCount 15
.\Visio-Enterprise-Audit.ps1 -OutputPath "C:\Reports\Visio"
```

## 6. Understanding Tracking Fields

The audit now exposes these usage/user fields per computer:

### `CurrentUser`
- Source: `Win32_ComputerSystem.UserName`
- Meaning: Interactive logged-on user at scan time (if available)
- Common fallback: `Unknown`

### `LastUsedDate`
- Best available timestamp for likely recent Visio usage

### `LastUsageSource`
- Indicates where `LastUsedDate` came from:
- `RunningProcess`: `VISIO.EXE` currently running (timestamp is scan time)
- `Prefetch`: latest `C:\Windows\Prefetch\VISIO*.pf` `LastWriteTime`
- `ExecutableLastAccess`: fallback to Visio executable file access time
- `N/A`: no usable usage signal found

### Source reliability guidance

| Source | Reliability | Notes |
|---|---|---|
| `RunningProcess` | High (current state) | Confirms active use during scan |
| `Prefetch` | Medium/High | Depends on prefetch availability/policy |
| `ExecutableLastAccess` | Low/Medium | Last access time can be disabled/noisy |

## 7. Output Files and Columns

Default output location:
- `.\Output\VisioAudit`

Files generated per run:
- `VisioAudit_YYYYMMDD_HHMMSS.csv`
- `VisioAudit_YYYYMMDD_HHMMSS.html`

CSV columns:
- `ComputerName`
- `IsOnline`
- `VisioInstalled`
- `CurrentUser`
- `VisioVersion`
- `VisioEdition`
- `Office365`
- `LastUsedDate`
- `LastUsageSource`
- `InstallPath`
- `Error`

## 8. Usage Analytics Script

Script: `Visio-Usage-Analytics.ps1`

Primary use:
- Deep dive after the baseline audit identifies target computers.

Parameters:

| Parameter | Type | Default |
|---|---|---|
| `-OutputPath` | `string` | `.\Output\VisioAudit` |
| `-ComputerFilter` | `string` | `*` |
| `-ThreadCount` | `int` | `10` (`1..64`) |
| `-IncludeOfflineComputers` | `switch` | `false` |
| `-ComputerPrefix` | `string` | `GOT` |
| `-SearchBase` | `string` | OU default in script |
| `-ComputerNames` | `string[]` | empty |

Examples:

```powershell
.\Visio-Usage-Analytics.ps1 -ComputerNames @("PC001","PC002")
.\Visio-Usage-Analytics.ps1 -ComputerPrefix "ENG" -SearchBase "OU=Engineering,DC=contoso,DC=com"
```

## 9. Helper Script Operations

Script: `Visio-Helper-Utils.ps1`

Run:

```powershell
.\Visio-Helper-Utils.ps1
```

Provides menu-driven actions for:
- Running the full audit
- Viewing latest report summary
- Comparing latest two reports
- Cost analysis
- Department filtering
- Optional scheduled task/email helpers

## 10. Office Detector

Script: `Office-Version-Detector.ps1`

Examples:

```powershell
.\Office-Version-Detector.ps1
.\Office-Version-Detector.ps1 -VerboseLogging
.\Office-Version-Detector.ps1 -StrictErrorHandling -LogFilePath "C:\Logs\Office-Version-Detection.log"
```

Exit codes:
- `0`: supported Office version detected
- `1`: unsupported version or detection failure

## 11. Operational Runbook

Recommended weekly workflow:
1. Run `Visio-Enterprise-Audit.ps1` with appropriate `-SearchBase` and `-ThreadCount`.
2. Review CSV for:
   - `VisioInstalled = Yes`
   - `CurrentUser = Unknown` (follow-up candidate)
   - `LastUsageSource` quality (`RunningProcess`/`Prefetch` preferred)
3. Use `Visio-Helper-Utils.ps1` for summary and comparison.
4. Use `Visio-Usage-Analytics.ps1` for high-value or ambiguous endpoints.

## 12. Troubleshooting

### No computers returned
- Validate AD module load and permissions.
- Confirm `-ComputerPrefix` and `-SearchBase` are correct.

### Many offline systems
- Validate DNS/firewall/routing to endpoints.
- Reduce `-ThreadCount` to reduce scan pressure.

### `CurrentUser` often `Unknown`
- Expected on locked/no-user systems or inaccessible session data.
- Validate remote WMI/CIM access and endpoint policy.

### `LastUsageSource` mostly `ExecutableLastAccess`
- Prefetch may be disabled/cleared by policy.
- This is still valid fallback, but lower confidence.

### Access denied errors
- Run elevated PowerShell.
- Validate endpoint admin rights and remote management permissions.

## 13. Performance and Safety Recommendations

- Start with `-ThreadCount 10` and tune gradually.
- Large domains: scan per OU using `-SearchBase`.
- Keep historical CSV files for trend/comparison.
- Avoid very high thread counts during business hours.

## 14. Security and Data Handling

- Reports may include usernames and hostnames.
- Store outputs in access-controlled locations.
- Treat exported files as operationally sensitive inventory data.

## 15. Version Notes

This guide reflects current repository behavior including:
- Professional ASCII-safe scan interface
- `CurrentUser` tracking
- `LastUsageSource` tracking with source precedence
