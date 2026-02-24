# Visio Audit Guide

## Purpose

This document is the operational guide for running Visio audits in an Active Directory environment.

Canonical reference:
- `documentation/USER_GUIDE.md`

This file provides a concise, execution-focused runbook for day-to-day use.

## Scope

Primary scripts:
- `Visio-Enterprise-Audit.ps1`
- `Visio-Usage-Analytics.ps1`
- `Visio-Helper-Utils.ps1`

## Prerequisites

- Windows PowerShell 5.1 or later
- ActiveDirectory PowerShell module
- Administrative privileges
- Domain connectivity to target computers

Install AD module on Windows client systems:

```powershell
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
Import-Module ActiveDirectory
```

## Main Audit Workflow

Run from repository root:

```powershell
cd C:\PROJECTS\visio-enterprise-audit-suite
.\Visio-Enterprise-Audit.ps1
```

Key parameters:

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-OutputPath` | string | `.\Output\VisioAudit` | CSV/HTML output folder |
| `-ComputerFilter` | string | `*` | Compatibility parameter |
| `-ThreadCount` | int | `10` | Valid range: `1..64` |
| `-ComputerPrefix` | string | `GOT` | AD name prefix |
| `-SearchBase` | string | none | Scan full domain when omitted |

Examples:

```powershell
.\Visio-Enterprise-Audit.ps1 -ThreadCount 20
.\Visio-Enterprise-Audit.ps1 -ComputerPrefix "ENG"
.\Visio-Enterprise-Audit.ps1 -SearchBase "OU=Workstations,DC=contoso,DC=com" -ThreadCount 15
.\Visio-Enterprise-Audit.ps1 -OutputPath "C:\Reports\Visio"
```

## Report Outputs

Generated files:
- `VisioAudit_YYYYMMDD_HHMMSS.csv`
- `VisioAudit_YYYYMMDD_HHMMSS.html`

Default location:
- `.\Output\VisioAudit`

## Important Tracking Fields

The main audit now includes:
- `CurrentUser`
- `LastUsedDate`
- `LastUsageSource`

`LastUsageSource` values:
- `RunningProcess`
- `Prefetch`
- `ExecutableLastAccess`
- `N/A`

Recommended interpretation:
- Highest confidence: `RunningProcess`
- Medium confidence: `Prefetch`
- Lowest confidence: `ExecutableLastAccess`

## Usage Analytics Workflow

Use analytics after baseline inventory:

```powershell
.\Visio-Usage-Analytics.ps1 -ComputerNames @("PC001","PC002")
```

Useful for deeper investigation when:
- `CurrentUser` is `Unknown`
- `LastUsageSource` is low-confidence
- Endpoint usage is disputed

## Weekly Runbook

1. Run baseline inventory (`Visio-Enterprise-Audit.ps1`).
2. Review CSV for active Visio estate and weak-confidence usage data.
3. Compare with previous report using `Visio-Helper-Utils.ps1`.
4. Run targeted analytics for exceptions.
5. Archive outputs for trend/compliance evidence.

## Known Limits

- `CurrentUser` may be `Unknown` on locked or inaccessible endpoints.
- Prefetch artifacts may be unavailable due to endpoint policy.
- Last-access timestamps can be noisy in hardened environments.

## Related Documents

- `documentation/USER_GUIDE.md`
- `documentation/TROUBLESHOOTING.md`
- `documentation/DEPLOYMENT.md`

