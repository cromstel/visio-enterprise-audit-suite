# Deployment Guide

## Purpose

This guide covers production-style deployment of the Visio Enterprise Audit Suite for recurring enterprise audits.

## Recommended Deployment Model

Host location:
- Domain-joined management server or admin workstation

Execution identity:
- Service/admin account with:
- AD read permissions
- Remote endpoint audit permissions
- Local admin rights where required

Data location:
- Secure folder or controlled network share for CSV/HTML reports

## Package Layout

Expected key files:
- `Visio-Enterprise-Audit.ps1`
- `Visio-Usage-Analytics.ps1`
- `Visio-Helper-Utils.ps1`
- `Office-Version-Detector.ps1`
- `documentation/USER_GUIDE.md`

## Pre-Deployment Checklist

- PowerShell 5.1+ available
- ActiveDirectory module installed
- Execution policy allows script run
- Target network paths/firewall rules validated
- Output directory access validated
- Initial pilot OU identified

## Step 1: Install Prerequisites

```powershell
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
Import-Module ActiveDirectory
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```

## Step 2: Validate Environment

```powershell
Get-ADComputer -Filter * -ResultSetSize 5 | Select-Object Name
Test-Path .\Visio-Enterprise-Audit.ps1
```

Optional endpoint connectivity check:

```powershell
Test-Connection -ComputerName "PC001" -Count 1
```

## Step 3: Pilot Run

Use a narrow scope first:

```powershell
.\Visio-Enterprise-Audit.ps1 `
  -SearchBase "OU=Pilot,DC=contoso,DC=com" `
  -ThreadCount 5 `
  -OutputPath "C:\Reports\Visio\Pilot"
```

Validate:
- CSV and HTML generated
- `CurrentUser` and `LastUsageSource` columns populated as expected
- Error volume is acceptable

## Step 4: Production Rollout

Example full/large OU run:

```powershell
.\Visio-Enterprise-Audit.ps1 `
  -SearchBase "OU=Workstations,DC=contoso,DC=com" `
  -ThreadCount 15 `
  -OutputPath "C:\Reports\Visio\Production"
```

Tuning guidance:
- Start with `10` threads.
- Increase gradually to `15-25` only if network and endpoints remain stable.
- Hard limit in script is `64`.

## Step 5: Schedule Recurring Audits

Create a weekly task:

```powershell
$scriptPath = "C:\PROJECTS\visio-enterprise-audit-suite\Visio-Enterprise-Audit.ps1"
$taskName = "VisioAudit-Weekly"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At "02:00"
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""

Register-ScheduledTask `
  -TaskName $taskName `
  -Trigger $trigger `
  -Action $action `
  -RunLevel Highest `
  -Description "Weekly Visio installation and usage audit" `
  -Force
```

Verify:

```powershell
Get-ScheduledTask -TaskName "VisioAudit-Weekly" | Select-Object TaskName, State
```

## Step 6: Retention and Reporting Operations

Recommended retention:
- Keep at least 12 months of CSV history
- Archive older reports to compressed storage

Suggested structure:
- `C:\Reports\Visio\YYYY\MM\`

## Security Recommendations

- Restrict report directory permissions (user and machine inventory data is sensitive).
- Use least-privilege for scheduled identities.
- Avoid emailing raw reports to broad distribution lists.
- Prefer protected shares with audited access.

## Validation Commands (Post-Deployment)

Quick sanity checks:

```powershell
Import-Csv "C:\Reports\Visio\Production\VisioAudit_*.csv" | Select-Object -First 5
Import-Csv "C:\Reports\Visio\Production\VisioAudit_*.csv" | Group-Object LastUsageSource
```

## Rollback

If deployment causes operational issues:
1. Disable scheduled task.
2. Revert to previous script backup.
3. Run pilot OU with lower `-ThreadCount`.
4. Re-validate endpoint access before re-enabling schedule.

## Related Documentation

- `documentation/USER_GUIDE.md`
- `documentation/VISIO_AUDIT_GUIDE.md`
- `documentation/TROUBLESHOOTING.md`

