# Troubleshooting Guide

## Purpose

This guide provides common issues, root causes, and validated fixes for the Visio Enterprise Audit Suite.

Primary scripts covered:
- `Visio-Enterprise-Audit.ps1`
- `Visio-Usage-Analytics.ps1`
- `Visio-Helper-Utils.ps1`
- `Office-Version-Detector.ps1`

## Fast Triage Checklist

1. Run PowerShell as Administrator.
2. Confirm AD module is available (`Import-Module ActiveDirectory`).
3. Confirm domain connectivity and DNS resolution.
4. Start with `-ThreadCount 5` on unstable networks.
5. Run against a small OU first (`-SearchBase`).

## Execution Policy Errors

Symptoms:
- Script blocked due to signing policy.

Fix:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```

Temporary process-only bypass:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

## ActiveDirectory Module Missing

Symptoms:
- `Import-Module ActiveDirectory` fails.

Fix on Windows client:

```powershell
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
Import-Module ActiveDirectory
```

## AD Query Returns No Computers

Symptoms:
- Main audit exits with no matching computers.

Checks:
- Validate `-ComputerPrefix` (default is `GOT`).
- Validate OU DN passed to `-SearchBase`.
- If unsure, run without `-SearchBase` to scan full domain scope.

Verification command:

```powershell
Get-ADComputer -Filter "Name -like 'GOT*'" -ResultSetSize 10 | Select-Object Name
```

## Many Endpoints Show Offline

Symptoms:
- High offline count in summary.

Likely causes:
- DNS/routing/firewall issues
- ICMP blocked
- Endpoints actually offline

Actions:
- Confirm host resolution with `Resolve-DnsName`.
- Test representative systems with `Test-Connection`.
- Reduce `-ThreadCount` and re-run.

## WMI/CIM Access Failures

Symptoms:
- Frequent access errors or timeout behavior.

Actions:
- Ensure endpoint admin privileges.
- Verify remote management/firewall policy for WMI/CIM.
- Test one endpoint manually:

```powershell
$s = New-CimSession -ComputerName "PC001"
Get-CimInstance -CimSession $s -ClassName Win32_OperatingSystem
Remove-CimSession $s
```

## Performance Issues

Symptoms:
- Scan appears slow or unstable.

Actions:
- Start at `-ThreadCount 5`, increase gradually.
- Split by OU with `-SearchBase`.
- Avoid running very high parallelism during peak hours.

## `CurrentUser` Is Often `Unknown`

Explanation:
- `CurrentUser` is derived from `Win32_ComputerSystem.UserName`.
- Some endpoints do not expose this reliably (locked screens, session state, permissions).

Actions:
- Validate WMI/CIM permissions.
- Cross-check disputed hosts with `Visio-Usage-Analytics.ps1`.

## `LastUsageSource` Interpretation Issues

Definitions:
- `RunningProcess`: `VISIO.EXE` currently running.
- `Prefetch`: latest `VISIO*.pf` timestamp.
- `ExecutableLastAccess`: fallback from EXE file access metadata.

Guidance:
- Prefer `RunningProcess` and `Prefetch` for operational decisions.
- Treat `ExecutableLastAccess` as low-confidence fallback.

## Empty or Unexpected Report Values

Checks:
- Confirm output path is writable.
- Validate scanner account privileges.
- Confirm target systems have Visio installed and reachable.

Expected behavior:
- Offline systems show `Error` such as "Computer offline".
- Missing usage artifacts may produce `LastUsageSource = N/A`.

## Office Detector Errors

Script:
- `Office-Version-Detector.ps1`

Useful runs:

```powershell
.\Office-Version-Detector.ps1 -VerboseLogging
.\Office-Version-Detector.ps1 -StrictErrorHandling -LogFilePath "C:\Logs\Office-Version-Detection.log"
```

Exit codes:
- `0`: supported version detected
- `1`: unsupported version or failure

## Escalation Path

If unresolved:
1. Capture command used.
2. Save first 50 lines of console output.
3. Save generated CSV/HTML sample.
4. Capture endpoint-specific error examples.
5. Re-run a single-OU scope and compare.

## Related Documentation

- `documentation/USER_GUIDE.md`
- `documentation/VISIO_AUDIT_GUIDE.md`
- `documentation/DEPLOYMENT.md`

