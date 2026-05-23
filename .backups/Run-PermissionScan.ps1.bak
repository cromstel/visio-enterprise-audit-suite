#Requires -Version 5.1
<#
.SYNOPSIS
    Launcher wrapper for Scan-FolderPermissions.ps1.
    Provides preset resource profiles and optional Task Scheduler registration.

.DESCRIPTION
    Run this script instead of calling the scanner directly.
    Choose a resource profile that suits your server load, then optionally
    register a scheduled task to run the scan off-hours automatically.

    Reports are saved in a timestamped subfolder created next to this script.

.PARAMETER Profile
    Resource throttle profile:
      Safe     - Very slow; minimal I/O impact. Best during business hours.
      Balanced - Moderate; recommended default.
      Fast     - Faster; use only in maintenance windows.

.PARAMETER DriveRoot
    Root path to scan. Default: K:\

.PARAMETER SkipBuiltIn
    Omit built-in system/CREATOR OWNER accounts from the report.

.PARAMETER IncludeInherited
    Include inherited (not just explicit) permission entries.

.PARAMETER ScheduleTask
    Register a Windows Scheduled Task to run this scan nightly.

.PARAMETER TaskTime
    Time for the scheduled task (24h format, e.g. "02:00"). Default: 02:00

.PARAMETER TaskUser
    Windows account the scheduled task runs as. Default: SYSTEM

.EXAMPLE
    # Balanced scan on K:\ — reports saved next to this script
    .\Run-PermissionScan.ps1 -Profile Balanced

.EXAMPLE
    # Low-impact scan, skip built-in accounts
    .\Run-PermissionScan.ps1 -Profile Safe -SkipBuiltIn

.EXAMPLE
    # Schedule a nightly Safe-profile scan at 01:30 AM
    .\Run-PermissionScan.ps1 -Profile Safe -ScheduleTask -TaskTime "01:30"

.NOTES
    Both Run-PermissionScan.ps1 and Scan-FolderPermissions.ps1 must be in the same folder.
    Version 3.0: Fixed hardcoded report path; fixed %DATE% scheduled task token;
                 fixed ErrorActionPreference; added PSScriptRoot null guard.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("Safe","Balanced","Fast")]
    [string]$Profile = "Balanced",

    [Parameter(Mandatory = $false)]
    [string]$DriveRoot = "K:\",

    [Parameter(Mandatory = $false)]
    [switch]$SkipBuiltIn,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeInherited,

    [Parameter(Mandatory = $false)]
    [switch]$ScheduleTask,

    [Parameter(Mandatory = $false)]
    [string]$TaskTime = "02:00",

    [Parameter(Mandatory = $false)]
    [string]$TaskUser = "SYSTEM"
)

Set-StrictMode -Version Latest
# FIX: Was "Stop" — that caused the launcher to abort on any non-critical warning
# surfaced by the scanner child script. "Continue" is correct for a launcher.
$ErrorActionPreference = "Continue"

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Resolve Script Directory
# FIX: Guard against $PSScriptRoot being empty (e.g. when dot-sourced).
# ─────────────────────────────────────────────────────────────────────────────

$ScriptDir = if ($PSScriptRoot -and $PSScriptRoot -ne "") {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    (Get-Location).Path
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Profile Definitions
# ─────────────────────────────────────────────────────────────────────────────

$profiles = @{
    Safe = @{
        BatchSize   = 20
        SleepMs     = 800
        Description = "Very gentle — minimal CPU/IO impact. Safe during business hours."
    }
    Balanced = @{
        BatchSize   = 50
        SleepMs     = 200
        Description = "Moderate pacing. Recommended for off-hours use."
    }
    Fast = @{
        BatchSize   = 200
        SleepMs     = 50
        Description = "Aggressive throughput. Maintenance windows only."
    }
}

$selectedProfile = $profiles[$Profile]

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Output Path
# FIX: Reports now save in a timestamped subfolder NEXT TO this script,
# not the hardcoded C:\Reports\FolderPermissions path.
# ─────────────────────────────────────────────────────────────────────────────

$Timestamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFolder = Join-Path $ScriptDir "ScanReports\$Timestamp"

if (-not (Test-Path $ReportFolder)) {
    try {
        New-Item -ItemType Directory -Path $ReportFolder -Force | Out-Null
        Write-Host "[INFO] Created report folder: $ReportFolder" -ForegroundColor Cyan
    }
    catch {
        Write-Host "[ERROR] Could not create report folder '$ReportFolder': $_" -ForegroundColor Red
        exit 1
    }
}

$csvFile = Join-Path $ReportFolder "FolderPermissions_${Profile}_$Timestamp.csv"

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Locate Scanner Script
# ─────────────────────────────────────────────────────────────────────────────

$scannerScript = Join-Path $ScriptDir "Scan-FolderPermissions.ps1"

if (-not (Test-Path $scannerScript)) {
    Write-Host "[ERROR] Cannot find Scan-FolderPermissions.ps1 at: $scannerScript" -ForegroundColor Red
    Write-Host "        Both scripts must reside in the same folder." -ForegroundColor Red
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Optional Scheduled Task Registration
# FIX: Removed %DATE% CMD environment variable token — it is not expanded in
# PowerShell argument strings and produced a literal "%DATE%" filename.
# The scanner now generates its own timestamped filename when OutputPath is
# omitted, so the scheduled task simply does not pass -OutputPath.
# ─────────────────────────────────────────────────────────────────────────────

if ($ScheduleTask) {
    Write-Host "`n[*] Registering Windows Scheduled Task..." -ForegroundColor Yellow

    $taskName = "FileShare_PermissionScan_$Profile"

    # Build the -File argument for powershell.exe.
    # Note: OutputPath is intentionally omitted so the scanner generates its own
    # timestamped path relative to the script location on each run.
    $skipArg    = if ($SkipBuiltIn)      { " -SkipBuiltIn" }      else { "" }
    $inheritArg = if ($IncludeInherited) { " -IncludeInherited" }  else { "" }

    $scriptArgs = "-NonInteractive -NoProfile -ExecutionPolicy Bypass " +
                  "-File `"$scannerScript`" " +
                  "-TargetDrive `"$DriveRoot`" " +
                  "-BatchSize $($selectedProfile.BatchSize) " +
                  "-SleepMs $($selectedProfile.SleepMs)" +
                  $skipArg + $inheritArg

    try {
        $action    = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $scriptArgs
        $trigger   = New-ScheduledTaskTrigger -Daily -At $TaskTime
        $settings  = New-ScheduledTaskSettingsSet `
                         -RunOnlyIfIdle:$false `
                         -StartWhenAvailable `
                         -ExecutionTimeLimit (New-TimeSpan -Hours 8) `
                         -Priority 7   # 7 = BelowNormal (5=Normal, 9=Idle)
        $principal = New-ScheduledTaskPrincipal -UserId $TaskUser -RunLevel Highest

        Register-ScheduledTask `
            -TaskName    $taskName `
            -Action      $action `
            -Trigger     $trigger `
            -Settings    $settings `
            -Principal   $principal `
            -Description "Automated file-share permission scan ($Profile profile) — FileShare Audit v3.0" `
            -Force | Out-Null

        Write-Host ""
        Write-Host "[OK] Scheduled Task registered: '$taskName'" -ForegroundColor Green
        Write-Host "     Runs daily at $TaskTime as: $TaskUser" -ForegroundColor Green
        Write-Host "     Reports saved to: $ScriptDir\ScanReports\<timestamp>\" -ForegroundColor Green
        Write-Host ""
        Write-Host "     Manage task:" -ForegroundColor DarkGray
        Write-Host "       View  : Get-ScheduledTask -TaskName '$taskName'" -ForegroundColor DarkGray
        Write-Host "       Run   : Start-ScheduledTask -TaskName '$taskName'" -ForegroundColor DarkGray
        Write-Host "       Remove: Unregister-ScheduledTask -TaskName '$taskName' -Confirm:`$false" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "[ERROR] Failed to register scheduled task: $_" -ForegroundColor Red
    }

    return
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Banner
# ─────────────────────────────────────────────────────────────────────────────

$maxW = 48  # Interior width for dynamic values in the banner box

function Show-BannerLine {
    param([string]$Label, [string]$Value)
    $display = if ($Value.Length -gt $maxW) { $Value.Substring(0, $maxW - 3) + "..." } else { $Value.PadRight($maxW) }
    Write-Host "║  $($Label.PadRight(12)): $display║" -ForegroundColor Magenta
}

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor DarkMagenta
Write-Host "║              PERMISSION SCAN LAUNCHER  v3.0                 ║" -ForegroundColor DarkMagenta
Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor DarkMagenta
Show-BannerLine "Profile"    "$Profile - $($selectedProfile.Description.Substring(0, [math]::Min(35, $selectedProfile.Description.Length)))"
Show-BannerLine "Drive"      $DriveRoot
Show-BannerLine "Batch Size" "$($selectedProfile.BatchSize) folders"
Show-BannerLine "Sleep"      "$($selectedProfile.SleepMs) ms between batches"
Show-BannerLine "Report Dir" $ReportFolder
Show-BannerLine "Output CSV" $csvFile
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor DarkMagenta
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Launch Scanner
# ─────────────────────────────────────────────────────────────────────────────

$scanParams = @{
    TargetDrive = $DriveRoot
    OutputPath  = $csvFile
    BatchSize   = $selectedProfile.BatchSize
    SleepMs     = $selectedProfile.SleepMs
}

if ($SkipBuiltIn)      { $scanParams["SkipBuiltIn"]      = $true }
if ($IncludeInherited) { $scanParams["IncludeInherited"] = $true }

Write-Host "[*] Launching scanner (Profile: $Profile)..." -ForegroundColor Cyan
Write-Host ""

& $scannerScript @scanParams

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Post-scan Excel Import Tips
# ─────────────────────────────────────────────────────────────────────────────

if (Test-Path $csvFile) {
    # Count data rows (subtract 1 for header line)
    $rowCount = ((Get-Content $csvFile -ErrorAction SilentlyContinue) |
                 Measure-Object -Line).Lines - 1

    Write-Host ""
    Write-Host "================================================================" -ForegroundColor DarkGray
    Write-Host " EXCEL IMPORT GUIDE" -ForegroundColor White
    Write-Host "================================================================" -ForegroundColor DarkGray
    Write-Host " Report file  : $csvFile" -ForegroundColor Cyan
    Write-Host " Data rows    : ~$rowCount permission entries" -ForegroundColor White
    Write-Host ""
    Write-Host " How to import:" -ForegroundColor White
    Write-Host "  1. Excel > Data > Get Data > From Text/CSV > select the file" -ForegroundColor DarkGray
    Write-Host "  2. File Origin: 65001 (UTF-8)   Delimiter: Comma" -ForegroundColor DarkGray
    Write-Host "  3. Click Load (or Transform Data to preview first)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host " Useful column filters:" -ForegroundColor White
    Write-Host "  - IdentityReference  : find a specific group or user" -ForegroundColor DarkGray
    Write-Host "  - AccessControlType  : filter = Deny to spot all deny rules" -ForegroundColor DarkGray
    Write-Host "  - ACLProtected = True: folders with broken inheritance" -ForegroundColor DarkGray
    Write-Host "  - FriendlyRights     : filter = Full Control for over-privileged access" -ForegroundColor DarkGray
    Write-Host "================================================================" -ForegroundColor DarkGray
    Write-Host ""
}
