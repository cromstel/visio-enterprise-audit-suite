#Requires -Version 5.1
<#
.SYNOPSIS
    Scans all folders on a specified drive/share and exports ACL (permission) data to CSV.

.DESCRIPTION
    Recursively enumerates all folders on the target drive (default: H:\),
    reads each folder's Access Control List (ACL), and exports a structured
    report of every user/group and their assigned permissions to a CSV file.

    Resource throttling is built-in:
      - Runs at BelowNormal process priority
      - Sleeps between folder batches to avoid I/O spikes
      - Skips inaccessible folders gracefully with logged warnings

    Automatically skips:
      - $RECYCLE.BIN / RECYCLER (Recycle Bin folders)
      - System Volume Information
      - Recovery (Windows Recovery partition folder)
    All other hidden folders ARE scanned.

    Output is saved in a timestamped subfolder created next to this script file.

.PARAMETER TargetDrive
    Root path to scan. Default: H:\

.PARAMETER OutputPath
    Full path (including filename) for the CSV report.
    Default: <ScriptDir>\ScanReports\<timestamp>\FolderPermissions_<timestamp>.csv

.PARAMETER BatchSize
    Number of folders to process before pausing. Default: 50

.PARAMETER SleepMs
    Milliseconds to sleep between batches. Default: 200

.PARAMETER MaxDepth
    Maximum folder depth to recurse. 0 = unlimited. Default: 0

.PARAMETER SkipBuiltIn
    If set, skips built-in system accounts (SYSTEM, TrustedInstaller, etc.).

.PARAMETER IncludeInherited
    If set, includes inherited permission entries (default: explicit only).

.EXAMPLE
    .\Scan-FolderPermissions.ps1 -TargetDrive "H:\"

.EXAMPLE
    .\Scan-FolderPermissions.ps1 -TargetDrive "H:\" -SkipBuiltIn -BatchSize 25 -SleepMs 500

.NOTES
    Author      : FileShare Audit Script
    Version     : 3.0
    Requires    : PowerShell 5.1+, Read access to target folders
    Run As      : Administrator (or account with SeBackupPrivilege for full ACL read)
    Changes v3.0: Fixed duplicate hashtable key crash; fixed log-before-directory bug;
                  output now saved relative to script location; added system/recycle
                  folder exclusions; hidden folders preserved; banner truncation fixed;
                  PSScriptRoot null guard added; scheduled task date token fixed.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDrive = "H:\",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 500)]
    [int]$BatchSize = 50,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 5000)]
    [int]$SleepMs = 200,

    [Parameter(Mandatory = $false)]
    [int]$MaxDepth = 0,

    [Parameter(Mandatory = $false)]
    [switch]$SkipBuiltIn,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeInherited
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# -----------------------------------------------------------------------------
# REGION: Resolve Script Directory
# FIX: $PSScriptRoot is empty when dot-sourced; fall back to MyInvocation path.
# -----------------------------------------------------------------------------

$ScriptDir = if ($PSScriptRoot -and $PSScriptRoot -ne "") {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    (Get-Location).Path
}

# -----------------------------------------------------------------------------
# REGION: Setup & Configuration
# -----------------------------------------------------------------------------

$ScriptVersion = "3.0"
$ScriptStart   = Get-Date
$Timestamp     = $ScriptStart.ToString("yyyyMMdd_HHmmss")

# FIX: Default output path now resolves to a timestamped subfolder NEXT TO this
# script file, not the unpredictable current working directory.
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $reportSubFolder = Join-Path $ScriptDir "ScanReports\$Timestamp"
    $OutputPath      = Join-Path $reportSubFolder "FolderPermissions_$Timestamp.csv"
}

# Log file sits alongside the CSV (same folder, .log extension).
$LogPath = [System.IO.Path]::ChangeExtension($OutputPath, ".log")

# -----------------------------------------------------------------------------
# REGION: Output Directory Creation
# FIX: Directory must be created BEFORE the first Write-Log call, because
# Write-Log uses Add-Content which needs the directory to already exist.
# -----------------------------------------------------------------------------

$outputDir = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($outputDir) -and -not (Test-Path $outputDir)) {
    try {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        # Cannot call Write-Log here yet  -  log file dir was just created now.
        Write-Host "[INFO] Created output directory: $outputDir" -ForegroundColor Cyan
    }
    catch {
        Write-Host "[ERROR] Cannot create output directory '$outputDir': $_" -ForegroundColor Red
        exit 1
    }
}

# -----------------------------------------------------------------------------
# REGION: Built-In Accounts & System Folder Exclusions
# -----------------------------------------------------------------------------

# Built-in Windows accounts to optionally omit from the report.
$BuiltInAccounts = @(
    "NT AUTHORITY\SYSTEM",
    "NT SERVICE\TrustedInstaller",
    "BUILTIN\Administrators",
    "CREATOR OWNER",
    "NT AUTHORITY\Authenticated Users",
    "Everyone"
)

# Folder NAMES (not full paths) that are always excluded from scanning.
# These cover the Recycle Bin variants and Windows system-reserved folders.
# All other hidden folders are still scanned normally.
$ExcludedFolderNames = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
@(
    '$RECYCLE.BIN',           # Windows Vista+ Recycle Bin
    '$Recycle.Bin',           # Alternate casing
    'RECYCLER',               # Windows XP / Server 2003 Recycle Bin
    'RECYCLED',               # Older Windows Recycle Bin variant
    'System Volume Information',  # NTFS journal / restore point data
    'Recovery'                # Windows Recovery partition folder
) | ForEach-Object { [void]$ExcludedFolderNames.Add($_) }

# -----------------------------------------------------------------------------
# REGION: Helper Functions
# -----------------------------------------------------------------------------

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    try {
        Add-Content -Path $LogPath -Value $entry -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch { }
    switch ($Level) {
        "ERROR"   { Write-Host $entry -ForegroundColor Red }
        "WARN"    { Write-Host $entry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $entry -ForegroundColor Green }
        default   { Write-Host $entry -ForegroundColor Cyan }
    }
}

function Set-ThrottledPriority {
    # Drops the current PowerShell process to BelowNormal to reduce server load.
    try {
        $proc = [System.Diagnostics.Process]::GetCurrentProcess()
        $proc.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::BelowNormal
        Write-Log "Process priority set to BelowNormal for resource throttling." "INFO"
    }
    catch {
        Write-Log "Could not set process priority (non-fatal): $_" "WARN"
    }
}

function Test-ShouldExcludeFolder {
    # Returns $true if a folder should be skipped (Recycle Bin / system dirs).
    # Does NOT exclude hidden folders  -  only the named exclusion list above.
    param([string]$FolderPath)
    $folderName = [System.IO.Path]::GetFileName($FolderPath)
    return $ExcludedFolderNames.Contains($folderName)
}

function Resolve-AccountType {
    # Attempts to classify an ACL identity as Local User, Local Group,
    # Built-In, Domain Group (inferred), or SID.
    param([string]$IdentityReference)

    # Unresolved SID pattern
    if ($IdentityReference -match "^S-1-\d") {
        return "SID (Unresolved)"
    }

    # Well-known built-in prefixes  -  quick return, no ADSI call needed
    if ($IdentityReference -match "^(BUILTIN|NT AUTHORITY|NT SERVICE)\\") {
        return "Built-In / System"
    }

    # Strip domain prefix for local SAM lookups
    $bare = ($IdentityReference -split "\\")[-1]

    try {
        $localGroup = [ADSI]"WinNT://./$bare,group"
        # ADSI returns an object even for missing entries; check SchemaClassName
        if ($localGroup.SchemaClassName -eq "Group") { return "Local Group" }
    }
    catch { }

    try {
        $localUser = [ADSI]"WinNT://./$bare,user"
        if ($localUser.SchemaClassName -eq "User") { return "Local User" }
    }
    catch { }

    # Heuristic: common suffixes used by AD security groups
    if ($bare -match "(Group|Grp|Team|Dept|Admin|Users|Operators|Access|Members|Staff|Read|Write|Modify)$") {
        return "Domain Group (Inferred)"
    }

    return "Domain User / Group"
}

function Get-FriendlyRights {
    # Converts FileSystemRights enum to a human-readable permission label.
    # FIX: Removed duplicate key 1245631 that caused InvalidOperation crash.
    # FIX: Use [long] cast to safely handle large/negative enum values.
    param([System.Security.AccessControl.FileSystemRights]$Rights)

    $rightsValue = [long][int]$Rights

    # Named combinations checked from most- to least-specific.
    # FIX: Each key is unique  -  the original had 1245631 ("Modify") listed twice
    # which caused "Item has already been added" InvalidOperation at runtime.
    switch ($rightsValue) {
        2032127      { return "Full Control" }
        1245631      { return "Modify" }
        1180063      { return "Read & Execute" }
        1179817      { return "Read" }
        278          { return "Write" }
        { $_ -lt 0 } { return "Special / Inherited Flags ($Rights)" }
    }

    # Fall back: enumerate individual flag names that are set
    $flagNames = [System.Collections.Generic.List[string]]::new()
    $allFlags  = [System.Enum]::GetValues([System.Security.AccessControl.FileSystemRights])
    foreach ($flag in $allFlags) {
        $flagVal = [long][int]$flag
        # Skip zero-value and composite aggregates to avoid noise
        if ($flagVal -le 0) { continue }
        if (($rightsValue -band $flagVal) -eq $flagVal) {
            $name = $flag.ToString()
            if (-not $flagNames.Contains($name)) {
                $flagNames.Add($name)
            }
        }
    }

    if ($flagNames.Count -gt 0) {
        return $flagNames -join " | "
    }

    return $Rights.ToString()
}

function Get-FolderDepth {
    # Returns how many directory levels deep $Path is relative to $BasePath.
    param([string]$Path, [string]$BasePath)
    $base = $BasePath.TrimEnd("\")
    if ($Path.Length -le $base.Length) { return 0 }
    $rel  = $Path.Substring($base.Length)
    # Split on backslash; subtract 1 because leading "\" creates an empty first element
    return ($rel -split "\\").Count - 1
}

function Get-TruncatedDisplay {
    # FIX: Safely pad or truncate a string to exactly $Width chars for banner display.
    param([string]$Text, [int]$Width)
    if ($Text.Length -gt $Width) {
        return $Text.Substring(0, $Width - 3) + "..."
    }
    return $Text.PadRight($Width)
}

# -----------------------------------------------------------------------------
# REGION: Pre-scan Validation
# -----------------------------------------------------------------------------

if (-not (Test-Path $TargetDrive)) {
    Write-Host "[ERROR] Target drive/path '$TargetDrive' is not accessible." -ForegroundColor Red
    exit 1
}

# -----------------------------------------------------------------------------
# REGION: Banner
# FIX: All dynamic values are routed through Get-TruncatedDisplay so long paths
# cannot overflow the 64-char box border.
# -----------------------------------------------------------------------------

Write-Host ""
Write-Host "+==============================================================+" -ForegroundColor DarkCyan
Write-Host "|    FILE SHARE FOLDER PERMISSIONS SCANNER  v$ScriptVersion             |" -ForegroundColor DarkCyan
Write-Host "+==============================================================+" -ForegroundColor DarkCyan
Write-Host "|  Target Drive  : $(Get-TruncatedDisplay $TargetDrive 44)|" -ForegroundColor Cyan
Write-Host "|  Output CSV    : $(Get-TruncatedDisplay $OutputPath 44)|" -ForegroundColor Cyan
Write-Host "|  Log File      : $(Get-TruncatedDisplay $LogPath 44)|" -ForegroundColor Cyan
Write-Host "|  Batch Size    : $(Get-TruncatedDisplay "$BatchSize folders" 44)|" -ForegroundColor Cyan
Write-Host "|  Sleep (ms)    : $(Get-TruncatedDisplay "$SleepMs ms between batches" 44)|" -ForegroundColor Cyan
Write-Host "|  Skip Built-In : $(Get-TruncatedDisplay "$SkipBuiltIn" 44)|" -ForegroundColor Cyan
Write-Host "|  Incl. Inherit.: $(Get-TruncatedDisplay "$IncludeInherited" 44)|" -ForegroundColor Cyan
Write-Host "|  Skip Sys Dirs : $(Get-TruncatedDisplay "RECYCLE.BIN, System Volume Info, Recovery" 44)|" -ForegroundColor Cyan
Write-Host "+==============================================================+" -ForegroundColor DarkCyan
Write-Host ""

Write-Log "Scan started. Target: $TargetDrive | Output: $OutputPath" "INFO"
Set-ThrottledPriority

# -----------------------------------------------------------------------------
# REGION: CSV Header
# Written immediately so the file exists even if the scan is interrupted.
# -----------------------------------------------------------------------------

$csvHeader = "FolderPath,FolderDepth,Owner,IdentityReference,AccountType,AccessControlType," +
             "FileSystemRights,FriendlyRights,IsInherited,InheritanceFlags,PropagationFlags," +
             "ACLProtected,FolderCreated,FolderLastModified,ScanTimestamp"

Set-Content -Path $OutputPath -Value $csvHeader -Encoding UTF8
Write-Log "CSV initialised with headers: $OutputPath" "INFO"

# -----------------------------------------------------------------------------
# REGION: Folder Enumeration
# Uses .NET Directory API for speed. Hidden folders are included by default.
# Recycle Bin and system-reserved folders are filtered here at source.
# -----------------------------------------------------------------------------

Write-Log "Starting folder enumeration on: $TargetDrive" "INFO"
Write-Host "`n[*] Enumerating folders  -  this may take a moment on large shares...`n" -ForegroundColor Yellow

$allFolders          = [System.Collections.Generic.List[string]]::new()
$enumErrors          = 0
$scanErrors          = 0
$totalRecords        = 0
$processedFolders    = 0
$skippedSystemFolders = 0

# Include the drive root itself
$rootPath = $TargetDrive.TrimEnd("\")
$allFolders.Add($rootPath)

# BFS enumeration queue
$enumQueue = [System.Collections.Generic.Queue[string]]::new()
$enumQueue.Enqueue($rootPath)

while ($enumQueue.Count -gt 0) {
    $currentDir = $enumQueue.Dequeue()

    # Depth check (0 = unlimited)
    if ($MaxDepth -gt 0) {
        $depth = Get-FolderDepth -Path $currentDir -BasePath $TargetDrive
        if ($depth -ge $MaxDepth) { continue }
    }

    try {
        $subDirs = [System.IO.Directory]::GetDirectories($currentDir)
        foreach ($sub in $subDirs) {

            # FIX: Skip Recycle Bin and Windows system-reserved folders.
            # All other hidden folders are intentionally preserved.
            if (Test-ShouldExcludeFolder -FolderPath $sub) {
                Write-Log "EXCLUDED (system/recycle): $sub" "INFO"
                $skippedSystemFolders++
                continue
            }

            $allFolders.Add($sub)
            $enumQueue.Enqueue($sub)
        }
    }
    catch [System.UnauthorizedAccessException] {
        Write-Log "ACCESS DENIED (enumeration skipped): $currentDir" "WARN"
        $enumErrors++
    }
    catch {
        Write-Log "Enumeration error on '$currentDir': $_" "WARN"
        $enumErrors++
    }
}

$totalFolders = $allFolders.Count
Write-Log "Enumeration complete. Folders to scan: $totalFolders | System folders skipped: $skippedSystemFolders | Errors: $enumErrors" "INFO"
Write-Host "[OK] $totalFolders folders queued for ACL scan ($skippedSystemFolders system/recycle folders excluded).`n" -ForegroundColor Green

# -----------------------------------------------------------------------------
# REGION: ACL Scan Loop
# -----------------------------------------------------------------------------

$batchBuffer = [System.Collections.Generic.List[string]]::new()
$batchCount  = 0
$scanTs      = $ScriptStart.ToString("yyyy-MM-dd HH:mm:ss")

foreach ($folder in $allFolders) {

    $processedFolders++
    $pct = [math]::Round(($processedFolders / $totalFolders) * 100, 1)

    # Update progress bar every 10 folders to minimise console I/O overhead
    if (($processedFolders % 10 -eq 0) -or ($processedFolders -eq $totalFolders)) {
        Write-Progress `
            -Activity  "Scanning Folder Permissions on $TargetDrive" `
            -Status    "[$processedFolders / $totalFolders]  $pct%   -   $folder" `
            -PercentComplete $pct
    }

    try {
        # Retrieve folder metadata (-Force ensures hidden items are accessible)
        $folderItem     = Get-Item -LiteralPath $folder -Force -ErrorAction Stop
        $folderCreated  = $folderItem.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
        $folderModified = $folderItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
        $folderDepth    = Get-FolderDepth -Path $folder -BasePath $TargetDrive

        # Read ACL
        $acl         = Get-Acl -LiteralPath $folder -ErrorAction Stop
        $owner       = if ($acl.Owner) { $acl.Owner } else { "(Unknown)" }
        $isProtected = $acl.AreAccessRulesProtected   # True = inheritance broken

        $accessRules = $acl.Access

        if ($accessRules.Count -eq 0) {
            # Record folders with no ACEs so they appear in the report
            # Column count: 15  -  FolderPath(0), Depth(1), Owner(2), Identity(3),
            # AccountType(4), ACType(5), FSRights(6), Friendly(7), IsInherited(8),
            # InhFlags(9), PropFlags(10), ACLProtected(11), Created(12), Modified(13), TS(14)
            # After "(No ACEs)" at col 3, we need 8 empty cols (4-11) = 9 commas.
            $safePath  = $folder -replace '"', '""'
            $safeOwner = $owner  -replace '"', '""'
            $row = '"{0}",{1},"{2}","(No ACEs)",,,,,,,,,"{3}","{4}","{5}"' -f `
                   $safePath, $folderDepth, $safeOwner,
                   $folderCreated, $folderModified, $scanTs
            $batchBuffer.Add($row)
            $totalRecords++
        }
        else {
            foreach ($rule in $accessRules) {

                # Skip inherited entries unless explicitly requested
                if ($rule.IsInherited -and -not $IncludeInherited) { continue }

                $identity = $rule.IdentityReference.ToString()

                # Optionally omit built-in / system accounts
                if ($SkipBuiltIn -and ($BuiltInAccounts -contains $identity)) { continue }

                $accountType    = Resolve-AccountType -IdentityReference $identity
                $friendlyRights = Get-FriendlyRights  -Rights $rule.FileSystemRights

                # Escape double-quotes for valid RFC 4180 CSV
                $safePath      = $folder                             -replace '"', '""'
                $safeOwner     = $owner                              -replace '"', '""'
                $safeIdentity  = $identity                           -replace '"', '""'
                $safeRights    = $rule.FileSystemRights.ToString()   -replace '"', '""'
                $safeFriendly  = $friendlyRights                     -replace '"', '""'
                $safeInhFlags  = $rule.InheritanceFlags.ToString()   -replace '"', '""'
                $safePropFlags = $rule.PropagationFlags.ToString()   -replace '"', '""'

                $row = '"{0}",{1},"{2}","{3}","{4}","{5}","{6}","{7}",{8},"{9}","{10}",{11},"{12}","{13}","{14}"' -f `
                       $safePath,
                       $folderDepth,
                       $safeOwner,
                       $safeIdentity,
                       $accountType,
                       $rule.AccessControlType,
                       $safeRights,
                       $safeFriendly,
                       $rule.IsInherited.ToString(),
                       $safeInhFlags,
                       $safePropFlags,
                       $isProtected.ToString(),
                       $folderCreated,
                       $folderModified,
                       $scanTs

                $batchBuffer.Add($row)
                $totalRecords++
            }
        }
    }
    catch [System.UnauthorizedAccessException] {
        Write-Log "ACCESS DENIED (ACL read skipped): $folder" "WARN"
        $scanErrors++
    }
    catch {
        Write-Log "ACL scan error on '$folder': $_" "WARN"
        $scanErrors++
    }

    # -- Batch flush + throttle pause ------------------------------------------
    $batchCount++
    if ($batchCount -ge $BatchSize) {
        $batchBuffer | Add-Content -Path $OutputPath -Encoding UTF8
        $batchBuffer.Clear()
        $batchCount = 0
        if ($SleepMs -gt 0) {
            Start-Sleep -Milliseconds $SleepMs
        }
    }
}

# Flush any remaining buffered rows
if ($batchBuffer.Count -gt 0) {
    $batchBuffer | Add-Content -Path $OutputPath -Encoding UTF8
    $batchBuffer.Clear()
}

Write-Progress -Activity "Scanning Folder Permissions" -Completed

# -----------------------------------------------------------------------------
# REGION: Completion Summary
# -----------------------------------------------------------------------------

$ScriptEnd   = Get-Date
$Duration    = $ScriptEnd - $ScriptStart
$durationStr = "{0}h {1}m {2}s" -f $Duration.Hours, $Duration.Minutes, $Duration.Seconds

Write-Host ""
Write-Host "+==============================================================+" -ForegroundColor DarkGreen
Write-Host "|                      SCAN COMPLETE                          |" -ForegroundColor DarkGreen
Write-Host "+==============================================================+" -ForegroundColor DarkGreen
Write-Host "|  Folders Scanned    : $(Get-TruncatedDisplay $totalFolders.ToString() 41)|" -ForegroundColor Green
Write-Host "|  ACL Records Written: $(Get-TruncatedDisplay $totalRecords.ToString() 41)|" -ForegroundColor Green
Write-Host "|  System Fldrs Skipped:$(Get-TruncatedDisplay $skippedSystemFolders.ToString() 41)|" -ForegroundColor Cyan
Write-Host "|  Enum Errors        : $(Get-TruncatedDisplay $enumErrors.ToString() 41)|" -ForegroundColor Yellow
Write-Host "|  Scan Errors        : $(Get-TruncatedDisplay $scanErrors.ToString() 41)|" -ForegroundColor Yellow
Write-Host "|  Duration           : $(Get-TruncatedDisplay $durationStr 41)|" -ForegroundColor Green
Write-Host "|  Output CSV         : $(Get-TruncatedDisplay $OutputPath 41)|" -ForegroundColor Green
Write-Host "|  Log File           : $(Get-TruncatedDisplay $LogPath 41)|" -ForegroundColor Green
Write-Host "+==============================================================+" -ForegroundColor DarkGreen
Write-Host ""

Write-Log "Scan complete. Folders: $totalFolders | Records: $totalRecords | SystemSkipped: $skippedSystemFolders | Errors: $scanErrors | Duration: $durationStr" "SUCCESS"
Write-Log "Output CSV : $OutputPath" "SUCCESS"
Write-Log "Output Log : $LogPath"   "SUCCESS"