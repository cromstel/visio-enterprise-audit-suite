#Requires -Version 5.1

<#
.SYNOPSIS
    Scans all folders on a specified drive/share and exports ACL (permission) data to CSV.

.DESCRIPTION
    Recursively enumerates all folders on the target drive (default: K:\),
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
    Root path to scan. Default: K:\

.PARAMETER OutputPath
    Full path (including filename) for the CSV report.

.PARAMETER BatchSize
    Number of folders to process before pausing. Default: 50

.PARAMETER SleepMs
    Milliseconds to sleep between batches. Default: 200

.PARAMETER MaxDepth
    Maximum folder depth to recurse. 0 = unlimited. Default: 0

.PARAMETER SkipBuiltIn
    If set, skips built-in system accounts.

.PARAMETER IncludeInherited
    If set, includes inherited permission entries.

.NOTES
    Version     : 3.1
    Requires    : PowerShell 5.1+
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDrive = "K:\",

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

# Enable long path support where possible
try {
    [System.AppContext]::SetSwitch("Switch.System.IO.UseLegacyPathHandling", $false)
    [System.AppContext]::SetSwitch("Switch.System.IO.BlockLongPaths", $false)
}
catch { }

# ─────────────────────────────────────────────────────────────────────────────
# Resolve Script Directory
# ─────────────────────────────────────────────────────────────────────────────

$ScriptDir = if ($PSScriptRoot -and $PSScriptRoot -ne "") {
    $PSScriptRoot
}
elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}
else {
    (Get-Location).Path
}

# ─────────────────────────────────────────────────────────────────────────────
# Setup
# ─────────────────────────────────────────────────────────────────────────────

$ScriptVersion = "3.1"
$ScriptStart   = Get-Date
$Timestamp     = $ScriptStart.ToString("yyyyMMdd_HHmmss")

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $reportSubFolder = Join-Path $ScriptDir "ScanReports\$Timestamp"
    $OutputPath      = Join-Path $reportSubFolder "FolderPermissions_$Timestamp.csv"
}

$LogPath = [System.IO.Path]::ChangeExtension($OutputPath, ".log")

# ─────────────────────────────────────────────────────────────────────────────
# Output Directory
# ─────────────────────────────────────────────────────────────────────────────

$outputDir = Split-Path -Parent $OutputPath

if (-not [string]::IsNullOrWhiteSpace($outputDir) -and -not (Test-Path $outputDir)) {
    try {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        Write-Host "[INFO] Created output directory: $outputDir" -ForegroundColor Cyan
    }
    catch {
        Write-Host "[ERROR] Cannot create output directory '$outputDir': $_" -ForegroundColor Red
        exit 1
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Built-in Accounts
# ─────────────────────────────────────────────────────────────────────────────

$BuiltInAccounts = @(
    "NT AUTHORITY\SYSTEM",
    "NT SERVICE\TrustedInstaller",
    "BUILTIN\Administrators",
    "CREATOR OWNER",
    "NT AUTHORITY\Authenticated Users",
    "Everyone"
)

# ─────────────────────────────────────────────────────────────────────────────
# Excluded Folder Names
# ─────────────────────────────────────────────────────────────────────────────

$ExcludedFolderNames = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

@(
    '$RECYCLE.BIN',
    '$Recycle.Bin',
    'RECYCLER',
    'RECYCLED',
    'System Volume Information',
    'Recovery'
) | ForEach-Object {
    [void]$ExcludedFolderNames.Add($_)
}

# ─────────────────────────────────────────────────────────────────────────────
# Helper Functions
# ─────────────────────────────────────────────────────────────────────────────

function Write-Log {

    param(
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $entry = "[{0}] [{1}] {2}" -f (
        Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    ), $Level, $Message

    try {
        Add-Content -Path $LogPath -Value $entry -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch { }

    switch ($Level) {

        "ERROR" {
            Write-Host $entry -ForegroundColor Red
        }

        "WARN" {
            Write-Host $entry -ForegroundColor Yellow
        }

        "SUCCESS" {
            Write-Host $entry -ForegroundColor Green
        }

        default {
            Write-Host $entry -ForegroundColor Cyan
        }
    }
}

function Set-ThrottledPriority {

    try {
        $proc = [System.Diagnostics.Process]::GetCurrentProcess()
        $proc.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::BelowNormal

        Write-Log "Process priority set to BelowNormal." "INFO"
    }
    catch {
        Write-Log "Could not set process priority: $_" "WARN"
    }
}

function Test-ShouldExcludeFolder {

    param(
        [string]$FolderPath
    )

    $folderName = [System.IO.Path]::GetFileName($FolderPath)

    return $ExcludedFolderNames.Contains($folderName)
}

function Resolve-AccountType {

    param(
        [string]$IdentityReference
    )

    if ($IdentityReference -match "^S-1-\d") {
        return "SID (Unresolved)"
    }

    if ($IdentityReference -match "^(BUILTIN|NT AUTHORITY|NT SERVICE)\\") {
        return "Built-In / System"
    }

    $bare = ($IdentityReference -split "\\")[-1]

    try {
        $localGroup = [ADSI]"WinNT://./$bare,group"

        if ($localGroup.SchemaClassName -eq "Group") {
            return "Local Group"
        }
    }
    catch { }

    try {
        $localUser = [ADSI]"WinNT://./$bare,user"

        if ($localUser.SchemaClassName -eq "User") {
            return "Local User"
        }
    }
    catch { }

    if ($bare -match "(Group|Grp|Team|Dept|Admin|Users|Operators|Access|Members|Staff|Read|Write|Modify)$") {
        return "Domain Group (Inferred)"
    }

    return "Domain User / Group"
}

function Get-FriendlyRights {

    param(
        [System.Security.AccessControl.FileSystemRights]$Rights
    )

    $rightsValue = [long][int]$Rights

    switch ($rightsValue) {

        2032127 { return "Full Control" }
        1245631 { return "Modify" }
        1180063 { return "Read & Execute" }
        1179817 { return "Read" }
        278     { return "Write" }

        { $_ -lt 0 } {
            return "Special / Inherited Flags ($Rights)"
        }
    }

    $flagNames = [System.Collections.Generic.List[string]]::new()

    $allFlags = [System.Enum]::GetValues(
        [System.Security.AccessControl.FileSystemRights]
    ) |
    Sort-Object { [int]$_ } -Descending |
    Select-Object -Unique

    foreach ($flag in $allFlags) {

        $flagVal = [long][int]$flag

        if ($flagVal -le 0) {
            continue
        }

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

    param(
        [string]$Path,
        [string]$BasePath
    )

    $base = $BasePath.TrimEnd("\")

    if ($Path.Length -le $base.Length) {
        return 0
    }

    $rel = $Path.Substring($base.Length)

    return ($rel -split "\\").Count - 1
}

function Get-TruncatedDisplay {

    param(
        [AllowNull()]
        [string]$Text,

        [int]$Width
    )

    if ($null -eq $Text) {
        $Text = ""
    }

    if ($Text.Length -gt $Width) {
        return $Text.Substring(0, $Width - 3) + "..."
    }

    return $Text.PadRight($Width)
}

function Write-BannerLine {

    param(
        [string]$Label,
        [string]$Value,

        [ConsoleColor]$Color = "Cyan"
    )

    $content = "{0,-16}: {1}" -f $Label, $Value
    $display = Get-TruncatedDisplay $content 60

    Write-Host "║ $display ║" -ForegroundColor $Color
}

# ─────────────────────────────────────────────────────────────────────────────
# Validation
# ─────────────────────────────────────────────────────────────────────────────

if (-not (Test-Path $TargetDrive)) {

    Write-Host "[ERROR] Target drive/path '$TargetDrive' is not accessible." -ForegroundColor Red
    exit 1
}

# Normalize root path
$rootPath = [System.IO.Path]::GetFullPath($TargetDrive).TrimEnd("\")

# ─────────────────────────────────────────────────────────────────────────────
# Banner
# ─────────────────────────────────────────────────────────────────────────────

Write-Host ""

Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor DarkCyan
Write-Host "║    FILE SHARE FOLDER PERMISSIONS SCANNER  v$ScriptVersion             ║" -ForegroundColor DarkCyan
Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor DarkCyan

Write-BannerLine "Target Drive" $TargetDrive
Write-BannerLine "Output CSV" $OutputPath
Write-BannerLine "Log File" $LogPath
Write-BannerLine "Batch Size" "$BatchSize folders"
Write-BannerLine "Sleep (ms)" "$SleepMs ms between batches"
Write-BannerLine "Skip Built-In" $SkipBuiltIn
Write-BannerLine "Incl. Inherit." $IncludeInherited
Write-BannerLine "Skip Sys Dirs" "RECYCLE.BIN, System Volume Info, Recovery"

Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor DarkCyan
Write-Host ""

Write-Log "Scan started. Target: $TargetDrive | Output: $OutputPath" "INFO"

Set-ThrottledPriority

# ─────────────────────────────────────────────────────────────────────────────
# CSV Header
# ─────────────────────────────────────────────────────────────────────────────

$csvHeader = @(
    "FolderPath",
    "FolderDepth",
    "Owner",
    "IdentityReference",
    "AccountType",
    "AccessControlType",
    "FileSystemRights",
    "FriendlyRights",
    "IsInherited",
    "InheritanceFlags",
    "PropagationFlags",
    "ACLProtected",
    "FolderCreated",
    "FolderLastModified",
    "ScanTimestamp"
) -join ","

Set-Content -Path $OutputPath -Value $csvHeader -Encoding UTF8

Write-Log "CSV initialized: $OutputPath" "INFO"

# ─────────────────────────────────────────────────────────────────────────────
# Folder Enumeration
# ─────────────────────────────────────────────────────────────────────────────

Write-Log "Starting folder enumeration on: $TargetDrive" "INFO"

Write-Host "`n[*] Enumerating folders — this may take a while...`n" -ForegroundColor Yellow

$allFolders           = [System.Collections.Generic.List[string]]::new()
$enumErrors           = 0
$scanErrors           = 0
$totalRecords         = 0
$processedFolders     = 0
$skippedSystemFolders = 0

$allFolders.Add($rootPath)

$enumQueue = [System.Collections.Generic.Queue[string]]::new()
$enumQueue.Enqueue($rootPath)

while ($enumQueue.Count -gt 0) {

    $currentDir = $enumQueue.Dequeue()

    if ($MaxDepth -gt 0) {

        $depth = Get-FolderDepth -Path $currentDir -BasePath $rootPath

        if ($depth -ge $MaxDepth) {
            continue
        }
    }

    try {

        $subDirs = [System.IO.Directory]::GetDirectories($currentDir)

        foreach ($sub in $subDirs) {

            if (Test-ShouldExcludeFolder -FolderPath $sub) {

                Write-Log "EXCLUDED (system/recycle): $sub" "INFO"

                $skippedSystemFolders++

                continue
            }

            $dirInfo = Get-Item -LiteralPath $sub -Force -ErrorAction SilentlyContinue

            if (
                $dirInfo -and
                ($dirInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint)
            ) {

                Write-Log "SKIPPED REPARSE POINT: $sub" "INFO"

                continue
            }

            $allFolders.Add($sub)
            $enumQueue.Enqueue($sub)
        }
    }
    catch [System.UnauthorizedAccessException] {

        Write-Log "ACCESS DENIED: $currentDir" "WARN"
        $enumErrors++
    }
    catch [System.IO.PathTooLongException] {

        Write-Log "PATH TOO LONG: $currentDir" "WARN"
        $enumErrors++
    }
    catch [System.IO.IOException] {

        Write-Log "I/O ERROR: $currentDir : $_" "WARN"
        $enumErrors++
    }
    catch {

        Write-Log "Enumeration error on '$currentDir': $_" "WARN"
        $enumErrors++
    }
}

$totalFolders = $allFolders.Count

Write-Log "Enumeration complete. Folders: $totalFolders" "INFO"

Write-Host "[OK] $totalFolders folders queued.`n" -ForegroundColor Green

# ─────────────────────────────────────────────────────────────────────────────
# ACL Scan
# ─────────────────────────────────────────────────────────────────────────────

$batchBuffer = [System.Collections.Generic.List[string]]::new()
$batchCount  = 0
$scanTs      = $ScriptStart.ToString("yyyy-MM-dd HH:mm:ss")

foreach ($folder in $allFolders) {

    $processedFolders++

    if ($totalFolders -gt 0) {
        $pct = [math]::Round(($processedFolders / $totalFolders) * 100, 1)
    }
    else {
        $pct = 0
    }

    if (
        ($processedFolders % 10 -eq 0) -or
        ($processedFolders -eq $totalFolders)
    ) {

        Write-Progress `
            -Activity "Scanning Folder Permissions on $TargetDrive" `
            -Status "[$processedFolders / $totalFolders]  $pct%  —  $folder" `
            -PercentComplete $pct
    }

    try {

        $folderItem = Get-Item -LiteralPath $folder -Force -ErrorAction Stop

        $folderCreated  = $folderItem.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
        $folderModified = $folderItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")

        $folderDepth = Get-FolderDepth -Path $folder -BasePath $rootPath

        $acl = Get-Acl -LiteralPath $folder -ErrorAction Stop

        if ($null -eq $acl) {
            throw "ACL object is null."
        }

        $owner       = if ($acl.Owner) { $acl.Owner } else { "(Unknown)" }
        $isProtected = $acl.AreAccessRulesProtected

        $accessRules = @($acl.Access)

        if ($accessRules.Count -eq 0) {

            $safePath  = $folder -replace '"', '""'
            $safeOwner = $owner -replace '"', '""'

            $row = '"{0}",{1},"{2}","(No ACEs)","","","","","","","","","{3}","{4}","{5}"' -f `
                $safePath,
                $folderDepth,
                $safeOwner,
                $folderCreated,
                $folderModified,
                $scanTs

            $batchBuffer.Add($row)

            $totalRecords++
        }
        else {

            foreach ($rule in $accessRules) {

                if ($rule.IsInherited -and -not $IncludeInherited) {
                    continue
                }

                $identity = $rule.IdentityReference.ToString()

                if (
                    $SkipBuiltIn -and
                    ($BuiltInAccounts -contains $identity)
                ) {
                    continue
                }

                $accountType = Resolve-AccountType -IdentityReference $identity

                $friendlyRights = Get-FriendlyRights `
                    -Rights $rule.FileSystemRights

                $safePath      = $folder -replace '"', '""'
                $safeOwner     = $owner -replace '"', '""'
                $safeIdentity  = $identity -replace '"', '""'
                $safeRights    = $rule.FileSystemRights.ToString() -replace '"', '""'
                $safeFriendly  = $friendlyRights -replace '"', '""'
                $safeInhFlags  = $rule.InheritanceFlags.ToString() -replace '"', '""'
                $safePropFlags = $rule.PropagationFlags.ToString() -replace '"', '""'

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

    $batchCount++

    if ($batchCount -ge $BatchSize) {

        [System.IO.File]::AppendAllLines($OutputPath, $batchBuffer)

        $batchBuffer.Clear()

        $batchCount = 0

        if ($SleepMs -gt 0) {
            Start-Sleep -Milliseconds $SleepMs
        }
    }
}

if ($batchBuffer.Count -gt 0) {

    [System.IO.File]::AppendAllLines($OutputPath, $batchBuffer)

    $batchBuffer.Clear()
}

Write-Progress -Activity "Scanning Folder Permissions" -Completed

# ─────────────────────────────────────────────────────────────────────────────
# Completion Summary
# ─────────────────────────────────────────────────────────────────────────────

$ScriptEnd = Get-Date
$Duration  = $ScriptEnd - $ScriptStart

$durationStr = "{0}h {1}m {2}s" -f `
    $Duration.Hours,
    $Duration.Minutes,
    $Duration.Seconds

Write-Host ""

Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor DarkGreen
Write-Host "║                      SCAN COMPLETE                          ║" -ForegroundColor DarkGreen
Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor DarkGreen

Write-BannerLine "Folders Scanned" $totalFolders Green
Write-BannerLine "ACL Records" $totalRecords Green
Write-BannerLine "System Skipped" $skippedSystemFolders Cyan
Write-BannerLine "Enum Errors" $enumErrors Yellow
Write-BannerLine "Scan Errors" $scanErrors Yellow
Write-BannerLine "Duration" $durationStr Green
Write-BannerLine "Output CSV" $OutputPath Green
Write-BannerLine "Log File" $LogPath Green

Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor DarkGreen

Write-Host ""

Write-Log "Scan complete. Folders: $totalFolders | Records: $totalRecords | Errors: $scanErrors | Duration: $durationStr" "SUCCESS"

Write-Log "Output CSV : $OutputPath" "SUCCESS"
Write-Log "Output Log : $LogPath" "SUCCESS"
