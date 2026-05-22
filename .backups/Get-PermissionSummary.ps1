#Requires -Version 5.1
<#
.SYNOPSIS
    Reads a FolderPermissions CSV and generates condensed summary reports.

.DESCRIPTION
    After running Scan-FolderPermissions.ps1, use this script to produce:
      1. By-Group/User    : every identity ranked by folder count, with rights summary
      2. Deny ACEs        : all explicit Deny permission entries
      3. Broken Inheritance: folders where ACL inheritance has been manually broken

    Summary CSVs are saved in a timestamped subfolder next to this script file
    (or in a folder you specify via -OutputFolder).

.PARAMETER InputCsv
    Path to the CSV produced by Scan-FolderPermissions.ps1. (Required)

.PARAMETER OutputFolder
    Where to write the three summary reports.
    Default: <ScriptDir>\ScanReports\Summaries_<timestamp>\

.EXAMPLE
    .\Get-PermissionSummary.ps1 -InputCsv ".\ScanReports\20240915_020031\FolderPermissions_Balanced_20240915_020031.csv"

.EXAMPLE
    .\Get-PermissionSummary.ps1 -InputCsv ".\ScanReports\20240915_020031\FolderPermissions_Balanced_20240915_020031.csv" -OutputFolder "D:\AuditReports"

.NOTES
    Version 3.0: Fixed Format-Table Substring/PadRight truncation bug;
                 OutputFolder now defaults to script location not InputCsv parent;
                 added OutputFolder auto-creation; fixed null/empty collection guards.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "Input CSV not found: $_"
        }
        return $true
    })]
    [string]$InputCsv,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Resolve Script Directory
# FIX: Guard against $PSScriptRoot being empty when dot-sourced.
# ─────────────────────────────────────────────────────────────────────────────

$ScriptDir = if ($PSScriptRoot -and $PSScriptRoot -ne "") {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    (Get-Location).Path
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Output Folder
# FIX: Default is now relative to the script file, not the parent of InputCsv.
# Auto-creates the folder if it does not exist.
# ─────────────────────────────────────────────────────────────────────────────

$ts = Get-Date -Format "yyyyMMdd_HHmmss"

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = Join-Path $ScriptDir "ScanReports\Summaries_$ts"
}

if (-not (Test-Path $OutputFolder)) {
    try {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
        Write-Host "[INFO] Created output folder: $OutputFolder" -ForegroundColor Cyan
    }
    catch {
        Write-Host "[ERROR] Cannot create output folder '$OutputFolder': $_" -ForegroundColor Red
        exit 1
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Output File Paths
# ─────────────────────────────────────────────────────────────────────────────

$baseName          = [System.IO.Path]::GetFileNameWithoutExtension($InputCsv)
$summaryByGroupCsv = Join-Path $OutputFolder "${baseName}_Summary_ByGroup_$ts.csv"
$summaryDenyCsv    = Join-Path $OutputFolder "${baseName}_Summary_DenyACEs_$ts.csv"
$summaryBrokenCsv  = Join-Path $OutputFolder "${baseName}_Summary_BrokenInheritance_$ts.csv"

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Banner
# ─────────────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor DarkYellow
Write-Host "║       PERMISSION SUMMARY ANALYSER  v3.0                     ║" -ForegroundColor DarkYellow
Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor DarkYellow
$truncInput  = if ($InputCsv.Length  -gt 55) { "..." + $InputCsv.Substring($InputCsv.Length - 52)  } else { $InputCsv }
$truncOutput = if ($OutputFolder.Length -gt 55) { "..." + $OutputFolder.Substring($OutputFolder.Length - 52) } else { $OutputFolder }
Write-Host "║  Input : $($truncInput.PadRight(54))║" -ForegroundColor Yellow
Write-Host "║  Output: $($truncOutput.PadRight(54))║" -ForegroundColor Yellow
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor DarkYellow
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Load CSV
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[*] Loading CSV: $InputCsv" -ForegroundColor Cyan

$data = Import-Csv -Path $InputCsv -Encoding UTF8

if (-not $data -or $data.Count -eq 0) {
    Write-Host "[ERROR] CSV is empty or could not be parsed: $InputCsv" -ForegroundColor Red
    exit 1
}

$totalRows     = $data.Count
$uniqueFolders = ($data |
                  Select-Object -ExpandProperty FolderPath |
                  Sort-Object -Unique).Count
$uniqueIdents  = ($data |
                  Where-Object { $_.IdentityReference -ne "(No ACEs)" } |
                  Select-Object -ExpandProperty IdentityReference |
                  Sort-Object -Unique).Count

Write-Host "[OK] Loaded $totalRows rows | $uniqueFolders unique folders | $uniqueIdents unique identities" -ForegroundColor Green
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Report 1 — Summary by Group / User
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[*] Building: Summary by Group/User..." -ForegroundColor Cyan

$byGroup = $data |
    Where-Object { $_.IdentityReference -ne "(No ACEs)" } |
    Group-Object -Property IdentityReference |
    ForEach-Object {
        $ident   = $_.Name
        $entries = $_.Group

        # Unique access control types for this identity (Allow / Deny / both)
        $acTypes = ($entries |
                    Select-Object -ExpandProperty AccessControlType |
                    Sort-Object -Unique) -join " | "

        # Unique friendly rights labels
        $rights  = ($entries |
                    Select-Object -ExpandProperty FriendlyRights |
                    Sort-Object -Unique) -join " | "

        # Number of distinct folders this identity appears on
        $fCount  = ($entries |
                    Select-Object -ExpandProperty FolderPath |
                    Sort-Object -Unique).Count

        # Account type (first unique value; should be consistent per identity)
        $acType  = ($entries |
                    Select-Object -ExpandProperty AccountType |
                    Sort-Object -Unique) -join " | "

        # True if any rule for this identity is a Deny
        $hasDeny = ($entries |
                    Where-Object { $_.AccessControlType -eq "Deny" } |
                    Measure-Object).Count -gt 0

        # Up to 3 sample paths for quick reference
        $samplePaths = ($entries |
                        Select-Object -ExpandProperty FolderPath |
                        Sort-Object -Unique |
                        Select-Object -First 3) -join " ; "

        [PSCustomObject]@{
            IdentityReference  = $ident
            AccountType        = $acType
            TotalFolderCount   = $fCount
            AccessControlTypes = $acTypes
            RightsAssigned     = $rights
            HasDenyACE         = $hasDeny
            SampleFolderPaths  = $samplePaths
        }
    } |
    Sort-Object -Property TotalFolderCount -Descending

$byGroup | Export-Csv -Path $summaryByGroupCsv -NoTypeInformation -Encoding UTF8
Write-Host "[OK] By-Group summary saved: $summaryByGroupCsv ($($byGroup.Count) identities)" -ForegroundColor Green

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Report 2 — Explicit Deny ACEs
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[*] Building: Deny ACE report..." -ForegroundColor Cyan

# FIX: Explicit Count check instead of implicit boolean on collection,
# which is unreliable in PowerShell when result is a single-item array.
$denyRows = @($data | Where-Object { $_.AccessControlType -eq "Deny" })

if ($denyRows.Count -gt 0) {
    $denyRows |
        Select-Object FolderPath, FolderDepth, Owner, IdentityReference,
                      AccountType, FileSystemRights, FriendlyRights,
                      IsInherited, ACLProtected, FolderCreated, FolderLastModified |
        Export-Csv -Path $summaryDenyCsv -NoTypeInformation -Encoding UTF8
    Write-Host "[WARN] Deny ACEs found — saved: $summaryDenyCsv ($($denyRows.Count) entries)" -ForegroundColor Yellow
}
else {
    Write-Host "[OK] No Deny ACEs found on the share." -ForegroundColor Green
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Report 3 — Folders with Broken Inheritance
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[*] Building: Broken inheritance report..." -ForegroundColor Cyan

# FIX: Use Group-Object to de-duplicate by FolderPath first, THEN select
# columns — avoids the unreliable Select-Object -Unique multi-property dedup.
$brokenFolders = @($data |
    Where-Object { $_.ACLProtected -eq "True" } |
    Group-Object -Property FolderPath |
    ForEach-Object {
        $first = $_.Group[0]
        # Collect all identities and rights for this folder
        $identities = ($_.Group |
                       Select-Object -ExpandProperty IdentityReference |
                       Sort-Object -Unique) -join " ; "
        $rights     = ($_.Group |
                       Select-Object -ExpandProperty FriendlyRights |
                       Sort-Object -Unique) -join " ; "

        [PSCustomObject]@{
            FolderPath         = $first.FolderPath
            FolderDepth        = $first.FolderDepth
            Owner              = $first.Owner
            AllIdentities      = $identities
            AllRightsAssigned  = $rights
            ExplicitACECount   = $_.Group.Count
            FolderCreated      = $first.FolderCreated
            FolderLastModified = $first.FolderLastModified
        }
    } |
    Sort-Object -Property FolderPath)

if ($brokenFolders.Count -gt 0) {
    $brokenFolders | Export-Csv -Path $summaryBrokenCsv -NoTypeInformation -Encoding UTF8
    Write-Host "[WARN] Broken inheritance found — saved: $summaryBrokenCsv ($($brokenFolders.Count) folders)" -ForegroundColor Yellow
}
else {
    Write-Host "[OK] No folders with broken inheritance found." -ForegroundColor Green
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Console Summary Table — Top 15 Identities
# FIX: Replaced PadRight(N).Substring(0, actualLength) anti-pattern which
# stripped the padding it had just applied. Correct form: Substring first
# to cap at N, then PadRight to guarantee minimum width.
# ─────────────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "================================================================" -ForegroundColor DarkGray
Write-Host " TOP 15 IDENTITIES BY FOLDER ACCESS COUNT" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor DarkGray

$byGroup | Select-Object -First 15 |
    Format-Table -AutoSize -Wrap:$false -Property `
        @{
            Label      = "Identity (truncated to 40)"
            Expression = {
                $s = $_.IdentityReference
                # FIX: Substring to cap, THEN PadRight to fill — correct order
                $s.Substring(0, [math]::Min(40, $s.Length)).PadRight(40)
            }
        },
        @{
            Label      = "Type"
            Expression = {
                $s = $_.AccountType
                $s.Substring(0, [math]::Min(22, $s.Length)).PadRight(22)
            }
        },
        @{
            Label      = "Folders"
            Expression = { $_.TotalFolderCount }
        },
        @{
            Label      = "Rights (truncated)"
            Expression = {
                $s = $_.RightsAssigned
                $s.Substring(0, [math]::Min(28, $s.Length)).PadRight(28)
            }
        },
        @{
            Label      = "HasDeny"
            Expression = { $_.HasDenyACE }
        }

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Final Output File Summary
# ─────────────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "================================================================" -ForegroundColor DarkGray
Write-Host " OUTPUT FILES" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor DarkGray
Write-Host "  By-Group Summary      : $summaryByGroupCsv" -ForegroundColor Cyan
if (Test-Path $summaryDenyCsv)   {
    Write-Host "  Deny ACEs Report      : $summaryDenyCsv" -ForegroundColor Yellow
}
if (Test-Path $summaryBrokenCsv) {
    Write-Host "  Broken Inheritance    : $summaryBrokenCsv" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "  All files are UTF-8 encoded and Excel-compatible." -ForegroundColor DarkGray
Write-Host "  Import via: Data > Get Data > From Text/CSV (delimiter: Comma)" -ForegroundColor DarkGray
Write-Host "================================================================" -ForegroundColor DarkGray
Write-Host ""
