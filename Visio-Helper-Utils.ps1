#requires -RunAsAdministrator
#requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Visio Audit Helper Utilities
    Common tasks and shortcuts for managing Visio audit data
    Supports: Visio 2021, 2019, 2016, and Office 365 (x64/x86)

.DESCRIPTION
    Provides quick access to:
    - Report generation and analysis
    - Email notifications
    - Excel exports
    - Automated scheduling
    - Data visualization
#>

# ============================================================================
# UTILITY CONFIGURATION
# ============================================================================
[PSCredential]$script:VisioScanCredential = $null

function Get-VisioScanCredential {
    [CmdletBinding()]
    param(
        [switch]$Force
    )

    if (-not $Force -and $script:VisioScanCredential) {
        return $script:VisioScanCredential
    }

    $useAlt = Read-Host "Use an alternate local admin credential for remote scans? (Y/N)"
    if ($useAlt -match '^[Yy]') {
        $script:VisioScanCredential = Get-Credential -Message "Enter the local admin credential used by the audit"
    }

    return $script:VisioScanCredential
}

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

function Show-Menu {
    Write-Host ("`n" + ("=" * 72)) -ForegroundColor Cyan
    Write-Host "VISIO ENTERPRISE AUDIT - HELPER UTILITIES" -ForegroundColor Cyan
    Write-Host "Supports: Visio 2021/2019/2016, Office 365, and local-admin credential prompts" -ForegroundColor Cyan
    Write-Host (("=" * 72)) -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1.  Run Full Installation Audit" -ForegroundColor Yellow
    Write-Host "2.  Run Usage Analytics" -ForegroundColor Yellow
    Write-Host "3.  Find Unused Visio Installations (6+ months)" -ForegroundColor Yellow
    Write-Host "4.  Export Latest Report to Excel" -ForegroundColor Yellow
    Write-Host "5.  New Cost Analysis" -ForegroundColor Yellow
    Write-Host "6.  View Last Report Summary" -ForegroundColor Yellow
    Write-Host "7.  Compare Two Reports (detect changes)" -ForegroundColor Yellow
    Write-Host "8.  Send Report via Email" -ForegroundColor Yellow
    Write-Host "9.  Create Scheduled Task" -ForegroundColor Yellow
    Write-Host "10. Select Report by Department" -ForegroundColor Yellow
    Write-Host "11. Generate Department Summary" -ForegroundColor Yellow
    Write-Host "12. Exit" -ForegroundColor Yellow
    Write-Host "13. Show Access Error 397 Guidance" -ForegroundColor Yellow
    Write-Host ""
}

function Invoke-FullAudit {
    param(
        [string]$OutputPath,
        [string]$ComputerPrefix = "GOT",
        [int]$Threads = 10,
        [PSCredential]$ScanCredential
    )
    
    Write-Host "`n[*] Starting full Visio audit..." -ForegroundColor Cyan
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    
    if ([string]::IsNullOrEmpty($OutputPath)) {
        $OutputPath = "$scriptPath\Output\VisioAudit"
    }

    $credential = if ($ScanCredential) { $ScanCredential } else { Get-VisioScanCredential }
    $auditArgs = @(
        "-OutputPath", $OutputPath,
        "-ComputerPrefix", $ComputerPrefix,
        "-ThreadCount", $Threads
    )
    if ($credential) {
        $auditArgs += "-ScanCredential"
        $auditArgs += $credential
    }

    & "$scriptPath\visio-enterprise-audit.ps1" @auditArgs

    Write-Host "[+] Audit complete! Check $OutputPath for reports." -ForegroundColor Green
}

function Invoke-UsageAnalytics {
    param(
        [PSCredential]$ScanCredential
    )

    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    Write-Host "`n[*] Starting usage analytics..." -ForegroundColor Cyan
    $credential = if ($ScanCredential) { $ScanCredential } else { Get-VisioScanCredential }
    $analyticsArgs = @()
    if ($credential) {
        $analyticsArgs += "-ScanCredential"
        $analyticsArgs += $credential
    }

    & "$scriptPath\visio-Usage-analytics.ps1" @analyticsArgs
    Write-Host "[+] Usage analytics complete." -ForegroundColor Green
}

function Show-AccessErrorGuidance {
    Write-Host "`n[*] Access Denied (CIM 397) Guidance" -ForegroundColor Cyan
    Write-Host "CIM 397 indicates the audit host could not establish CIM/WMI communication." -ForegroundColor Yellow
    Write-Host "Best options to work around it:" -ForegroundColor Green
    Write-Host "  1. Re-run the audit with -ScanCredential (local administrator) so CIM runs under elevated context." -ForegroundColor White
    Write-Host "  2. Enable WinRM/remote registry on the target hosts or make sure the firewall allows DCOM/CIM traffic." -ForegroundColor White
    Write-Host "  3. When remoting is blocked, deploy the Visio scripts locally (scheduling or remote execution tools) and centralize results later." -ForegroundColor White
    Write-Host "  4. Run Office-Version-Detector locally to prove the version and then rerun the audit with valid credentials." -ForegroundColor White
    Write-Host ""
}

function Find-UnusedVisio {
    param(
        [string]$ReportPath,
        [int]$MonthsInactive = 6
    )
    
    # Use dynamic script path
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    
    if ([string]::IsNullOrEmpty($ReportPath)) {
        $ReportPath = "$scriptPath\Output\VisioAudit"
    }

    Write-Host "`n[*] Finding Visio installations unused for $MonthsInactive+ months..." -ForegroundColor Cyan

    $cutoffDate = (Get-Date).AddMonths(-$MonthsInactive)
    $latestReport = Get-ChildItem -Path $ReportPath -Filter "VisioAudit_*.csv" | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

    if (!$latestReport) {
        Write-Host "[-] No reports found in $ReportPath" -ForegroundColor Red
        return
    }

    Write-Host "[+] Using report: $($latestReport.Name)" -ForegroundColor Green

    $report = Import-Csv $latestReport.FullName
    $unused = $report | Where-Object {
        $_.VisioInstalled -eq "Yes" -and
        ![string]::IsNullOrEmpty($_.LastUsedDate) -and
        [datetime]$_.LastUsedDate -lt $cutoffDate
    } | Sort-Object -Property LastUsedDate

    Write-Host "`n[+] Found $($unused.Count) unused Visio installations:" -ForegroundColor Green
    Write-Host ""

    $unused | Format-Table -Property ComputerName, VisioVersion, VisioEdition, LastUsedDate -AutoSize | Out-Host

    # Export to CSV
    $outputFile = Join-Path (Split-Path $latestReport.FullName) "UnusedVisio_$($MonthsInactive)Months_$(Get-Date -Format 'yyyyMMdd').csv"
    $unused | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Host "[+] Exported to: $outputFile" -ForegroundColor Green

    return $unused
}

function Export-ToExcel {
    param(
        [string]$ReportPath
    )
    
    # Use dynamic script path
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    
    if ([string]::IsNullOrEmpty($ReportPath)) {
        $ReportPath = "$scriptPath\Output\VisioAudit"
    }

    Write-Host "`n[*] Exporting to Excel..." -ForegroundColor Cyan

    # Check if ImportExcel module exists
    $moduleExists = Get-Module -ListAvailable -Name ImportExcel
    if (!$moduleExists) {
        Write-Host "[-] ImportExcel module not found. Installing..." -ForegroundColor Yellow
        Install-Module ImportExcel -Force -Scope CurrentUser
    }

    $latestReport = Get-ChildItem -Path $ReportPath -Filter "VisioAudit_*.csv" | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

    if (!$latestReport) {
        Write-Host "[-] No reports found" -ForegroundColor Red
        return
    }

    $data = Import-Csv $latestReport.FullName
    $excelFile = Join-Path (Split-Path $latestReport.FullName) "VisioAudit_$(Get-Date -Format 'yyyyMMdd').xlsx"

    $data | Export-Excel -Path $excelFile `
        -WorksheetName "Installations" `
        -AutoFilter `
        -FreezeTopRow `
        -TableStyle Light10 `
        -ChartType ColumnClustered `
        -ChartTitle "Visio Installation Summary"

    Write-Host "[+] Excel report created: $excelFile" -ForegroundColor Green
}

function Show-ReportSummary {
    param(
        [string]$ReportPath
    )
    
    # Use dynamic script path
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    
    if ([string]::IsNullOrEmpty($ReportPath)) {
        $ReportPath = "$scriptPath\Output\VisioAudit"
    }

    $latestReport = Get-ChildItem -Path $ReportPath -Filter "VisioAudit_*.csv" | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

    if (!$latestReport) {
        Write-Host "[-] No reports found" -ForegroundColor Red
        return
    }

    Write-Host ("`n" + ("=" * 72)) -ForegroundColor Cyan
    Write-Host "LATEST AUDIT REPORT SUMMARY" -ForegroundColor Cyan
    Write-Host (("=" * 72)) -ForegroundColor Cyan

    $data = Import-Csv $latestReport.FullName

    $summary = @{
        Total          = $data.Count
        Online         = ($data | Where-Object { $_.IsOnline -eq "Yes" }).Count
        Offline        = ($data | Where-Object { $_.IsOnline -eq "No" }).Count
        WithVisio      = ($data | Where-Object { $_.VisioInstalled -eq "Yes" }).Count
        Office365      = ($data | Where-Object { $_.Office365 -eq "Yes" }).Count
        Errors         = ($data | Where-Object { $_.Error -ne "None" -and ![string]::IsNullOrEmpty($_.Error) }).Count
    }

    $visioStandard = ($data | Where-Object { $_.VisioEdition -eq "Standard" }).Count
    $visioProfessional = ($data | Where-Object { $_.VisioEdition -eq "Professional" }).Count

    Write-Host "`nReport File: $($latestReport.Name)" -ForegroundColor Yellow
    Write-Host "Generated: $($latestReport.LastWriteTime)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Total Computers Scanned:    $($summary.Total)" -ForegroundColor Green
    Write-Host "  - Online:                 $($summary.Online)" -ForegroundColor Green
    Write-Host "  - Offline:                $($summary.Offline)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Visio Installations:        $($summary.WithVisio)" -ForegroundColor Green
    Write-Host "  - Standard Edition:       $visioStandard" -ForegroundColor Cyan
    Write-Host "  - Professional Edition:   $visioProfessional" -ForegroundColor Cyan
    Write-Host "  - Office 365:             $($summary.Office365)" -ForegroundColor Cyan
    Write-Host ""
    $installationRate = if ($summary.Total -gt 0) { (($summary.WithVisio / $summary.Total) * 100).ToString("F1") } else { "0.0" }
    Write-Host "Installation Rate:          $installationRate%" -ForegroundColor Green
    Write-Host "Access Errors:              $($summary.Errors)" -ForegroundColor Red
    Write-Host ""

    # Show top 10 computers with Visio
    Write-Host "Top 10 Computers with Visio:" -ForegroundColor Cyan
    $data | Where-Object { $_.VisioInstalled -eq "Yes" } | Select-Object -First 10 -Property ComputerName, VisioVersion, Office365, LastUsedDate | Format-Table
}

function Compare-Reports {
    # Use dynamic script path
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $reportPath = "$scriptPath\Output\VisioAudit"

    Write-Host "`n[*] Comparing reports to find changes..." -ForegroundColor Cyan

    $reports = Get-ChildItem -Path $reportPath -Filter "VisioAudit_*.csv" | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 2

    if ($reports.Count -lt 2) {
        Write-Host "[-] Need at least 2 reports to compare" -ForegroundColor Red
        return
    }

    $newer = Import-Csv $reports[0].FullName
    $older = Import-Csv $reports[1].FullName

    # Find new installations
    $newInstalls = @()
    foreach ($computer in $newer) {
        $oldRecord = $older | Where-Object { $_.ComputerName -eq $computer.ComputerName }
        if (!$oldRecord -or ($oldRecord.VisioInstalled -eq "No" -and $computer.VisioInstalled -eq "Yes")) {
            $newInstalls += $computer
        }
    }

    # Find removed installations
    $removed = @()
    foreach ($computer in $older) {
        $newRecord = $newer | Where-Object { $_.ComputerName -eq $computer.ComputerName }
        if (!$newRecord -or ($computer.VisioInstalled -eq "Yes" -and $newRecord.VisioInstalled -eq "No")) {
            $removed += $computer
        }
    }

    Write-Host ("`n" + ("=" * 72)) -ForegroundColor Cyan
    Write-Host "CHANGES DETECTED" -ForegroundColor Cyan
    Write-Host (("=" * 72)) -ForegroundColor Cyan

    Write-Host "`nNew Installations: $($newInstalls.Count)" -ForegroundColor Green
    if ($newInstalls.Count -gt 0) {
        $newInstalls | Format-Table -Property ComputerName, VisioVersion, VisioEdition, Office365
    }

    Write-Host "`nRemoved Installations: $($removed.Count)" -ForegroundColor Yellow
    if ($removed.Count -gt 0) {
        $removed | Format-Table -Property ComputerName, VisioVersion, VisioEdition
    }
}

function New-CostAnalysis {
    param(
        [string]$ReportPath,
        [double]$Office365CostPerMonth = 60,
        [double]$DesktopCostPerUser = 300
    )
    
    # Use dynamic script path
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    
    if ([string]::IsNullOrEmpty($ReportPath)) {
        $ReportPath = "$scriptPath\Output\VisioAudit"
    }

    $latestReport = Get-ChildItem -Path $ReportPath -Filter "VisioAudit_*.csv" | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

    if (!$latestReport) {
        Write-Host "[-] No reports found" -ForegroundColor Red
        return
    }

    $data = Import-Csv $latestReport.FullName

    $visioStandard = ($data | Where-Object { $_.VisioEdition -eq "Standard" -and $_.IsOnline -eq "Yes" }).Count
    $visioProfessional = ($data | Where-Object { $_.VisioEdition -eq "Professional" -and $_.IsOnline -eq "Yes" }).Count
    $office365Count = ($data | Where-Object { $_.Office365 -eq "Yes" -and $_.IsOnline -eq "Yes" }).Count
    $desktopStandard = ($data | Where-Object { $_.VisioInstalled -eq "Yes" -and $_.Office365 -ne "Yes" -and $_.VisioEdition -eq "Standard" -and $_.IsOnline -eq "Yes" }).Count
    $desktopProfessional = ($data | Where-Object { $_.VisioInstalled -eq "Yes" -and $_.Office365 -ne "Yes" -and $_.VisioEdition -eq "Professional" -and $_.IsOnline -eq "Yes" }).Count

    $office365Annual = $office365Count * $Office365CostPerMonth * 12
    $desktopAnnual = ($desktopStandard + $desktopProfessional) * $DesktopCostPerUser
    $totalAnnual = $office365Annual + $desktopAnnual

    Write-Host ("`n" + ("=" * 72)) -ForegroundColor Cyan
    Write-Host "VISIO LICENSE COST ANALYSIS" -ForegroundColor Cyan
    Write-Host (("=" * 72)) -ForegroundColor Cyan

    Write-Host "`nLicense Summary:" -ForegroundColor Green
    Write-Host "  Total Standard Edition:     $visioStandard" -ForegroundColor Cyan
    Write-Host "  Total Professional Edition: $visioProfessional" -ForegroundColor Cyan
    Write-Host "  ------------------------------------" -ForegroundColor Green
    Write-Host "  Office 365 Subscriptions:   $office365Count" -ForegroundColor Cyan
    Write-Host "  Desktop Standard Licenses:  $desktopStandard" -ForegroundColor Cyan
    Write-Host "  Desktop Professional:       $desktopProfessional" -ForegroundColor Cyan
    Write-Host "  Total Active Installations: $($office365Count + $desktopStandard + $desktopProfessional)" -ForegroundColor Green

    Write-Host "`nCost Breakdown (Annual):" -ForegroundColor Green
    Write-Host "  Office 365 Cost:             `$$([Math]::Round($office365Annual, 2))" -ForegroundColor Green
    Write-Host "  Desktop License Cost:        `$$([Math]::Round($desktopAnnual, 2))" -ForegroundColor Green
    Write-Host "  ------------------------------------" -ForegroundColor Green
    Write-Host "  TOTAL ANNUAL COST:           `$$([Math]::Round($totalAnnual, 2))" -ForegroundColor Yellow

    Write-Host "`nMonthly Cost:                 `$$([Math]::Round($totalAnnual / 12, 2))" -ForegroundColor Yellow
}

function Send-EmailReport {
    param(
        [string]$ReportPath,
        [string]$Recipients = "it-admin@company.com",
        [string]$SmtpServer = "smtp.company.com"
    )
    
    # Use dynamic script path
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    
    if ([string]::IsNullOrEmpty($ReportPath)) {
        $ReportPath = "$scriptPath\Output\VisioAudit"
    }

    Write-Host "`n[*] Preparing email report..." -ForegroundColor Cyan

    $latestCSV = Get-ChildItem -Path $ReportPath -Filter "VisioAudit_*.csv" | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1
    $latestHTML = Get-ChildItem -Path $ReportPath -Filter "VisioAudit_*.html" | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

    if (!$latestHTML) {
        Write-Host "[-] No HTML report found" -ForegroundColor Red
        return
    }

    try {
        $emailParams = @{
            To          = $Recipients
            From        = "visio-audit@$([System.Net.Dns]::GetHostName())"
            Subject     = "Weekly Visio Installation Audit Report - $(Get-Date -Format 'MMMM dd, yyyy')"
            Body        = Get-Content $latestHTML.FullName -Raw
            BodyAsHtml  = $true
            SmtpServer  = $SmtpServer
            Attachments = @($latestCSV.FullName)
        }

        Send-MailMessage @emailParams

        Write-Host "[+] Report sent to: $Recipients" -ForegroundColor Green
    }
    catch {
        Write-Host "[-] Error sending email: $_" -ForegroundColor Red
    }
}

function New-ScheduledAudit {
    param(
        [ValidateSet("Daily", "Weekly", "Monthly")]
        [string]$Frequency = "Weekly",

        [ValidateSet("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
        [string]$DayOfWeek = "Sunday",

        [int]$Hour = 2
    )

    Write-Host "`n[*] Creating scheduled task for $Frequency Visio audit..." -ForegroundColor Cyan

    $scriptPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "visio-enterprise-audit.ps1"

    if (!(Test-Path $scriptPath)) {
        Write-Host "[-] Script not found at $scriptPath" -ForegroundColor Red
        return
    }

    $taskName = "VisioAudit-$Frequency"
    $taskTime = "{0:D2}:00" -f $Hour

    $trigger = switch ($Frequency) {
        "Daily" { New-ScheduledTaskTrigger -Daily -At $taskTime }
        "Weekly" { New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek -At $taskTime }
        "Monthly" { New-ScheduledTaskTrigger -Monthly -DayOfMonth 1 -At $taskTime }
    }

    $action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""

    try {
        Register-ScheduledTask `
            -TaskName $taskName `
            -Trigger $trigger `
            -Action $action `
            -RunLevel Highest `
            -Description "Automated Visio installation audit" `
            -Force

        Write-Host "[+] Task created: $taskName" -ForegroundColor Green
        Write-Host "    Frequency: $Frequency" -ForegroundColor Green
        Write-Host "    Time: $taskTime" -ForegroundColor Green
    }
    catch {
        Write-Host "[-] Error creating task: $_" -ForegroundColor Red
    }
}

function Select-ReportByDepartment {
    param(
        [string]$Department,
        [string]$ReportPath
    )
    
    # Use dynamic script path
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    
    if ([string]::IsNullOrEmpty($ReportPath)) {
        $ReportPath = "$scriptPath\Output\VisioAudit"
    }

    Write-Host "`n[*] Filtering report for department: $Department" -ForegroundColor Cyan

    $latestReport = Get-ChildItem -Path $ReportPath -Filter "VisioAudit_*.csv" | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

    if (!$latestReport) {
        Write-Host "[-] No reports found" -ForegroundColor Red
        return
    }

    $data = Import-Csv $latestReport.FullName
    $filtered = $data | Where-Object { $_.ComputerName -like "*$Department*" }

    Write-Host "`n[+] Found $($filtered.Count) computers in '$Department'" -ForegroundColor Green
    $filtered | Format-Table -Property ComputerName, VisioInstalled, VisioVersion, VisioEdition, LastUsedDate

    $outputFile = Join-Path (Split-Path $latestReport.FullName) "VisioAudit_${Department}_$(Get-Date -Format 'yyyyMMdd').csv"
    $filtered | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Host "[+] Exported to: $outputFile" -ForegroundColor Green
}

# ============================================================================
# MAIN MENU LOOP
# ============================================================================

function Start-InteractiveMenu {
    do {
        Show-Menu
        $choice = Read-Host "Enter selection (1-12)"

        switch ($choice) {
            "1" {
                $prefix = Read-Host "Enter computer prefix (default: GOT)"
                if ([string]::IsNullOrEmpty($prefix)) { $prefix = "GOT" }
                Invoke-FullAudit -ComputerPrefix $prefix
            }
            "2" {
                Invoke-UsageAnalytics
            }
            "3" {
                $months = Read-Host "Months inactive (default: 6)"
                if ([string]::IsNullOrEmpty($months)) { $months = 6 }
                Find-UnusedVisio -MonthsInactive $months
            }
            "4" {
                Export-ToExcel
                Pause
            }
            "5" {
                New-CostAnalysis
                Pause
            }
            "6" {
                Show-ReportSummary
                Pause
            }
            "7" {
                Compare-Reports
                Pause
            }
            "8" {
                $recipients = Read-Host "Email recipients (comma-separated)"
                Send-EmailReport -Recipients $recipients
                Pause
            }
            "9" {
                Write-Host "`n1. Daily"
                Write-Host "2. Weekly"
                Write-Host "3. Monthly"
                $freq = Read-Host "Select frequency"
                $freqMap = @{ "1" = "Daily"; "2" = "Weekly"; "3" = "Monthly" }
                if ($freqMap.ContainsKey($freq)) {
                    New-ScheduledAudit -Frequency $freqMap[$freq]
                }
                else {
                    Write-Host "Invalid frequency selection. Please choose 1, 2, or 3." -ForegroundColor Red
                }
                Pause
            }
            "10" {
                $dept = Read-Host "Enter department name (e.g., SALES, DESIGN)"
                Select-ReportByDepartment -Department $dept
                Pause
            }
            "11" {
                Write-Host "`nDepartment summary feature - enter wildcard (e.g., 'SALES*')" -ForegroundColor Cyan
                $pattern = Read-Host "Enter pattern"
                # Use dynamic script path
                $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
                $reportPath = "$scriptPath\Output\VisioAudit"
                $latestReport = Get-ChildItem -Path $reportPath -Filter "VisioAudit_*.csv" | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

                if ($latestReport) {
                    $data = Import-Csv $latestReport.FullName
                    $grouped = $data | Where-Object { $_.ComputerName -like $pattern } | Group-Object -Property { $_.ComputerName -replace '^([A-Z]+).*', '$1' }

                    foreach ($group in $grouped) {
                        $withVisio = $group.Group | Where-Object { $_.VisioInstalled -eq "Yes" }
                        $standardEdition = $withVisio | Where-Object { $_.VisioEdition -eq "Standard" }
                        $professionalEdition = $withVisio | Where-Object { $_.VisioEdition -eq "Professional" }
                        Write-Host "`n$($group.Name):" -ForegroundColor Cyan
                        Write-Host "  Total: $($group.Group.Count)" -ForegroundColor Yellow
                        Write-Host "  With Visio: $($withVisio.Count)" -ForegroundColor Green
                        Write-Host "    - Standard: $($standardEdition.Count)" -ForegroundColor Cyan
                        Write-Host "    - Professional: $($professionalEdition.Count)" -ForegroundColor Cyan
                    }
                }
                Pause
            }
            "12" {
                Write-Host "`nExiting..." -ForegroundColor Yellow
                exit
            }
            "13" {
                Show-AccessErrorGuidance
                Pause
            }
            default {
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
        Clear-Host
    } while ($true)
}

# ============================================================================
# ENTRY POINT
# ============================================================================

Clear-Host
Write-Host "`n" -ForegroundColor Cyan
Start-InteractiveMenu
