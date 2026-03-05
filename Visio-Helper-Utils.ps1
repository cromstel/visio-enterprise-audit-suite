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

function Get-CredentialCachePath {
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    return Join-Path $scriptPath "VisioScanCredential.txt"
}

function Save-VisioScanCredentialCache {
    param(
        [Parameter(Mandatory = $true)]
        [PSCredential]$Credential
    )

    $cachePath = Get-CredentialCachePath
    $cacheDir = Split-Path $cachePath
    if (-not (Test-Path $cacheDir)) {
        New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null
    }

    $payload = @{
        UserName = $Credential.UserName
        Password = $Credential.Password | ConvertFrom-SecureString
        SavedAt  = (Get-Date).ToString("o")
    }

    $payload | ConvertTo-Json | Set-Content -Path $cachePath -Encoding UTF8
}

function Load-VisioScanCredentialCache {
    $cachePath = Get-CredentialCachePath
    if (-not (Test-Path $cachePath)) {
        return $null
    }

    try {
        $json = Get-Content -Path $cachePath -Raw -ErrorAction Stop
        $payload = $json | ConvertFrom-Json
        if ($payload.UserName -and $payload.Password) {
            $secureString = $payload.Password | ConvertTo-SecureString
            return New-Object System.Management.Automation.PSCredential($payload.UserName, $secureString)
        }
    }
    catch {
        Write-Host "[-] Unable to load cached credential: $_" -ForegroundColor Yellow
    }
    return $null
}

function Clear-VisioScanCredentialCache {
    $cachePath = Get-CredentialCachePath
    if (Test-Path $cachePath) {
        Remove-Item -Path $cachePath -Force -ErrorAction SilentlyContinue
    }
    $script:VisioScanCredential = $null
    Write-Host "[+] Cached credential cleared" -ForegroundColor Yellow
}

function Get-VisioScanCredential {
    [CmdletBinding()]
    param(
        [switch]$Force
    )

    if ($Force) {
        Clear-VisioScanCredentialCache
    }

    if (-not $Force -and $script:VisioScanCredential) {
        return $script:VisioScanCredential
    }

    $cached = if (-not $Force) { Load-VisioScanCredentialCache } else { $null }
    if ($cached) {
        $script:VisioScanCredential = $cached
        Write-Host "[*] Loaded cached credential for $($cached.UserName)" -ForegroundColor Green
        return $cached
    }

    $useAlt = Read-Host "Enter local admin credential now? (Y/N)"
    if ($useAlt -match '^[Yy]') {
        $script:VisioScanCredential = Get-Credential -Message "Enter the local admin credential used by the audit"
        if ($script:VisioScanCredential) {
            Save-VisioScanCredentialCache -Credential $script:VisioScanCredential
        }
    }

    return $script:VisioScanCredential
}

function Get-DefaultOutputPath {
    $base = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    return Join-Path $base "Output\VisioAudit"
}

function Get-SnapshotDirectory {
    $snapshotsRoot = Join-Path (Get-DefaultOutputPath) "Snapshots"
    if (-not (Test-Path $snapshotsRoot)) {
        New-Item -ItemType Directory -Path $snapshotsRoot -Force | Out-Null
    }
    return $snapshotsRoot
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
    Write-Host "8.  Send Report Notification (Email + Webhook)" -ForegroundColor Yellow
    Write-Host "9.  Schedule Recurring Audit" -ForegroundColor Yellow
    Write-Host "10. Select Report by Department" -ForegroundColor Yellow
    Write-Host "11. Generate Department Summary" -ForegroundColor Yellow
    Write-Host "12. Exit" -ForegroundColor Yellow
    Write-Host "13. Show Access Error 397 Guidance" -ForegroundColor Yellow
    Write-Host "14. Show Scheduled Task Status" -ForegroundColor Yellow
    Write-Host "15. Clear Cached Credential" -ForegroundColor Yellow
    Write-Host "16. Cleanup Old Reports" -ForegroundColor Yellow
    Write-Host "17. Run Health Check" -ForegroundColor Yellow
    Write-Host "18. Remediate Access 397 endpoints" -ForegroundColor Yellow
    Write-Host "19. Export snapshot & push JSON" -ForegroundColor Yellow
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

function Get-LatestAuditReportFiles {
    param(
        [string]$ReportPath
    )

    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    if ([string]::IsNullOrEmpty($ReportPath)) {
        $ReportPath = "$scriptPath\Output\VisioAudit"
    }

    $latestCsv = Get-ChildItem -Path $ReportPath -Filter "VisioAudit_*.csv" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $latestHtml = Get-ChildItem -Path $ReportPath -Filter "VisioAudit_*.html" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1

    return [PSCustomObject]@{
        Csv = if ($latestCsv) { $latestCsv.FullName } else { $null }
        Html = if ($latestHtml) { $latestHtml.FullName } else { $null }
    }
}

function Send-ReportEmail {
    param(
        [string]$Recipients = "it-admin@company.com",
        [string]$SmtpServer = "smtp.company.com",
        [string]$Subject = "Weekly Visio Installation Audit Report",
        [string]$Body = "",
        [string[]]$Attachments = @()
    )

    if (-not $Recipients) {
        Write-Host "[-] No recipients provided for email" -ForegroundColor Red
        return
    }

    $emailParams = @{
        To          = $Recipients -split "[,;]" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        From        = "visio-audit@$([System.Net.Dns]::GetHostName())"
        Subject     = $Subject
        Body        = $Body
        BodyAsHtml  = $true
        SmtpServer  = $SmtpServer
    }

    if ($Attachments.Count -gt 0) {
        $emailParams.Attachments = $Attachments
    }

    try {
        Send-MailMessage @emailParams -ErrorAction Stop
        Write-Host "[+] Email notification queued to: $Recipients" -ForegroundColor Green
    }
    catch {
        Write-Host "[-] Failed to send email: $_" -ForegroundColor Red
    }
}

function Invoke-ReportWebhook {
    param(
        [string]$WebhookUrl,
        [string]$Summary,
        [string]$Title = "Visio Audit Report"
    )

    if ([string]::IsNullOrEmpty($WebhookUrl)) {
        Write-Host "[-] Webhook URL is empty; skipping webhook notification" -ForegroundColor Yellow
        return
    }

    $payload = @{
        title = $Title
        text  = $Summary
    } | ConvertTo-Json -Depth 3

    try {
        Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType "application/json" -Body $payload -ErrorAction Stop | Out-Null
        Write-Host "[+] Webhook notification sent" -ForegroundColor Green
    }
    catch {
        Write-Host "[-] Webhook notification failed: $_" -ForegroundColor Red
    }
}

function Send-ReportNotification {
    param(
        [string]$ReportPath,
        [string]$Recipients = "",
        [string]$SmtpServer = "smtp.company.com",
        [string]$WebhookUrl,
        [string]$Subject = "Weekly Visio Installation Audit Report",
        [switch]$IncludeAttachments = $true,
        [switch]$UseZip
    )

    $files = Get-LatestAuditReportFiles -ReportPath $ReportPath

    if (-not $files.Csv -or -not $files.Html) {
        Write-Host "[-] No reports found in $ReportPath" -ForegroundColor Red
        return
    }

    $csvData = Import-Csv -Path $files.Csv -ErrorAction SilentlyContinue
    $total = $csvData.Count
    $visioInstalls = ($csvData | Where-Object { $_.VisioInstalled -eq "Yes" }).Count
    $withVisioOffline = ($csvData | Where-Object { $_.VisioInstalled -eq "Yes" -and $_.IsOnline -eq "No" }).Count
    $errors = ($csvData | Where-Object { $_.Error -and $_.Error -ne "None" }).Count

    $body = @"
<p>Visio audit report generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
<ul>
  <li>Total computers scanned: $total</li>
  <li>Visio installations: $visioInstalls</li>
  <li>Visio offline counts: $withVisioOffline</li>
  <li>Access errors: $errors</li>
</ul>
<p>HTML report attached: $(Split-Path $files.Html -Leaf)</p>
"@

    $attachments = @()
    if ($IncludeAttachments) {
        if ($UseZip) {
            $zipPath = Join-Path $env:TEMP "VisioAuditReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').zip"
            Compress-Archive -Path @($files.Csv, $files.Html) -DestinationPath $zipPath -Force
            $attachments += $zipPath
        }
        else {
            $attachments += $files.Html
            $attachments += $files.Csv
        }
    }

    if ($Recipients) {
        Send-ReportEmail -Recipients $Recipients -SmtpServer $SmtpServer -Subject $Subject -Body $body -Attachments $attachments
    }
    elseif (-not $WebhookUrl) {
        Write-Host "[-] No recipients or webhook defined; no notification sent" -ForegroundColor Yellow
    }

    if ($WebhookUrl) {
        $summary = "Total: $total | Visio: $visioInstalls | Offline: $withVisioOffline | Errors: $errors"
        Invoke-ReportWebhook -WebhookUrl $WebhookUrl -Summary $summary -Title $Subject
    }

}

function Cleanup-OldReports {
    param(
        [int]$DaysToKeep = 30,
        [int]$MaxFiles = 0,
        [string]$ReportPath
    )

    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    if ([string]::IsNullOrEmpty($ReportPath)) {
        $ReportPath = "$scriptPath\Output\VisioAudit"
    }

    $files = @()
    $files += Get-ChildItem -Path $ReportPath -File -Filter "VisioAudit_*.csv" -ErrorAction SilentlyContinue
    $files += Get-ChildItem -Path $ReportPath -File -Filter "VisioAudit_*.html" -ErrorAction SilentlyContinue
    if (!$files) {
        Write-Host "[*] No reports found in $ReportPath" -ForegroundColor Yellow
        return
    }

    $removeList = @()
    if ($DaysToKeep -gt 0) {
        $cutoff = (Get-Date).AddDays(-$DaysToKeep)
        $removeList += $files | Where-Object { $_.LastWriteTime -lt $cutoff }
    }
    if ($MaxFiles -gt 0) {
        $recent = $files | Sort-Object -Property LastWriteTime -Descending | Select-Object -First $MaxFiles
        $removeList += $files | Where-Object { $recent -notcontains $_ }
    }

    $toDelete = $removeList | Sort-Object -Unique
    if (!$toDelete) {
        Write-Host "[*] No files met the retention criteria" -ForegroundColor Green
        return
    }

    foreach ($file in $toDelete) {
        try {
            Remove-Item -Path $file.FullName -Force -ErrorAction Stop
            Write-Host "[+] Removed: $($file.Name)" -ForegroundColor Yellow
        }
        catch {
            Write-Host "[-] Failed to remove $($file.Name): $_" -ForegroundColor Red
        }
    }
}

function Invoke-VisioHealthCheck {
    param(
        [string]$TaskName = "VisioAudit-Weekly",
        [string]$ReportPath,
        [int]$SampleCount = 3
    )

    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    if ([string]::IsNullOrEmpty($ReportPath)) {
        $ReportPath = "$scriptPath\Output\VisioAudit"
    }

    $status = [ordered]@{}

    try {
        Get-ADDomain -ErrorAction Stop > $null
        $status.AD = "PASS"
    }
    catch {
        $status.AD = "FAIL: $_"
    }

    $latest = Get-LatestAuditReportFiles -ReportPath $ReportPath
    if ($latest.Csv) {
        $status.Reporting = "PASS (CSV: $(Split-Path $latest.Csv -Leaf))"
    }
    else {
        $status.Reporting = "MISSING"
    }

    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        $status.ScheduledTask = "PASS (Next run: $($task.NextRunTime))"
    }
    catch {
        $status.ScheduledTask = "WARN: $($_.Exception.Message)"
    }

    $winrmStatuses = @{}
    if ($latest.Csv) {
        try {
            $reportData = Import-Csv -Path $latest.Csv -ErrorAction Stop
            $targets = $reportData | Where-Object { $_.IsOnline -eq "Yes" } | Select-Object -ExpandProperty ComputerName -Unique | Select-Object -First $SampleCount
        }
        catch {
            $targets = @()
        }
    }
    else {
        $targets = @()
    }

    if ($targets.Count -eq 0) {
        $status.WinRM = "WARN: no online targets available"
    }
    else {
        foreach ($target in $targets) {
            try {
                Test-WSMan -ComputerName $target -ErrorAction Stop | Out-Null
                $winrmStatuses[$target] = "PASS"
            }
            catch {
                $winrmStatuses[$target] = "FAIL: $_"
            }
        }
        $status.WinRM = ($winrmStatuses.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" }) -join "; "
    }

    Write-Host "`nVISIO HEALTH CHECK" -ForegroundColor Cyan
    foreach ($key in $status.Keys) {
        $value = $status[$key]
        $color = if ($value -like "PASS*") { "Green" } elseif ($value -like "WARN*") { "Yellow" } else { "Red" }
        Write-Host ("{0}: {1}" -f $key, $value) -ForegroundColor $color
    }

    $dashboardPath = Join-Path $ReportPath "VisioHealthStatus.html"
    $rows = $status.GetEnumerator() | ForEach-Object {
        "<tr><td>$($_.Key)</td><td>$($_.Value)</td></tr>"
    }

    $html = @"
<!DOCTYPE html>
<html><head><meta charset='UTF-8'><title>Visio Health Check</title></head><body>
<h2>Visio Health Check</h2>
<table border='1' cellpadding='6'>
<tr><th>Check</th><th>Status</th></tr>
$($rows -join "`n")
</table>
<p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
</body></html>
"@

    $html | Out-File -FilePath $dashboardPath -Encoding UTF8
    Write-Host "[+] Health dashboard saved: $dashboardPath" -ForegroundColor Green
}

function New-ScheduledAudit {
    param(
        [ValidateSet("Daily", "Weekly", "Monthly")]
        [string]$Frequency = "Weekly",

        [ValidateSet("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
        [string]$DayOfWeek = "Sunday",

        [int]$Hour = 2,
        [int]$ThreadCount = 10,
        [string]$ComputerPrefix = "GOT",
        [string]$SearchBase,
        [string]$OutputPath
    )

    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    if ([string]::IsNullOrEmpty($OutputPath)) {
        $OutputPath = "$scriptPath\Output\VisioAudit"
    }

    $taskActionArgs = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "$scriptPath\visio-enterprise-audit.ps1",
        "-ThreadCount", $ThreadCount,
        "-ComputerPrefix", $ComputerPrefix,
        "-OutputPath", $OutputPath
    )

    if ($SearchBase) {
        $taskActionArgs += "-SearchBase"
        $taskActionArgs += "`"$SearchBase`""
    }

    $quotedActionArgs = $taskActionArgs | ForEach-Object {
        if ($_ -match '\s') {
            "`"$_`""
        }
        else {
            $_
        }
    }

    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument ($quotedActionArgs -join " ")

    $taskTime = "{0:D2}:00" -f [math]::Max(0, [math]::Min(23, $Hour))

    $trigger = switch ($Frequency) {
        "Daily"   { New-ScheduledTaskTrigger -Daily -At $taskTime }
        "Weekly"  { New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek -At $taskTime }
        "Monthly" { New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At $taskTime }
    }

    $taskName = "VisioAudit-$Frequency"
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew

    try {
        Register-ScheduledTask -TaskName $taskName -Trigger $trigger -Action $action -Settings $settings -Principal $principal -Force
        Write-Host "[+] Scheduled task registered: $taskName" -ForegroundColor Green
        Write-Host "    Frequency: $Frequency" -ForegroundColor Cyan
        Write-Host "    Day/Time : $($DayOfWeek) @ $taskTime" -ForegroundColor Cyan
        Write-Host "    Threads  : $ThreadCount" -ForegroundColor Cyan
        Write-Host "    SearchBase: $($SearchBase -or 'Full domain')" -ForegroundColor Cyan
    }
    catch {
        Write-Host "[-] Unable to register scheduled task: $_" -ForegroundColor Red
    }
}

function Show-ScheduledAuditStatus {
    param(
        [string]$TaskName = "VisioAudit-Weekly"
    )

    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        $lastRun = $task.LastRunTime
        $nextRun = $task.NextRunTime
        $result = $task.LastTaskResult
        Write-Host "`nScheduled Task: $TaskName" -ForegroundColor Cyan
        Write-Host "  Next Run : $nextRun" -ForegroundColor Green
        Write-Host "  Last Run : $lastRun" -ForegroundColor Yellow
        Write-Host "  Last Result: $result" -ForegroundColor Yellow
    }
    catch {
        Write-Host "[-] Unable to retrieve task status: $_" -ForegroundColor Red
    }
}

function Show-AccessErrorGuidance {
    Write-Host "`n[*] Access Denied (CIM 397) Guidance" -ForegroundColor Cyan
    Write-Host "CIM 397 indicates the audit host could not establish CIM/WMI communication." -ForegroundColor Yellow
    Write-Host "Best options to work around it:" -ForegroundColor Green
    Write-Host "  1. Re-run the audit with -ScanCredential (local administrator) so CIM runs under elevated context." -ForegroundColor White
    Write-Host "  2. Enable WinRM/remote registry on the target hosts or make sure the firewall allows DCOM/CIM traffic." -ForegroundColor White
    Write-Host "  3. When remoting is blocked, deploy the Visio scripts locally (scheduling or remote execution tools) and centralize results later." -ForegroundColor White
    Write-Host "  4. Run Office-Version-Detector locally to prove the 10.0.60910/Visio 2016 Standard or Professional install and then rerun the audit with valid credentials." -ForegroundColor White
    Write-Host ""
}

function Invoke-Access397Remediation {
    [CmdletBinding()]
    param(
        [string[]]$ComputerNames,
        [string]$SearchBase,
        [string]$ComputerPrefix = "GOT",
        [PSCredential]$ScanCredential,
        [bool]$ApplyLocalAccountTokenFilterPolicy = $true,
        [bool]$EnableWmiFirewallRule = $true,
        [bool]$EnsurePsRemoting = $true,
        [string]$SaveReportPath
    )

    if (-not $ComputerNames -or $ComputerNames.Count -eq 0) {
        $domainComputers = Get-DomainComputers -ComputerPrefix $ComputerPrefix -SearchBase $SearchBase
        $ComputerNames = $domainComputers | Select-Object -ExpandProperty Name
    }

    if (-not $ComputerNames -or $ComputerNames.Count -eq 0) {
        Write-Host "[-] No target computers found for remediation" -ForegroundColor Red
        return
    }

    $remediationScript = {
        param($applyToken, $enableFirewall, $ensureRemoting)
        $jobStatus = [ordered]@{
            ComputerName                   = $env:COMPUTERNAME
            LocalAccountTokenFilterPolicy  = "Not run"
            WmiFirewallStatus              = "Not run"
            PsRemotingStatus               = "Not run"
            WinRMServiceStatus             = $null
            WsManConnectivity             = $null
            Timestamp                      = (Get-Date).ToString("o")
            Error                          = $null
        }

        try {
            if ($applyToken) {
                $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                New-Item -Path $regPath -Force | Out-Null
                New-ItemProperty -Path $regPath -Name "LocalAccountTokenFilterPolicy" -Value 1 -PropertyType DWord -Force | Out-Null
                $jobStatus.LocalAccountTokenFilterPolicy = (Get-ItemProperty -Path $regPath -Name "LocalAccountTokenFilterPolicy" -ErrorAction Stop).LocalAccountTokenFilterPolicy
            }
        }
        catch {
            $jobStatus.LocalAccountTokenFilterPolicy = "Error: $($_.Exception.Message)"
        }

        try {
            if ($enableFirewall) {
                if (Get-Command -Name Set-NetFirewallRule -ErrorAction SilentlyContinue) {
                    $rules = Get-NetFirewallRule -Group "Windows Management Instrumentation (WMI)" -ErrorAction SilentlyContinue
                    if ($rules) {
                        $rules | Set-NetFirewallRule -Enabled True -ErrorAction SilentlyContinue
                        $jobStatus.WmiFirewallStatus = "Enabled"
                    }
                    else {
                        $jobStatus.WmiFirewallStatus = "No rules found"
                    }
                }
                else {
                    Enable-NetFirewallRule -DisplayGroup "Windows Management Instrumentation (WMI)" -ErrorAction SilentlyContinue | Out-Null
                    $jobStatus.WmiFirewallStatus = "Firewall group enabled"
                }
            }
        }
        catch {
            $jobStatus.WmiFirewallStatus = "Error: $($_.Exception.Message)"
        }

        try {
            if ($ensureRemoting) {
                Enable-PSRemoting -Force -SkipNetworkProfileCheck -ErrorAction SilentlyContinue | Out-Null
                $winRm = Get-Service -Name WinRM -ErrorAction SilentlyContinue
                $jobStatus.WinRMServiceStatus = if ($winRm) { $winRm.Status } else { "Unavailable" }
                $jobStatus.PsRemotingStatus = "Configured"
            }
        }
        catch {
            $jobStatus.PsRemotingStatus = "Error: $($_.Exception.Message)"
        }

        try {
            Test-WSMan -ComputerName $env:COMPUTERNAME -ErrorAction Stop | Out-Null
            $jobStatus.WsManConnectivity = "Self-check passed"
        }
        catch {
            $jobStatus.WsManConnectivity = "Failed: $($_.Exception.Message)"
        }

        return $jobStatus
    }

    $results = @()
    foreach ($computer in $ComputerNames) {
        $status = [ordered]@{
            ComputerName                  = $computer
            LocalAccountTokenFilterPolicy = "Pending"
            WmiFirewallStatus             = "Pending"
            PsRemotingStatus              = "Pending"
            WinRMServiceStatus            = "Pending"
            WsManConnectivity            = "Pending"
            Timestamp                     = (Get-Date).ToString("o")
            Error                         = $null
        }

        if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
            $status.Error = "Host offline or unreachable"
            $results += [PSCustomObject]$status
            continue
        }

        $invokeParams = @{
            ComputerName = $computer
            ScriptBlock  = $remediationScript
            ArgumentList = @($ApplyLocalAccountTokenFilterPolicy, $EnableWmiFirewallRule, $EnsurePsRemoting)
            ErrorAction  = "Stop"
        }
        if ($ScanCredential) {
            $invokeParams.Credential = $ScanCredential
        }

        try {
            $remoteResult = Invoke-Command @invokeParams
            if ($remoteResult) {
                $status = $remoteResult | Select-Object -First 1
            }
        }
        catch {
            $status.Error = $_.Exception.Message
            $status.LocalAccountTokenFilterPolicy = "Failed"
            $status.WmiFirewallStatus = "Failed"
            $status.PsRemotingStatus = "Failed"
            $status.WinRMServiceStatus = "Failed"
            $status.WsManConnectivity = "Failed"
        }

        $results += [PSCustomObject]$status
    }

    $reportRoot = if ([string]::IsNullOrEmpty($SaveReportPath)) { Join-Path (Get-DefaultOutputPath) "RemediationReports" } else { $SaveReportPath }
    if (-not (Test-Path $reportRoot)) {
        New-Item -ItemType Directory -Path $reportRoot -Force | Out-Null
    }
    $outputFile = Join-Path $reportRoot "Access397Remediation_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $results | ConvertTo-Json -Depth 5 | Out-File -FilePath $outputFile -Encoding UTF8
    Write-Host "[+] Remediation report saved: $outputFile" -ForegroundColor Green

    return $results
}

function Export-VisioAuditSnapshot {
    [CmdletBinding()]
    param(
        [string]$ReportPath,
        [string]$SnapshotDirectory,
        [string]$WebhookUrl
    )

    if ([string]::IsNullOrEmpty($ReportPath)) {
        $ReportPath = Get-DefaultOutputPath
    }

    $latestCsv = Get-ChildItem -Path $ReportPath -Filter "VisioAudit_*.csv" -ErrorAction SilentlyContinue |
        Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1
    $latestHtml = Get-ChildItem -Path $ReportPath -Filter "VisioAudit_*.html" -ErrorAction SilentlyContinue |
        Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

    if (-not $latestCsv) {
        Write-Host "[-] No CSV reports found in $ReportPath" -ForegroundColor Red
        return
    }

    $data = Import-Csv -Path $latestCsv.FullName -ErrorAction SilentlyContinue
    $summary = [ordered]@{
        Timestamp           = (Get-Date).ToString("o")
        CsvReport           = Split-Path $latestCsv.FullName -Leaf
        HtmlReport          = if ($latestHtml) { Split-Path $latestHtml.FullName -Leaf } else { $null }
        Totals              = [ordered]@{
            TotalComputers     = $data.Count
            VisioInstalled     = ($data | Where-Object { $_.VisioInstalled -eq "Yes" }).Count
            VisioProfessional  = ($data | Where-Object { $_.VisioEdition -eq "Professional" }).Count
            VisioStandard      = ($data | Where-Object { $_.VisioEdition -eq "Standard" }).Count
            Office365Installs  = ($data | Where-Object { $_.Office365Install -eq "True" }).Count
            OfflineComputers   = ($data | Where-Object { $_.IsOnline -eq "No" }).Count
            AccessErrors       = ($data | Where-Object { $_.Error -and $_.Error -ne "None" }).Count
            Cim397Errors       = ($data | Where-Object { $_.Error -match "CIM 397" }).Count
        }
        Samples             = $data | Select-Object -First 5 ComputerName, VisioVersion, VisioEdition, LastUsedDate, Error
    }

    $snapshotDir = if ([string]::IsNullOrEmpty($SnapshotDirectory)) { Get-SnapshotDirectory } else { $SnapshotDirectory }
    if (-not (Test-Path $snapshotDir)) {
        New-Item -ItemType Directory -Path $snapshotDir -Force | Out-Null
    }

    $snapshotFile = Join-Path $snapshotDir "VisioAuditSnapshot_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $summary | ConvertTo-Json -Depth 6 | Out-File -FilePath $snapshotFile -Encoding UTF8
    Write-Host "[+] JSON snapshot saved: $snapshotFile" -ForegroundColor Green

    if (-not [string]::IsNullOrEmpty($WebhookUrl)) {
        try {
            $jsonPayload = Get-Content -Path $snapshotFile -Raw
            Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType "application/json" -Body $jsonPayload -ErrorAction Stop | Out-Null
            Write-Host "[+] Snapshot posted to webhook: $WebhookUrl" -ForegroundColor Green
        }
        catch {
            Write-Host "[-] Failed to POST snapshot: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    return $snapshotFile
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
        $choice = Read-Host "Enter selection (1-19)"

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
                Write-Host "`n[*] Prepare report notification" -ForegroundColor Cyan
                $reportPath = Read-Host "Report folder (default: Output\VisioAudit)"
                if ([string]::IsNullOrEmpty($reportPath)) {
                    $reportPath = Get-DefaultOutputPath
                }
                $recipients = Read-Host "Email recipients (comma-separated, leave blank to skip)"
                $smtpServer = Read-Host "SMTP server (default: smtp.company.com)"
                if ([string]::IsNullOrEmpty($smtpServer)) { $smtpServer = "smtp.company.com" }
                $subject = Read-Host "Email subject (default: Visio audit summary)"
                if ([string]::IsNullOrEmpty($subject)) { $subject = "Visio Installation Audit Report" }
                $useZip = Read-Host "Compress attachments into a ZIP? (Y/N)"
                $webhook = Read-Host "Optional webhook URL (leave blank to skip)"
                $zipSwitch = $useZip -match '^[Yy]'
                if ($zipSwitch) {
                    Send-ReportNotification -ReportPath $reportPath -Recipients $recipients -SmtpServer $smtpServer -Subject $subject -WebhookUrl $webhook -IncludeAttachments -UseZip
                }
                else {
                    Send-ReportNotification -ReportPath $reportPath -Recipients $recipients -SmtpServer $smtpServer -Subject $subject -WebhookUrl $webhook -IncludeAttachments
                }
                Pause
            }
            "9" {
                Write-Host "`n[*] Schedule recurring Visio audit" -ForegroundColor Cyan
                Write-Host "1. Daily"
                Write-Host "2. Weekly"
                Write-Host "3. Monthly"
                $freq = Read-Host "Select frequency"
                $freqMap = @{ "1" = "Daily"; "2" = "Weekly"; "3" = "Monthly" }
                if (-not $freqMap.ContainsKey($freq)) {
                    Write-Host "Invalid selection, choose 1/2/3" -ForegroundColor Red
                    Pause
                    break
                }
                $hour = Read-Host "Hour of day (0-23, default 2)"
                [int]$parsedHour = 2
                if (-not [int]::TryParse($hour, [ref]$parsedHour)) { $parsedHour = 2 }
                $day = Read-Host "Day of week for weekly schedule (default Sunday)"
                if ([string]::IsNullOrEmpty($day)) { $day = "Sunday" }
                $threads = Read-Host "Thread count (1-64, default 10)"
                [int]$parsedThreads = 10
                if (-not [int]::TryParse($threads, [ref]$parsedThreads)) { $parsedThreads = 10 }
                $prefix = Read-Host "Computer prefix (default GOT)"
                if ([string]::IsNullOrEmpty($prefix)) { $prefix = "GOT" }
                $searchBase = Read-Host "LDAP SearchBase (leave blank for domain)"
                New-ScheduledAudit -Frequency $freqMap[$freq] -Hour $parsedHour -DayOfWeek $day -ThreadCount $parsedThreads -ComputerPrefix $prefix -SearchBase $searchBase
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
            "14" {
                $taskName = Read-Host "Scheduled task name (default VisioAudit-Weekly)"
                if ([string]::IsNullOrEmpty($taskName)) { $taskName = "VisioAudit-Weekly" }
                Show-ScheduledAuditStatus -TaskName $taskName
                Pause
            }
            "15" {
                Clear-VisioScanCredentialCache
                Pause
            }
            "16" {
                $reportPath = Read-Host "Report folder (default Output\\VisioAudit)"
                if ([string]::IsNullOrEmpty($reportPath)) {
                    $reportPath = Get-DefaultOutputPath
                }
                $days = Read-Host "Keep reports for how many days? (default 30)"
                [int]$parsedDays = 30
                if (-not [int]::TryParse($days, [ref]$parsedDays)) { $parsedDays = 30 }
                $maxFiles = Read-Host "Maximum files to keep (0 for no limit, default 0)"
                [int]$parsedMax = 0
                if (-not [int]::TryParse($maxFiles, [ref]$parsedMax)) { $parsedMax = 0 }
                Cleanup-OldReports -ReportPath $reportPath -DaysToKeep $parsedDays -MaxFiles $parsedMax
                Pause
            }
            "17" {
                $taskName = Read-Host "Scheduled task name (default VisioAudit-Weekly)"
                if ([string]::IsNullOrEmpty($taskName)) { $taskName = "VisioAudit-Weekly" }
                $sampleCount = Read-Host "WinRM target sample count (default 3)"
                [int]$parsedSample = 3
                if (-not [int]::TryParse($sampleCount, [ref]$parsedSample)) { $parsedSample = 3 }
                $reportPath = Read-Host "Report folder for health check (default Output\\VisioAudit)"
                if ([string]::IsNullOrEmpty($reportPath)) {
                    $reportPath = Get-DefaultOutputPath
                }
                Invoke-VisioHealthCheck -TaskName $taskName -ReportPath $reportPath -SampleCount $parsedSample
                Pause
            }
            "18" {
                Write-Host "`n[*] Access 397 remediation helper" -ForegroundColor Cyan
                $computerInput = Read-Host "Computer names (comma-separated) or leave blank to scan OU"
                $computerList = if (-not [string]::IsNullOrEmpty($computerInput)) {
                    $computerInput -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                }
                $searchBase = Read-Host "LDAP SearchBase (optional, leave blank for entire domain)"
                $prefix = Read-Host "Computer prefix (default: GOT)"
                if ([string]::IsNullOrEmpty($prefix)) { $prefix = "GOT" }
                $credential = Get-VisioScanCredential
                Invoke-Access397Remediation -ComputerNames $computerList -SearchBase $searchBase -ComputerPrefix $prefix -ScanCredential $credential
                Pause
            }
            "19" {
                Write-Host "`n[*] Exporting JSON snapshot" -ForegroundColor Cyan
                $reportPath = Read-Host "Report folder (default Output\\VisioAudit)"
                if ([string]::IsNullOrEmpty($reportPath)) { $reportPath = Get-DefaultOutputPath }
                $webhook = Read-Host "Webhook URL to POST snapshot (optional)"
                Export-VisioAuditSnapshot -ReportPath $reportPath -WebhookUrl $webhook
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
