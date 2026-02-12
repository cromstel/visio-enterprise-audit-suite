#requires -RunAsAdministrator
#requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Enterprise Visio Installation Audit Script
    Scans all domain computers for Visio installations and last usage

.DESCRIPTION
    This script queries Active Directory for all computers, then uses WMI/Registry
    to check for Visio installations. Supports Office 365 and Office 2019 only.
    Generates CSV and HTML reports.

.PARAMETER OutputPath
    Directory to save reports (default: script directory\Output\VisioAudit)

.PARAMETER ComputerFilter
    Filter for AD computer search (default: all enabled computers)

.PARAMETER ThreadCount
    Number of parallel jobs (default: 10)

.PARAMETER SearchBase
    LDAP path to the OU to search for computers

.EXAMPLE
    .\Visio-Enterprise-Audit.ps1 -OutputPath "C:\Reports" -ThreadCount 20
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$ComputerFilter = "*",

    [Parameter(Mandatory = $false)]
    [int]$ThreadCount = 10,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeOfflineComputers = $false,

    [Parameter(Mandatory = $false)]
    [string]$ComputerPrefix = "GOT",

    # Target specific OU within the domain
    [Parameter(Mandatory = $false)]
    [string]$SearchBase
)

# ============================================================================
# CONFIGURATION
# ============================================================================

# Determine script directory for output operations
if ($PSScriptRoot -or $MyInvocation.MyCommand.Path) {
    $ScriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
}
else {
    Write-Error "Unable to determine script location. Exiting."
    exit 1
}

# Set default OutputPath if not provided
if ([string]::IsNullOrEmpty($OutputPath)) {
    $OutputPath = "$ScriptDirectory\Output\VisioAudit"
}

$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"

# Set default SearchBase if not provided
if ([string]::IsNullOrEmpty($SearchBase)) {
    $SearchBase = "OU=Workstations,OU=NEOS CIB 64,OU=SE,OU=CRDF,DC=euro,DC=net,DC=intra"
}

# Office 365 and Office 2019 installation paths (x64 only)
$VisioPaths = @(
    "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE"         # Office 365/2019 x64
)

# Registry paths for detecting Office 365/2019 (x64 only)
$RegistryPaths = @(
    "HKLM:\Software\Microsoft\Office\16.0\Common\InstallRoot"          # Office 365/2019
)

# ============================================================================
# FUNCTIONS
# ============================================================================

function Initialize-AuditEnvironment {
    try {
        if (!(Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        Write-Host "Output directory: $OutputPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create output directory '$OutputPath': $_.Exception.Message"
        Write-Error "Please check permissions and path validity."
        exit 1
    }
}

function Get-DomainComputers {
    param(
        [string]$Filter = "*",
        [string]$SearchBase
    )

    Write-Host "`n[*] Querying Active Directory for computers..." -ForegroundColor Cyan
    Write-Host "[*] Targeting OU: $SearchBase" -ForegroundColor Yellow
    
    try {
        $getADParams = @{
            Filter      = "Name -like '$Filter'"
            Properties  = @("Name", "OperatingSystem", "LastLogonDate")
            ErrorAction = "Stop"
            SearchBase  = $SearchBase
        }
        
        if ($Domain) {
            $getADParams.Server = $Domain
        }
        
        $computers = Get-ADComputer @getADParams |
            Where-Object { $_.OperatingSystem -like "*Windows*" } |
            Sort-Object -Property Name

        Write-Host "[+] Found $($computers.Count) computers in Active Directory" -ForegroundColor Green
        return $computers
    }
    catch {
        Write-Host "[-] Error querying Active Directory: $_" -ForegroundColor Red
        exit 1
    }
}

function Test-ComputerConnectivity {
    param(
        [string]$ComputerName
    )

    $ping = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet
    return $ping
}

function Get-VisioInstallationInfo {
    param(
        [string]$ComputerName,
        [string[]]$Paths,
        [string[]]$RegPaths
    )

    $result = @{
        ComputerName      = $ComputerName
        IsOnline          = $false
        VisioInstalled    = $false
        VisioVersion      = $null
        InstallPath       = $null
        LastAccessTime    = $null
        LastUsedDate      = $null
        Office365Install  = $false
        OfficeVersion     = $null
        Error             = $null
    }

    # Test connectivity
    if (!(Test-ComputerConnectivity -ComputerName $ComputerName)) {
        $result.Error = "Computer offline"
        return $result
    }

    $result.IsOnline = $true

    try {
        # Initialize session variable before try block
        $session = $null
        $session = New-CimSession -ComputerName $ComputerName -ErrorAction Stop

        # Query installed Office products
        $officeProducts = Get-CimInstance -CimSession $session `
            -ClassName Win32_Product `
            -Filter "Name LIKE '%Office%' OR Name LIKE '%Visio%'" `
            -ErrorAction SilentlyContinue

        # Check for Office 365 subscription
        $office365 = $officeProducts | Where-Object { $_.Name -match "Microsoft 365|Office 365" }
        if ($office365) {
            $result.Office365Install = $true
            $result.OfficeVersion = $office365.Name
        }

        # Check for Visio specifically
        $visioProduct = $officeProducts | Where-Object { $_.Name -match "Visio" }
        if ($visioProduct) {
            $result.VisioInstalled = $true
            $result.VisioVersion = $visioProduct.Version
        }

        # Check file system for Visio executable
        foreach ($path in $Paths) {
            $remoteFile = "\\$ComputerName\$($path -replace ':', '$')"
            
            if (Test-Path $remoteFile -ErrorAction SilentlyContinue) {
                $fileInfo = Get-Item $remoteFile -ErrorAction SilentlyContinue
                if ($fileInfo) {
                    $result.VisioInstalled = $true
                    $result.InstallPath = $path
                    $result.LastAccessTime = $fileInfo.LastAccessTime
                    $result.LastUsedDate = $fileInfo.LastAccessTime.ToString("yyyy-MM-dd HH:mm:ss")
                    break
                }
            }
        }

        # Check registry for installation details
        foreach ($regPath in $RegPaths) {
            try {
                $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
                    [Microsoft.Win32.RegistryHive]::LocalMachine,
                    $ComputerName
                )
                
                $key = $regKey.OpenSubKey($regPath.Replace("HKLM:\", ""))
                if ($key) {
                    $installPath = $key.GetValue("Path")
                    if ($installPath) {
                        $result.InstallPath = $installPath
                        $result.Office365Install = $true
                    }
                }
            }
            catch {
                # Continue to next registry path
            }
        }

        Remove-CimSession $session
    }
    catch {
        $result.Error = "WMI access denied or unavailable"
    }

    return $result
}

function Get-VisioLastUsedFromRegistry {
    param(
        [string]$ComputerName
    )

    try {
        $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            [Microsoft.Win32.RegistryHive]::CurrentUser,
            $ComputerName
        )

        # Office 365/2019 recent document registry path
        $path = "Software\Microsoft\Office\16.0\Common\Open Find"
        $key = $regKey.OpenSubKey($path)

        if ($key) {
            $lastOpen = $key.GetValue("LastOpenedPath")
            if ($lastOpen) {
                return $lastOpen
            }
        }

        # Try alternative path for Visio specifically
        $path = "Software\Microsoft\Office\16.0\Visio\Recent"
        $key = $regKey.OpenSubKey($path)

        if ($key) {
            $subValues = $key.GetValueNames()
            if ($subValues.Count -gt 0) {
                return "Recent documents found"
            }
        }
    }
    catch {
        # Return null if registry access fails
    }

    return $null
}

function Invoke-VisioScan {
    param(
        [array]$Computers,
        [int]$ThreadCount
    )

    Write-Host "`n[*] Starting Visio scan on $($Computers.Count) computers..." -ForegroundColor Cyan
    Write-Host "[*] Using $ThreadCount parallel threads`n" -ForegroundColor Cyan

    $results = @()
    $scanned = 0

    # Use RunspacePool for parallel processing
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $ThreadCount)
    $runspacePool.Open()

    $jobs = @()

    foreach ($computer in $Computers) {
        $scriptBlock = {
            param($ComputerName, $VisioPaths, $RegistryPaths)
            
            # Inline function logic for runspace scope compatibility
            $result = @{
                ComputerName      = $ComputerName
                IsOnline          = $false
                VisioInstalled    = $false
                VisioVersion      = $null
                InstallPath       = $null
                LastAccessTime    = $null
                LastUsedDate      = $null
                Office365Install  = $false
                OfficeVersion     = $null
                Error             = $null
            }
            
            # Test connectivity
            $ping = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet
            if (!$ping) {
                $result.Error = "Computer offline"
                return $result
            }
            
            $result.IsOnline = $true
            
            try {
                # Check installed software via WMI
                $session = $null
                $session = New-CimSession -ComputerName $ComputerName -ErrorAction Stop

                # Query installed Office products
                $officeProducts = Get-CimInstance -CimSession $session `
                    -ClassName Win32_Product `
                    -Filter "Name LIKE '%Office%' OR Name LIKE '%Visio%'" `
                    -ErrorAction SilentlyContinue

                # Check for Office 365 subscription
                $office365 = $officeProducts | Where-Object { $_.Name -match "Microsoft 365|Office 365" }
                if ($office365) {
                    $result.Office365Install = $true
                    $result.OfficeVersion = $office365.Name
                }

                # Check for Visio specifically
                $visioProduct = $officeProducts | Where-Object { $_.Name -match "Visio" }
                if ($visioProduct) {
                    $result.VisioInstalled = $true
                    $result.VisioVersion = $visioProduct.Version
                }

                # Check file system for Visio executable
                foreach ($path in $VisioPaths) {
                    $remoteFile = "\\$ComputerName\$($path -replace ':', '$')"
                    
                    if (Test-Path $remoteFile -ErrorAction SilentlyContinue) {
                        $fileInfo = Get-Item $remoteFile -ErrorAction SilentlyContinue
                        if ($fileInfo) {
                            $result.VisioInstalled = $true
                            $result.InstallPath = $path
                            $result.LastAccessTime = $fileInfo.LastAccessTime
                            $result.LastUsedDate = $fileInfo.LastAccessTime.ToString("yyyy-MM-dd HH:mm:ss")
                            break
                        }
                    }
                }

                # Check registry for installation details
                foreach ($regPath in $RegistryPaths) {
                    try {
                        $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
                            [Microsoft.Win32.RegistryHive]::LocalMachine,
                            $ComputerName
                        )
                        
                        $key = $regKey.OpenSubKey($regPath.Replace("HKLM:\", ""))
                        if ($key) {
                            $installPath = $key.GetValue("Path")
                            if ($installPath) {
                                $result.InstallPath = $installPath
                                $result.Office365Install = $true
                            }
                        }
                    }
                    catch {
                        # Continue to next registry path
                    }
                }

                Remove-CimSession $session
            }
            catch {
                $result.Error = "WMI access denied or unavailable"
            }
            
            return $result
        }

        $job = [powershell]::Create().AddScript($scriptBlock).AddArgument($computer.Name).AddArgument($VisioPaths).AddArgument($RegistryPaths)
        $job.RunspacePool = $runspacePool
        $jobs += @{
            Job    = $job
            Handle = $job.BeginInvoke()
        }
    }

    # Collect results
    foreach ($jobItem in $jobs) {
        $result = $jobItem.Job.EndInvoke($jobItem.Handle)
        $results += $result

        [int]$scanned++
        $installed = if ($result.VisioInstalled) { "[X] Visio" } else { "[ ] No Visio" }
        $status = if ($result.IsOnline) { "Online" } else { "Offline" }

        Write-Progress -Activity "Scanning Computers" -Status "$scanned/$($Computers.Count) - $($result.ComputerName) [$status] $installed" -PercentComplete (($scanned / $Computers.Count) * 100)
    }

    $runspacePool.Close()
    $runspacePool.Dispose()

    return $results
}

function ConvertTo-HtmlReport {
    param(
        [array]$Results,
        [string]$OutputFile
    )

    $visioInstalled = $Results | Where-Object { $_.VisioInstalled }
    $visioNotInstalled = $Results | Where-Object { !$_.VisioInstalled }
    $offline = $Results | Where-Object { !$_.IsOnline }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Visio Installation Audit Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 40px 20px;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
        }
        
        .header p {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .metrics {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 40px;
            background: #f8f9fa;
            border-bottom: 1px solid #e9ecef;
        }
        
        .metric {
            background: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }
        
        .metric:hover {
            transform: translateY(-4px);
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.15);
        }
        
        .metric-value {
            font-size: 2.5em;
            font-weight: bold;
            color: #667eea;
            margin-bottom: 8px;
        }
        
        .metric-label {
            font-size: 0.95em;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .content {
            padding: 40px;
        }
        
        .section {
            margin-bottom: 50px;
        }
        
        .section h2 {
            color: #333;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #667eea;
            font-size: 1.8em;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
        }
        
        th {
            background: #f8f9fa;
            color: #333;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            border-bottom: 2px solid #e9ecef;
            font-size: 0.95em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #e9ecef;
        }
        
        tr:hover {
            background: #f8f9fa;
        }
        
        .status-online {
            color: #28a745;
            font-weight: bold;
        }
        
        .status-offline {
            color: #dc3545;
            font-weight: bold;
        }
        
        .badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.85em;
            font-weight: 600;
            text-transform: uppercase;
        }
        
        .badge-success {
            background: #d4edda;
            color: #155724;
        }
        
        .badge-danger {
            background: #f8d7da;
            color: #721c24;
        }
        
        .badge-warning {
            background: #fff3cd;
            color: #856404;
        }
        
        .badge-office365 {
            background: #cce5ff;
            color: #004085;
        }
        
        .footer {
            background: #f8f9fa;
            padding: 20px 40px;
            text-align: center;
            color: #666;
            border-top: 1px solid #e9ecef;
            font-size: 0.9em;
        }
        
        .empty-state {
            text-align: center;
            padding: 40px;
            color: #999;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Visio Installation Audit Report</h1>
            <p>Enterprise-wide scan of Visio installations across domain computers</p>
        </div>
        
        <div class="metrics">
            <div class="metric">
                <div class="metric-value">$($Results.Count)</div>
                <div class="metric-label">Total Computers</div>
            </div>
            <div class="metric">
                <div class="metric-value">$($visioInstalled.Count)</div>
                <div class="metric-label">Visio Installed</div>
            </div>
            <div class="metric">
                <div class="metric-value">$($offline.Count)</div>
                <div class="metric-label">Offline</div>
            </div>
            <div class="metric">
                <div class="metric-value">$(if ($Results.Count -gt 0) { ($visioInstalled.Count / $Results.Count * 100).ToString('F1') } else { '0' })%</div>
                <div class="metric-label">Installation Rate</div>
            </div>
        </div>
        
        <div class="content">
            <div class="section">
                <h2>Computers with Visio Installed</h2>
"@

    if ($visioInstalled.Count -gt 0) {
        $html += @"
                    <table>
                        <thead>
                            <tr>
                                <th>Computer Name</th>
                                <th>Status</th>
                                <th>Visio Version</th>
                                <th>Office 365</th>
                                <th>Last Used</th>
                                <th>Install Path</th>
                            </tr>
                        </thead>
                        <tbody>
"@
        foreach ($computer in $visioInstalled) {
            $office365Badge = if ($computer.Office365Install) { '<span class="badge badge-office365">Office 365</span>' } else { '<span class="badge badge-warning">Desktop</span>' }
            $html += @"
                            <tr>
                                <td><strong>$($computer.ComputerName)</strong></td>
                                <td><span class="status-online">Online</span></td>
                                <td>$($computer.VisioVersion)</td>
                                <td>$office365Badge</td>
                                <td>$(if ($computer.LastUsedDate) { $computer.LastUsedDate } else { 'N/A' })</td>
                                <td style="font-size: 0.9em; color: #666;">$(if ($computer.InstallPath) { $computer.InstallPath } else { 'Standard' })</td>
                            </tr>
"@
        }
        $html += @"
                        </tbody>
                    </table>
"@
    }
    else {
        $html += @"
                    <div class="empty-state">
                        <p>No computers with Visio installed found</p>
                    </div>
"@
    }

    $html += @"
            </div>
            
            <div class="section">
                <h2>Computers without Visio</h2>
"@

    if (($visioNotInstalled | Where-Object { $_.IsOnline }).Count -gt 0) {
        $html += @"
                    <table>
                        <thead>
                            <tr>
                                <th>Computer Name</th>
                                <th>Status</th>
                                <th>Last Logon</th>
                            </tr>
                        </thead>
                        <tbody>
"@
        foreach ($computer in $visioNotInstalled | Where-Object { $_.IsOnline }) {
            $html += @"
                            <tr>
                                <td><strong>$($computer.ComputerName)</strong></td>
                                <td><span class="status-online">Online</span></td>
                                <td>N/A</td>
                            </tr>
"@
        }
        $html += @"
                        </tbody>
                    </table>
"@
    }
    else {
        $html += @"
                    <div class="empty-state">
                        <p>All online computers checked</p>
                    </div>
"@
    }

    $html += @"
            </div>
            
            <div class="section">
                <h2>Offline Computers</h2>
"@

    if ($offline.Count -gt 0) {
        $html += @"
                    <table>
                        <thead>
                            <tr>
                                <th>Computer Name</th>
                                <th>Status</th>
                            </tr>
                        </thead>
                        <tbody>
"@
        foreach ($computer in $offline) {
            $html += @"
                            <tr>
                                <td><strong>$($computer.ComputerName)</strong></td>
                                <td><span class="status-offline">Offline</span></td>
                            </tr>
"@
        }
        $html += @"
                        </tbody>
                    </table>
"@
    }
    else {
        $html += @"
                    <div class="empty-state">
                        <p>All computers are online</p>
                    </div>
"@
    }

    $html += @"
            </div>
        </div>
        
        <div class="footer">
            <p>Report generated on $(Get-Date -Format "dddd, MMMM dd, yyyy 'at' HH:mm:ss")</p>
            <p>Enterprise Visio Installation Audit System</p>
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $OutputFile -Encoding UTF8
    Write-Host "[+] HTML report saved: $OutputFile" -ForegroundColor Green
}

function Export-ResultsToCSV {
    param(
        [array]$Results,
        [string]$OutputFile
    )

    $Results | Select-Object `
        @{ Name = "ComputerName"; Expression = { $_.ComputerName } },
        @{ Name = "IsOnline"; Expression = { if ($_.IsOnline) { "Yes" } else { "No" } } },
        @{ Name = "VisioInstalled"; Expression = { if ($_.VisioInstalled) { "Yes" } else { "No" } } },
        @{ Name = "VisioVersion"; Expression = { if ($_.VisioVersion) { $_.VisioVersion } else { "N/A" } } },
        @{ Name = "Office365"; Expression = { if ($_.Office365Install) { "Yes" } else { "No" } } },
        @{ Name = "LastUsedDate"; Expression = { if ($_.LastUsedDate) { $_.LastUsedDate } else { "N/A" } } },
        @{ Name = "InstallPath"; Expression = { if ($_.InstallPath) { $_.InstallPath } else { "N/A" } } },
        @{ Name = "Error"; Expression = { if ($_.Error) { $_.Error } else { "None" } } } |
    Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

    Write-Host "[+] CSV report saved: $OutputFile" -ForegroundColor Green
}

function Get-AuditSummary {
    param(
        [array]$Results
    )

    $summary = @{
        TotalComputers    = $Results.Count
        OnlineComputers   = ($Results | Where-Object { $_.IsOnline }).Count
        OfflineComputers  = ($Results | Where-Object { !$_.IsOnline }).Count
        VisioInstalled    = ($Results | Where-Object { $_.VisioInstalled }).Count
        Office365Installs = ($Results | Where-Object { $_.Office365Install }).Count
        AccessErrors      = ($Results | Where-Object { $_.Error }).Count
    }

    return $summary
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

function Main {
    Write-Host ("`n" + ("=" * 80))
    Write-Host "  ENTERPRISE VISIO INSTALLATION AUDIT" -ForegroundColor Cyan
    Write-Host "  Office 365 / 2019 (x64 Only)" -ForegroundColor Cyan
    Write-Host (("=" * 80) + "`n")

    Initialize-AuditEnvironment

    # Get computers from Active Directory
    Write-Host "[*] Targeting OU: $SearchBase" -ForegroundColor Yellow
    $computers = Get-DomainComputers -Filter $ComputerFilter -SearchBase $SearchBase

    if ($computers.Count -eq 0) {
        Write-Host "[-] No computers found matching filter" -ForegroundColor Red
        exit 1
    }

    # Perform scan
    $results = Invoke-VisioScan -Computers $computers -ThreadCount $ThreadCount

    # Generate reports
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $OutputPath "VisioAudit_$timestamp.csv"
    $htmlPath = Join-Path $OutputPath "VisioAudit_$timestamp.html"

    Export-ResultsToCSV -Results $results -OutputFile $csvPath
    ConvertTo-HtmlReport -Results $results -OutputFile $htmlPath

    # Display summary
    $summary = Get-AuditSummary -Results $results

    Write-Host ("`n" + ("=" * 80))
    Write-Host "  AUDIT SUMMARY" -ForegroundColor Green
    Write-Host (("=" * 80))
    Write-Host "Total Computers Scanned: $($summary.TotalComputers)" -ForegroundColor Yellow
    Write-Host "Online Computers: $($summary.OnlineComputers)" -ForegroundColor Green
    Write-Host "Offline Computers: $($summary.OfflineComputers)" -ForegroundColor Yellow
    Write-Host "Computers with Visio: $($summary.VisioInstalled)" -ForegroundColor Green
    Write-Host "Office 365 Installations: $($summary.Office365Installs)" -ForegroundColor Cyan
    Write-Host "Access Errors: $($summary.AccessErrors)" -ForegroundColor Red
    Write-Host (("=" * 80) + "`n")

    Write-Host "[+] Audit complete!" -ForegroundColor Green
    Write-Host "[+] Reports available at: $OutputPath" -ForegroundColor Green
    Write-Host "[+] CSV: $csvPath" -ForegroundColor Yellow
    Write-Host "[+] HTML: $htmlPath" -ForegroundColor Yellow
}

Main
