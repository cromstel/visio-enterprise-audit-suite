#requires -RunAsAdministrator
#requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Visio Usage Analytics Script
    Tracks detailed Visio usage patterns and generates analytics

.DESCRIPTION
    Collects advanced usage metrics for Visio 2021/2019/2016 including:
    - Last user to run Visio
    - Number of Visio processes
    - Recent Visio documents
    - File association metadata
    - License information for Office 365
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$ComputerFilter = "*",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 64)]
    [int]$ThreadCount = 10,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeOfflineComputers = $false,

    [Parameter(Mandatory = $false)]
    [string]$ComputerPrefix = "GOT",

    # Target specific OU within the domain
    [Parameter(Mandatory = $false)]
    [string]$SearchBase,

    [Parameter(Mandatory = $false)]
    [string[]]$ComputerNames = @(),

    [Parameter(Mandatory = $false)]
    [int]$UsageDaysBack = 90,

    [Parameter(Mandatory = $false)]
    [PSCredential]$ScanCredential
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

# Set default SearchBase if not provided
if ([string]::IsNullOrEmpty($SearchBase)) {
    $SearchBase = "OU=Workstations,OU=NEOS CIB 64,OU=SE,OU=CRDF,DC=euro,DC=net,DC=intra"
}

[PSCredential]$script:UsageScanCredential = $ScanCredential
[int]$script:UsageFallbackWindow = $UsageDaysBack

# Fallback helper scripts for credential-limited hosts
$VisioUsageFallbackScript = {
    param($DaysBack)

    $cutoffDate = (Get-Date).AddDays(-$DaysBack)
    $usage = [ordered]@{
        ProcessRunning        = $false
        ActiveUser            = $null
        RecentDocuments       = @()
        VisioTempFiles        = @()
        FileAssociations      = @()
        LicenseStatus         = "Unknown"
        LastUserRun           = $null
        RunCount              = 0
        EstimatedUsageHours   = 0
        Error                 = $null
    }

    try {
        $visioProcesses = Get-Process -Name VISIO -ErrorAction SilentlyContinue
        if ($visioProcesses) {
            $usage.ProcessRunning = $true
            foreach ($proc in $visioProcesses) {
                $owner = Get-WmiObject -Class Win32_Process -Filter "ProcessId=$($proc.Id)" -ErrorAction SilentlyContinue
                if ($owner) {
                    $ownerInfo = $owner.GetOwner()
                    if ($ownerInfo.ReturnValue -eq 0) {
                        $usage.ActiveUser = "$($ownerInfo.Domain)\$($ownerInfo.User)"
                        break
                    }
                }
            }
        }

        $docPaths = Get-ChildItem -Path "C:\Users\*\Documents\*.vsd*" -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -gt $cutoffDate }
        if ($docPaths) {
            $usage.RecentDocuments = $docPaths | Select-Object -Property FullName, LastWriteTime -First 10
            $usage.RunCount = $docPaths.Count
        }

        $tempFiles = Get-ChildItem -Path "C:\Users\*\AppData\Local\Microsoft\Office\16.0\*" -Recurse -Filter "*Visio*" -ErrorAction SilentlyContinue
        if ($tempFiles) {
            $usage.VisioTempFiles = $tempFiles | Select-Object -Property FullName, LastAccessTime
        }

        $assocPatterns = @("*.vsd", "*.vsdx", "*.vsdm")
        foreach ($pattern in $assocPatterns) {
            $matches = Get-ChildItem -Path "C:\Users\*\Documents\*$pattern" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 3
            if ($matches) {
                $usage.FileAssociations += $matches | Select-Object -Property FullName, Extension
            }
        }

    }
    catch {
        $usage.Error = "Fallback usage analysis failed: $_"
    }

    return [pscustomobject]$usage
}

$VisioLicenseFallbackScript = {
    $licenseInfo = [ordered]@{
        IsLicensed       = $false
        LicenseStatus    = "Unknown"
        SubscriptionType = $null
        LastActivation   = $null
        Error            = $null
    }

    $licensePaths = @(
        "HKLM:\Software\Microsoft\Office\16.0\Common\Identity\Licenses",
    )

    foreach ($path in $licensePaths) {
        if (Test-Path $path) {
            $keys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
            foreach ($key in $keys) {
                $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                if ($props) {
                    if ($props.Status) {
                        $licenseInfo.IsLicensed = $true
                        $licenseInfo.LicenseStatus = $props.Status
                    }
                    if ($props.SubscriptionType) {
                        $licenseInfo.SubscriptionType = $props.SubscriptionType
                    }
                    if ($props.LastActivation) {
                        $licenseInfo.LastActivation = $props.LastActivation
                    }
                    break
                }
            }
            if ($licenseInfo.IsLicensed) { break }
        }
    }

    return [pscustomobject]$licenseInfo
}

$VisioConfigFallbackScript = {
    $config = [ordered]@{
        StartupLocation     = $null
        AutoRecoveryEnabled = $false
        AutoRecoveryInterval= $null
        DefaultFileFormat   = $null
        RecentFilesCount    = 0
        AddInsInstalled     = @()
        Error               = $null
    }

    try {
        $optionsPath = "HKCU:\Software\Microsoft\Office\16.0\Visio\Options"
        if (Test-Path $optionsPath) {
            $options = Get-ItemProperty -Path $optionsPath -ErrorAction SilentlyContinue
            if ($options) {
                $config.AutoRecoveryEnabled = [bool]$options.AutoRecovery
                $config.AutoRecoveryInterval = $options.AutoRecoveryInterval
                $config.DefaultFileFormat = $options.DefaultSaveFormat
            }
        }

        $addinPath = "HKCU:\Software\Microsoft\Office\16.0\Visio\Resiliency"
        if (Test-Path $addinPath) {
            $addinKey = Get-Item -Path $addinPath -ErrorAction SilentlyContinue
            if ($addinKey) {
                $config.AddInsInstalled = $addinKey.GetSubKeyNames()
            }
        }
    }
    catch {
        $config.Error = "Fallback configuration read failed: $_"
    }

    return [pscustomobject]$config
}

function New-UsageCimSession {
    param([string]$ComputerName)

    $sessionParams = @{
        ComputerName = $ComputerName
        ErrorAction  = "Stop"
    }
    if ($script:UsageScanCredential) {
        $sessionParams.Credential = $script:UsageScanCredential
    }

    return New-CimSession @sessionParams
}

function Invoke-UsageFallbackCommand {
    param(
        [string]$ComputerName,
        [scriptblock]$ScriptBlock,
        $ArgumentList = @()
    )

    $invokeParams = @{
        ComputerName = $ComputerName
        ScriptBlock  = $ScriptBlock
        ErrorAction  = "SilentlyContinue"
    }
    if ($script:UsageScanCredential) {
        $invokeParams.Credential = $script:UsageScanCredential
    }
    if ($ArgumentList -and $ArgumentList.Count -gt 0) {
        $invokeParams.ArgumentList = $ArgumentList
    }

    return Invoke-Command @invokeParams
}

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
        Write-Error "Failed to create output directory '$OutputPath': $($_.Exception.Message)"
        Write-Error "Please check permissions and path validity."
        exit 1
    }
}

function Get-DomainComputers {
    param(
        [string]$Filter = "*",
        [string]$SearchBase,
        [string]$ComputerPrefix = "GOT"
    )

    Write-Host "`n[*] Querying Active Directory for computers..." -ForegroundColor Cyan
    Write-Host "[*] Targeting OU: $SearchBase" -ForegroundColor Yellow
    Write-Host "[*] Using computer prefix filter: $ComputerPrefix*" -ForegroundColor Yellow
    
    try {
        # Build filter using ComputerPrefix
        $prefixFilter = "$ComputerPrefix*"
        $getADParams = @{
            Filter      = "Name -like '$prefixFilter'"
            Properties  = @("Name", "OperatingSystem", "LastLogonDate")
            ErrorAction = "Stop"
            SearchBase  = $SearchBase
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

# ============================================================================
# DETAILED VISIO USAGE ANALYSIS
# ============================================================================

function Get-DetailedVisioUsage {
    param(
        [string]$ComputerName,
        [int]$DaysBack = $script:UsageFallbackWindow
    )

    $usage = @{
        ComputerName          = $ComputerName
        IsOnline              = $false
        ProcessRunning        = $false
        ActiveUser            = $null
        RecentDocuments       = @()
        VisioTempFiles        = @()
        FileAssociations      = @()
        LicenseStatus         = $null
        LastUserRun           = $null
        RunCount              = 0
        EstimatedUsageHours   = 0
        Error                 = $null
    }

    if (!(Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
        $usage.Error = "Computer offline"
        return $usage
    }

    $usage.IsOnline = $true

    try {
        $session = New-UsageCimSession -ComputerName $ComputerName

        # Check if Visio is currently running
        $visioProcess = Get-CimInstance -CimSession $session `
            -ClassName Win32_Process `
            -Filter "Name='VISIO.EXE'" `
            -ErrorAction SilentlyContinue

        if ($visioProcess) {
            $usage.ProcessRunning = $true

            $ownerFound = $false
            foreach ($process in @($visioProcess)) {
                $ownerInfo = Invoke-CimMethod -InputObject $process -MethodName GetOwner -ErrorAction SilentlyContinue
                if ($ownerInfo -and $ownerInfo.ReturnValue -eq 0 -and $ownerInfo.User) {
                    $usage.ActiveUser = $ownerInfo.User
                    $ownerFound = $true
                    break
                }
            }
            if (-not $ownerFound) {
                $usage.ActiveUser = "Unknown"
            }
        }

        # Get file associations for Visio (VSD, VSDX, etc.)
        $visioFiles = Get-CimInstance -CimSession $session `
            -ClassName CIM_DataFile `
            -Filter "Name LIKE '%.vsd%' OR Name LIKE '%.vsdx%'" `
            -ErrorAction SilentlyContinue

        foreach ($file in $visioFiles) {
            $usage.RecentDocuments += @{
                Path           = $file.Name
                LastModified   = $file.LastModified
                FileSize       = $file.FileSize
            }
        }

        # Get Visio temp/cache files
        $tempPath = "\\$ComputerName\C$\Users\*\AppData\Local\Microsoft\Office\16.0\*"
        $tempFiles = Get-Item $tempPath -Include "*Visio*" -ErrorAction SilentlyContinue

        foreach ($file in $tempFiles) {
            $usage.VisioTempFiles += @{
                Path           = $file.FullName
                LastAccessTime = $file.LastAccessTime
            }
        }

        # Check Office 365 license status
        # Registry path for Office 365 license information
        # Software\Microsoft\Office\16.0\Common\Identity

        if ($session) {
            Remove-CimSession $session
        }
    }
    catch {
        $usage.Error = "Usage analysis failed (CIM fallback will run): $_"
        Write-Host "[WARN] Access denied (CIM 397) on $ComputerName - running local fallback script" -ForegroundColor Yellow
        $fallbackResult = Invoke-UsageFallbackCommand -ComputerName $ComputerName -ScriptBlock $VisioUsageFallbackScript -ArgumentList $DaysBack
        if ($fallbackResult -is [System.Collections.IEnumerable]) {
            $fallbackResult = $fallbackResult | Select-Object -First 1
        }

        if ($fallbackResult) {
            $usage.ProcessRunning = $fallbackResult.ProcessRunning
            $usage.ActiveUser = $fallbackResult.ActiveUser
            $usage.RecentDocuments = $fallbackResult.RecentDocuments
            $usage.VisioTempFiles = $fallbackResult.VisioTempFiles
            $usage.FileAssociations = $fallbackResult.FileAssociations
            $usage.LicenseStatus = $fallbackResult.LicenseStatus
            $usage.LastUserRun = $fallbackResult.LastUserRun
            $usage.RunCount = $fallbackResult.RunCount
            $usage.EstimatedUsageHours = $fallbackResult.EstimatedUsageHours
            if ($fallbackResult.Error) {
                $usage.Error = "Usage fallback warning: $($fallbackResult.Error)"
            }
        }
    }

    return $usage
}

function Measure-VisioDocuments {
    param(
        [string]$ComputerName,
        [int]$DaysBack = 90
    )

    $cutoffDate = (Get-Date).AddDays(-$DaysBack)
    $results = @()

    try {
        $vsdPath = "\\$ComputerName\C$\Users\*\Documents\*.vsd*"
        $visioFiles = Get-ChildItem -Path $vsdPath -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.LastAccessTime -gt $cutoffDate }

        foreach ($file in $visioFiles) {
            $results += @{
                FileName       = $file.Name
                FullPath       = $file.FullName
                LastModified   = $file.LastWriteTime
                LastAccessed   = $file.LastAccessTime
                FileSize       = $file.Length
                DaysInactive   = ([int]((Get-Date) - $file.LastAccessTime).TotalDays)
            }
        }
    }
    catch {
        Write-Error "Error analyzing documents on $ComputerName : $_"
    }

    return $results
}

function Get-Office365LicenseStatus {
    param(
        [string]$ComputerName
    )

    $licenseInfo = @{
        ComputerName     = $ComputerName
        IsLicensed       = $false
        LicenseStatus    = "Unknown"
        SubscriptionType = $null
        LastActivation   = $null
        Error            = $null
    }

    $regKey = $null
    try {
        $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            $ComputerName
        )

        # Check Office 365 license registry paths
        $licensePaths = @(
            "Software\Microsoft\Office\16.0\Common\Identity\Licenses"
        )

        foreach ($path in $licensePaths) {
            $key = $regKey.OpenSubKey($path)
            if ($key) {
                $licenseStatus = $key.GetValue("Status")
                if ($licenseStatus) {
                    $licenseInfo.IsLicensed = $true
                    $licenseInfo.LicenseStatus = $licenseStatus
                }
            }
        }
    }
    catch {
        $licenseInfo.Error = "Cannot access license information: $_"
        $fallbackResult = Invoke-UsageFallbackCommand -ComputerName $ComputerName -ScriptBlock $VisioLicenseFallbackScript
        if ($fallbackResult -is [System.Collections.IEnumerable]) {
            $fallbackResult = $fallbackResult | Select-Object -First 1
        }
        if ($fallbackResult) {
            $licenseInfo.IsLicensed = $fallbackResult.IsLicensed
            $licenseInfo.LicenseStatus = $fallbackResult.LicenseStatus
            $licenseInfo.SubscriptionType = $fallbackResult.SubscriptionType
            $licenseInfo.LastActivation = $fallbackResult.LastActivation
            if ($fallbackResult.Error) {
                $licenseInfo.Error = "License fallback warning: $($fallbackResult.Error)"
            }
        }
    }

    return $licenseInfo
}

function Get-VisioConfiguration {
    param(
        [string]$ComputerName
    )

    $config = @{
        ComputerName          = $ComputerName
        StartupLocation       = $null
        AutoRecoveryEnabled   = $false
        AutoRecoveryInterval  = $null
        DefaultFileFormat     = $null
        RecentFilesCount      = 0
        AddInsInstalled       = @()
        Error                 = $null
    }

    try {
        $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            [Microsoft.Win32.RegistryHive]::CurrentUser,
            $ComputerName
        )

        # Visio 365/2021/2019/2016 options registry path
        $path = "Software\Microsoft\Office\16.0\Visio\Options"
        $key = $regKey.OpenSubKey($path)

        if ($key) {
            $config.AutoRecoveryEnabled = [bool]$key.GetValue("AutoRecovery")
            $config.AutoRecoveryInterval = $key.GetValue("AutoRecoveryInterval")
            $config.DefaultFileFormat = $key.GetValue("DefaultSaveFormat")
        }

        # Get add-ins
        $addinPath = "Software\Microsoft\Office\16.0\Visio\Resiliency"
        $addinKey = $regKey.OpenSubKey($addinPath)

        if ($addinKey) {
            $addins = $addinKey.GetSubKeyNames()
            $config.AddInsInstalled = $addins
        }
    }
    catch {
        $config.Error = "Cannot read configuration"
    }

    return $config
}

function New-UsageAnalyticsReport {
    param(
        [array]$UsageData,
        [string]$OutputPath
    )

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Visio Usage Analytics Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 40px 20px;
        }
        
        .container {
            max-width: 1400px;
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
        
        .content {
            padding: 40px;
        }
        
        .section {
            margin-bottom: 40px;
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
        }
        
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #e9ecef;
        }
        
        tr:hover {
            background: #f8f9fa;
        }
        
        .status-active {
            color: #28a745;
            font-weight: bold;
        }
        
        .status-inactive {
            color: #dc3545;
            font-weight: bold;
        }
        
        .footer {
            background: #f8f9fa;
            padding: 20px 40px;
            text-align: center;
            color: #666;
            border-top: 1px solid #e9ecef;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Visio Usage Analytics Report</h1>
            <p>Detailed usage patterns and activity tracking</p>
        </div>
        
        <div class="content">
            <div class="section">
                <h2>Active Visio Usage</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Computer</th>
                            <th>Currently Running</th>
                            <th>Active User</th>
                            <th>Recent Documents</th>
                            <th>Last Modified</th>
                        </tr>
                    </thead>
                    <tbody>
"@

    foreach ($computer in $UsageData) {
        if ($computer.IsOnline) {
            $running = if ($computer.ProcessRunning) { '<span class="status-active">Running</span>' } else { '<span class="status-inactive">Not Running</span>' }
			$user = if ($null -ne $computer.ActiveUser -and $computer.ActiveUser -ne "") {
				$computer.ActiveUser
			} else {
				"N/A"
			}

            $docCount = $computer.RecentDocuments.Count

            $html += @"
                        <tr>
                            <td><strong>$($computer.ComputerName)</strong></td>
                            <td>$running</td>
                            <td>$user</td>
                            <td>$docCount files</td>
                            <td>N/A</td>
                        </tr>
"@
        }
    }

    $html += @"
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="footer">
            <p>Report generated on $(Get-Date -Format "dddd, MMMM dd, yyyy 'at' HH:mm:ss")</p>
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $OutputPath -Encoding UTF8
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

function Main {
    Write-Host ("`n" + ("=" * 80))
    Write-Host "  VISIO USAGE ANALYTICS" -ForegroundColor Cyan
    Write-Host "  Detailed usage tracking and activity monitoring" -ForegroundColor Cyan
    Write-Host (("=" * 80) + "`n")

    Initialize-AuditEnvironment

    # Get computers from Active Directory
    Write-Host "[*] Targeting OU: $SearchBase" -ForegroundColor Yellow
    Write-Host "[*] Scanning computers with prefix: $ComputerPrefix*" -ForegroundColor Yellow
    
    if ($ComputerNames.Count -eq 0) {
        $computers = Get-DomainComputers -Filter $ComputerFilter -SearchBase $SearchBase -ComputerPrefix $ComputerPrefix
        $ComputerNames = $computers.Name
    }

    if ($ComputerNames.Count -eq 0) {
        Write-Host "[-] No computers found matching filter" -ForegroundColor Red
        exit 1
    }

    $results = @()

    foreach ($computer in $ComputerNames) {
        Write-Host "[*] Analyzing $computer..." -ForegroundColor Yellow
       
        $usage = Get-DetailedVisioUsage -ComputerName $computer -DaysBack $UsageDaysBack
        $documents = Measure-VisioDocuments -ComputerName $computer -DaysBack $UsageDaysBack
        $license = Get-Office365LicenseStatus -ComputerName $computer
        $config = Get-VisioConfiguration -ComputerName $computer

        $results += @{
            Usage         = $usage
            Documents     = $documents
            License       = $license
            Configuration = $config
        }
    }

    # Generate report
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPath = Join-Path $OutputPath "VisioUsageAnalytics_$timestamp.html"

    New-UsageAnalyticsReport -UsageData $results.Usage -OutputPath $reportPath

    Write-Host ("`n" + ("=" * 80))
    Write-Host "  USAGE ANALYTICS SUMMARY" -ForegroundColor Green
    Write-Host (("=" * 80))
    Write-Host "Computers Analyzed: $($ComputerNames.Count)" -ForegroundColor Yellow
    Write-Host "Report saved: $reportPath" -ForegroundColor Green
    Write-Host (("=" * 80) + "`n")

    Write-Host "[+] Analysis complete!" -ForegroundColor Green
    Write-Host "[+] Report saved: $reportPath" -ForegroundColor Yellow
}

Main
