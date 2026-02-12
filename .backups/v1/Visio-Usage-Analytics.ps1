#requires -RunAsAdministrator
#requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Visio Usage Analytics Script
    Tracks detailed Visio usage patterns and generates analytics

.DESCRIPTION
    Collects advanced usage metrics including:
    - Last user to run Visio
    - Number of Visio processes
    - Recent Visio documents
    - File association metadata
    - License information for Office 365
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\Temp\VisioAudit",

    [Parameter(Mandatory = $false)]
    [string[]]$ComputerNames = @()
)

# ============================================================================
# DETAILED VISIO USAGE ANALYSIS
# ============================================================================

function Get-DetailedVisioUsage {
    param(
        [string]$ComputerName
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
        $session = New-CimSession -ComputerName $ComputerName -ErrorAction Stop

        # Check if Visio is currently running
        $visioProcess = Get-CimInstance -CimSession $session `
            -ClassName Win32_Process `
            -Filter "Name='VISIO.EXE'" `
            -ErrorAction SilentlyContinue

        if ($visioProcess) {
            $usage.ProcessRunning = $true
            $usage.ActiveUser = $visioProcess.GetOwner().User
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

        Remove-CimSession $session
    }
    catch {
        $usage.Error = "Analysis failed: $_"
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

    try {
        $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            $ComputerName
        )

        # Check Office 365 license registry paths
        $licensePaths = @(
            "Software\Microsoft\Office\16.0\Common\Identity\Licenses",
            "Software\Microsoft\Office\ClickToRun\Licensing"
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
        $licenseInfo.Error = "Cannot access license information"
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

        # Visio 365/2019 options registry path
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
            $running = if ($computer.ProcessRunning) { '<span class="status-active">● Running</span>' } else { '<span class="status-inactive">Not Running</span>' }
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

Write-Host "`n[*] Starting detailed Visio usage analysis..." -ForegroundColor Cyan

if ($ComputerNames.Count -eq 0) {
    $ComputerNames = (Get-ADComputer -Filter "Name -like '*'" -Properties Name).Name | Select-Object -First 20
}

$results = @()

foreach ($computer in $ComputerNames) {
    Write-Host "[*] Analyzing $computer..." -ForegroundColor Yellow
    
    $usage = Get-DetailedVisioUsage -ComputerName $computer
    $documents = Measure-VisioDocuments -ComputerName $computer
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

Write-Host "[+] Analysis complete!" -ForegroundColor Green
Write-Host "[+] Report saved: $reportPath" -ForegroundColor Yellow
