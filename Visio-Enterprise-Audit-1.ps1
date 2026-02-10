#requires -RunAsAdministrator
#requires -Modules ActiveDirectory

param(
    [string]$OutputPath = "C:\Temp\VisioAudit",
    [string]$ComputerFilter = "*",
    [int]$ThreadCount = 10,
    [switch]$IncludeOfflineComputers = $false
)

$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"

$VisioPaths = @(
    "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\VISIO.EXE",
    "C:\Program Files\Microsoft Office\Office16\VISIO.EXE",
    "C:\Program Files (x86)\Microsoft Office\Office16\VISIO.EXE",
    "C:\Program Files\Microsoft Office\Office15\VISIO.EXE",
    "C:\Program Files (x86)\Microsoft Office\Office15\VISIO.EXE"
)

$RegistryPaths = @(
    "HKLM:\Software\Microsoft\Office\16.0\Common\InstallRoot",
    "HKLM:\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot",
    "HKLM:\Software\Microsoft\Office\15.0\Common\InstallRoot",
    "HKLM:\Software\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot"
)

function Initialize-AuditEnvironment {
    if (!(Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
}

function Get-DomainComputers {
    param([string]$Filter)
    Get-ADComputer -Filter "Name -like '$Filter'" -Properties Name, OperatingSystem |
        Where-Object { $_.OperatingSystem -like "*Windows*" } |
        Sort-Object Name
}

function Test-ComputerConnectivity {
    param([string]$ComputerName)
    Test-Connection -ComputerName $ComputerName -Count 1 -Quiet
}

function Get-VisioInstallationInfo {
    param($ComputerName, $Paths, $RegPaths)

    $result = @{
        ComputerName     = $ComputerName
        IsOnline         = $false
        VisioInstalled   = $false
        VisioVersion     = $null
        InstallPath      = $null
        LastUsedDate     = $null
        Office365Install = $false
        Error            = $null
    }

    if (!(Test-ComputerConnectivity $ComputerName)) {
        $result.Error = "Offline"
        return $result
    }

    $result.IsOnline = $true

    try {
        $session = New-CimSession -ComputerName $ComputerName
        $products = Get-CimInstance -CimSession $session -ClassName Win32_Product |
                    Where-Object { $_.Name -match "Office|Visio" }

        $visio = $products | Where-Object { $_.Name -match "Visio" }
        if ($visio) {
            $result.VisioInstalled = $true
            $result.VisioVersion = $visio.Version
        }

        foreach ($path in $Paths) {
            $remote = "\\$ComputerName\$($path -replace ':', '$')"
            if (Test-Path $remote) {
                $file = Get-Item $remote
                $result.InstallPath = $path
                $result.LastUsedDate = $file.LastAccessTime.ToString("yyyy-MM-dd HH:mm:ss")
                break
            }
        }

        Remove-CimSession $session
    }
    catch {
        $result.Error = "Access error"
    }

    return $result
}

function Invoke-VisioScan {
    param($Computers, $ThreadCount)

    $pool = [runspacefactory]::CreateRunspacePool(1, $ThreadCount)
    $pool.Open()
    $jobs = @()

    foreach ($c in $Computers) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        $ps.AddScript({
            param($n,$p,$r)
            Get-VisioInstallationInfo $n $p $r
        }).AddArgument($c.Name).AddArgument($VisioPaths).AddArgument($RegistryPaths) | Out-Null

        $jobs += @{ PS=$ps; Handle=$ps.BeginInvoke() }
    }

    $results = foreach ($j in $jobs) {
        $j.PS.EndInvoke($j.Handle)
    }

    $pool.Close()
    $pool.Dispose()
    return $results
}

function ConvertTo-HtmlReport {
    param($Results, $OutputFile)

    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Visio Audit</title>
<style>
body { font-family: Segoe UI, Arial; background:#f4f6f8; padding:20px; }
table { border-collapse: collapse; width:100%; }
th, td { border:1px solid #ccc; padding:8px; }
th { background:#eee; }
</style>
</head>
<body>
<h1>Visio Installation Audit</h1>
<table>
<tr>
<th>Computer</th>
<th>Online</th>
<th>Visio</th>
<th>Version</th>
<th>Last Used</th>
<th>Path</th>
</tr>
"@

    foreach ($c in $Results) {
        $html += "<tr>
<td>$($c.ComputerName)</td>
<td>$($c.IsOnline)</td>
<td>$($c.VisioInstalled)</td>
<td>$($c.VisioVersion)</td>
<td>$(if ($c.LastUsedDate) { $c.LastUsedDate } else { 'N/A' })</td>
<td>$(if ($c.InstallPath) { $c.InstallPath } else { 'N/A' })</td>
</tr>"
    }

    $html += "</table></body></html>"
    $html | Out-File $OutputFile -Encoding UTF8
}

function Export-ResultsToCSV {
    param($Results, $OutputFile)
    $Results | Export-Csv $OutputFile -NoTypeInformation -Encoding UTF8
}

function Main {
    Initialize-AuditEnvironment
    $computers = Get-DomainComputers $ComputerFilter
    $results = Invoke-VisioScan $computers $ThreadCount

    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    Export-ResultsToCSV $results (Join-Path $OutputPath "Visio_$ts.csv")
    ConvertTo-HtmlReport $results (Join-Path $OutputPath "Visio_$ts.html")
}

Main
