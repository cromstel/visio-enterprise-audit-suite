Scanning Active Directory (AD) for specific installed software like Microsoft Visio is best achieved by combining Active Directory querying (to get a list of computers) with a method to check the registry or software inventory on those computers. 

**Here are the primary methods to scan for Visio installations:**

# Method 1: PowerShell Script (Most Efficient) 
Using PowerShell, you can query AD for all computer objects and then remotely check their registry for the Visio installation string. 

----powershell

# Get all active computers from AD
$Computers = Get-ADComputer -Filter 'Enabled -eq $true' | Select-Object -ExpandProperty Name

# Iterate through computers and check registry
foreach ($Computer in $Computers) {
    if (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
        Invoke-Command -ComputerName $Computer -ScriptBlock {
            $Visio = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | 
                     Where-Object { $_.DisplayName -match "Visio" }
            if ($Visio) {
                Write-Host "$env:COMPUTERNAME - Visio Installed"
            }
        }
    }
}

----

This script retrieves all enabled computer accounts from Active Directory, checks if they are online, and then remotely queries the registry for any installed software that matches "Visio".

**Note: This requires administrator rights on the remote machines.**

----
# Method 2: SCCM/MECM Query (Best for Enterprise)

If your organization uses Microsoft Endpoint Configuration Manager (SCCM/MECM), this is the most reliable method, as it relies on stored hardware inventory. 

1. Navigate to Monitoring > Reporting > Reports.
2. Run a search for software reports.
3. Use the following WQL query to find Visio Standard or Professional

----sql

select SMS_R_System.Name, SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName 
from SMS_R_System 
inner join SMS_G_System_ADD_REMOVE_PROGRAMS on SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceID = SMS_R_System.ResourceId 
where SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName like "%Visio%"

```