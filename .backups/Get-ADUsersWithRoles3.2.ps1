#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Extracts all ACTIVE users from OU=Users & Clients,OU=SE,OU=CRDF,DC=euro,DC=net,DC=intra
    along with every group they belong to, the resolved role per group, and permission-tracking metadata.

.DESCRIPTION
    Targets the following OU tree path in euro.net.intra:
        euro.net.intra > CRDF > SE > Users & Clients

    For every enabled (active) user the script:
      1. Collects ~40 identity, organisation, and account-health attributes.
      2. Resolves ALL group memberships — direct and nested up to 8 levels deep.
      3. Classifies each group into a human-readable Role (Administrator, Developer, etc.).
      4. Extracts permission-relevant flags from each group (Security vs Distribution,
         Domain Local / Global / Universal scope).
      5. Emits ONE row per user-group pair so the CSV is pivot/filter-ready.
      6. Writes a timestamped log alongside the CSV report.

.PARAMETER OutputPath
    Folder where the CSV report, log file, and summary are written.
    Defaults to the script's directory, or the current working directory if
    the script is run from a console session without being saved to disk first.

.PARAMETER MaxNestDepth
    Maximum recursion depth for nested group resolution. Default: 8.

.PARAMETER PermissionGroupsOnly
    When set, only rows where the group is a Security group are included —
    useful for pure permission-tracking audits.

.EXAMPLE
    # Full report — all active users, all groups
    .\Get-ADUsersWithRoles.ps1

    # Save to a custom folder, security groups only
    .\Get-ADUsersWithRoles.ps1 -OutputPath "C:\Audit\2025" -PermissionGroupsOnly

    # Limit nested group depth to 3
    .\Get-ADUsersWithRoles.ps1 -MaxNestDepth 3

.NOTES
    Domain      : euro.net.intra
    Target OU   : OU=Users & Clients,OU=SE,OU=CRDF,DC=euro,DC=net,DC=intra
    Requires    : RSAT ActiveDirectory module  (Windows 10/11: Settings > Optional Features > RSAT)
    Permissions : Domain Read (standard user account is sufficient for most environments)
    Version     : 3.1
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = '',

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 15)]
    [int]$MaxNestDepth = 8,

    [Parameter(Mandatory = $false)]
    [switch]$PermissionGroupsOnly
)

# ── Resolve OutputPath safely ──────────────────────────────────────────────────
# $PSScriptRoot is empty when the script is run directly from a console session
# or copy-pasted into ISE/VSCode without being saved first.
# Resolution order: (1) explicit -OutputPath arg, (2) $PSScriptRoot if populated,
# (3) $PWD (current working directory) as a guaranteed non-empty fallback.
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
}

if (-not (Test-Path $OutputPath -PathType Container)) {
    throw "OutputPath '$OutputPath' does not exist or is not a directory."
}

# ── SearchScope is intentionally NOT a parameter ───────────────────────────────
# Scope is hard-locked to OneLevel so the query is strictly limited to
# OU=Users & Clients,OU=SE,OU=CRDF,DC=euro,DC=net,DC=intra
# and never descends into any child OUs beneath it.
$SearchScope = 'OneLevel'

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ══════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ══════════════════════════════════════════════════════════════════
$Script:Config = @{
    # ── Target OU (euro.net.intra > CRDF > SE > Users & Clients) ─────────────
    # Scope is OneLevel — strictly this OU only, no child OUs are crawled.
    TargetOU        = 'OU=Users & Clients,OU=SE,OU=CRDF,DC=euro,DC=net,DC=intra'
    Domain          = 'euro.net.intra'

    # ── Output files ──
    ReportName      = 'AD_ActiveUsers_Roles_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    LogName         = 'AD_ActiveUsers_Roles_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    SummaryName     = 'AD_ActiveUsers_Roles_{0}_Summary.txt' -f (Get-Date -Format 'yyyyMMdd_HHmmss')

    # ── CSV options ──
    Delimiter       = ','
    Encoding        = 'UTF8'
}

$Script:ReportPath  = Join-Path -Path $OutputPath -ChildPath $Script:Config.ReportName
$Script:LogPath     = Join-Path -Path $OutputPath -ChildPath $Script:Config.LogName
$Script:SummaryPath = Join-Path -Path $OutputPath -ChildPath $Script:Config.SummaryName


# ══════════════════════════════════════════════════════════════════
#  LOGGING
# ══════════════════════════════════════════════════════════════════
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')]
        [string]$Level = 'INFO'
    )

    $Ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $Entry = "[$Ts][$Level] $Message"

    $Colour = switch ($Level) {
        'INFO'    { 'Cyan'    }
        'WARN'    { 'Yellow'  }
        'ERROR'   { 'Red'     }
        'SUCCESS' { 'Green'   }
        'DEBUG'   { 'DarkGray'}
    }

    Write-Host $Entry -ForegroundColor $Colour
    Add-Content -Path $Script:LogPath -Value $Entry -Encoding UTF8
}


# ══════════════════════════════════════════════════════════════════
#  PREREQUISITES
# ══════════════════════════════════════════════════════════════════
function Test-Prerequisites {
    Write-Log 'Verifying prerequisites...' INFO

    # Module check
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        Write-Log ('ActiveDirectory module not found. Install via: ' +
                   'Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0') ERROR
        throw 'Missing required module: ActiveDirectory'
    }

    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log 'ActiveDirectory module imported.' SUCCESS

    # Domain connectivity
    try {
        $null = Get-ADDomain -Identity $Script:Config.Domain -ErrorAction Stop
        Write-Log "Domain reachable: $($Script:Config.Domain)" SUCCESS
    }
    catch {
        Write-Log "Cannot reach domain '$($Script:Config.Domain)': $_" ERROR
        throw
    }

    # OU existence
    try {
        $null = Get-ADOrganizationalUnit -Identity $Script:Config.TargetOU -ErrorAction Stop
        Write-Log "Target OU confirmed: $($Script:Config.TargetOU)" SUCCESS
    }
    catch {
        Write-Log "Target OU not found or inaccessible: $($Script:Config.TargetOU)" ERROR
        throw "OU validation failed: $_"
    }
}


# ══════════════════════════════════════════════════════════════════
#  ROLE CLASSIFIER  (extend the map to match your naming conventions)
# ══════════════════════════════════════════════════════════════════
function Resolve-RoleFromGroupName {
    param ([string]$GroupName)

    # Ordered so more-specific patterns win over generic ones
    $RoleMap = [ordered]@{
        # ── Privileged / Admin ────────────────────────────
        'Domain Admins'          = 'Domain Administrator'
        'Enterprise Admins'      = 'Enterprise Administrator'
        'Schema Admins'          = 'Schema Administrator'
        'Administrators'         = 'Local Administrator'
        'Admin'                  = 'Administrator'
        'Admins'                 = 'Administrator'
        'Privileged'             = 'Privileged User'
        'SysAdmin'               = 'System Administrator'

        # ── IT & Infrastructure ───────────────────────────
        'Helpdesk'               = 'Help Desk'
        'Help Desk'              = 'Help Desk'
        'Service Desk'           = 'Help Desk'
        'IT'                     = 'IT Staff'
        'Network'                = 'Network Engineer'
        'Infrastructure'         = 'Infrastructure Engineer'
        'Engineer'               = 'Engineer'
        'Operator'               = 'Operator'
        'Backup'                 = 'Backup Operator'
        'Print'                  = 'Print Operator'
        'Server'                 = 'Server Operator'
        'Monitor'                = 'Systems Monitor'

        # ── Security & Audit ─────────────────────────────
        'Security'               = 'Security'
        'Audit'                  = 'Auditor'
        'Compliance'             = 'Compliance Officer'
        'SOC'                    = 'SOC Analyst'

        # ── Development ───────────────────────────────────
        'Developer'              = 'Developer'
        'Dev'                    = 'Developer'
        'DevOps'                 = 'DevOps Engineer'
        'QA'                     = 'QA Engineer'
        'Test'                   = 'Tester'
        'Release'                = 'Release Manager'

        # ── Management ────────────────────────────────────
        'Manager'                = 'Manager'
        'Director'               = 'Director'
        'Executive'              = 'Executive'
        'CXO'                    = 'C-Level Executive'
        'VP'                     = 'Vice President'

        # ── Business Functions ────────────────────────────
        'Finance'                = 'Finance'
        'Accounting'             = 'Accounting'
        'HR'                     = 'Human Resources'
        'Legal'                  = 'Legal'
        'Marketing'              = 'Marketing'
        'Sales'                  = 'Sales'
        'Procurement'            = 'Procurement'
        'Operations'             = 'Operations'
        'Logistics'              = 'Logistics'
        'Analyst'                = 'Analyst'

        # ── Access / Remote ───────────────────────────────
        'VPN'                    = 'VPN Access'
        'Remote'                 = 'Remote Access'
        'RDP'                    = 'Remote Desktop Access'
        'MFA'                    = 'MFA Enrolled'
        'Password'               = 'Self-Service Password Reset'

        # ── Read / View ───────────────────────────────────
        'ReadOnly'               = 'Read-Only Access'
        'Read Only'              = 'Read-Only Access'
        'Viewer'                 = 'Viewer'
        'Reporter'               = 'Reporter'

        # ── Power Users ───────────────────────────────────
        'Power Users'            = 'Power User'
        'Power'                  = 'Power User'

        # ── Guest / External ──────────────────────────────
        'Guest'                  = 'Guest'
        'Contractor'             = 'Contractor'
        'External'               = 'External User'
        'Vendor'                 = 'Vendor'

        # ── Application specific ──────────────────────────
        'SharePoint'             = 'SharePoint User'
        'Exchange'               = 'Exchange User'
        'Teams'                  = 'Microsoft Teams User'
        'Outlook'                = 'Outlook User'
        'SAP'                    = 'SAP User'
        'CRM'                    = 'CRM User'
        'ERP'                    = 'ERP User'
    }

    foreach ($Pattern in $RoleMap.Keys) {
        if ($GroupName -match [regex]::Escape($Pattern)) {
            return $RoleMap[$Pattern]
        }
    }

    return 'Standard User'
}


# ══════════════════════════════════════════════════════════════════
#  PERMISSION LEVEL CLASSIFIER
#  Maps role → a numeric tier for easy sorting / alerting
# ══════════════════════════════════════════════════════════════════
function Resolve-PermissionTier {
    param ([string]$Role)

    switch -Regex ($Role) {
        'Domain Administrator|Enterprise Administrator|Schema Administrator' { return 'Tier 0 — Critical' }
        'Administrator|Privileged|System Administrator'                       { return 'Tier 1 — High'     }
        'Security|Auditor|Compliance|SOC'                                     { return 'Tier 2 — Elevated' }
        'Manager|Director|Executive|VP|C-Level'                               { return 'Tier 2 — Elevated' }
        'IT Staff|Help Desk|Network|Engineer|DevOps|Developer'                { return 'Tier 3 — Standard IT' }
        'Operator|Backup Operator|Print Operator|Server Operator'             { return 'Tier 3 — Standard IT' }
        'Finance|Accounting|HR|Legal|Procurement'                             { return 'Tier 4 — Business'    }
        'VPN Access|Remote Access|Remote Desktop Access'                      { return 'Tier 4 — Business'    }
        'Read-Only Access|Viewer|Reporter'                                    { return 'Tier 5 — Read-Only'   }
        'Guest|Contractor|External|Vendor'                                    { return 'Tier 6 — External'    }
        default                                                               { return 'Tier 5 — Standard'    }
    }
}


# ══════════════════════════════════════════════════════════════════
#  NESTED GROUP RESOLVER
# ══════════════════════════════════════════════════════════════════
function Get-AllGroupMemberships {
    param (
        [string]$UserDN,
        [int]$MaxDepth
    )

    $Visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $Result  = [System.Collections.Generic.List[PSCustomObject]]::new()

    function Recurse {
        param ([string]$ObjectDN, [int]$CurrentDepth)

        if ($CurrentDepth -ge $MaxDepth) {
            Write-Log "Max nest depth ($MaxDepth) reached at: $ObjectDN" DEBUG
            return
        }

        try {
            $ADObj = Get-ADObject -Identity $ObjectDN `
                                  -Properties MemberOf `
                                  -ErrorAction SilentlyContinue
            if (-not $ADObj -or -not $ADObj.MemberOf) { return }

            foreach ($ParentDN in $ADObj.MemberOf) {
                if (-not $Visited.Add($ParentDN)) { continue }   # already visited

                try {
                    $Grp = Get-ADGroup -Identity $ParentDN `
                                       -Properties Description, GroupCategory, GroupScope, ManagedBy, Info `
                                       -ErrorAction SilentlyContinue
                    if (-not $Grp) { continue }

                    # Resolve group manager name
                    $GrpManagerName = ''
                    if ($Grp.ManagedBy) {
                        try {
                            $GMgr = Get-ADObject -Identity $Grp.ManagedBy `
                                                 -Properties DisplayName `
                                                 -ErrorAction SilentlyContinue
                            $GrpManagerName = $GMgr.DisplayName
                        } catch { }
                    }

                    $ResolvedRole = Resolve-RoleFromGroupName -GroupName $Grp.Name
                    $PermTier     = Resolve-PermissionTier   -Role $ResolvedRole

                    $Result.Add([PSCustomObject]@{
                        GroupName         = $Grp.Name
                        GroupSamAccount   = $Grp.SamAccountName
                        GroupDN           = $Grp.DistinguishedName
                        GroupCategory     = $Grp.GroupCategory.ToString()   # Security | Distribution
                        GroupScope        = $Grp.GroupScope.ToString()      # DomainLocal | Global | Universal
                        GroupDescription  = $Grp.Description
                        GroupManagedBy    = $GrpManagerName
                        GroupNotes        = $Grp.Info
                        MembershipType    = if ($CurrentDepth -eq 0) { 'Direct' } else { "Nested (L$CurrentDepth)" }
                        NestDepth         = $CurrentDepth
                        AssignedRole      = $ResolvedRole
                        PermissionTier    = $PermTier
                        IsSecurityGroup   = ($Grp.GroupCategory -eq 'Security')
                    })

                    # Recurse into this group's parent groups
                    Recurse -ObjectDN $ParentDN -CurrentDepth ($CurrentDepth + 1)
                }
                catch {
                    Write-Log "Could not resolve group '$ParentDN': $_" WARN
                }
            }
        }
        catch {
            Write-Log "Error processing object '$ObjectDN': $_" WARN
        }
    }

    Recurse -ObjectDN $UserDN -CurrentDepth 0
    return $Result
}


# ══════════════════════════════════════════════════════════════════
#  FETCH ACTIVE USERS FROM TARGET OU
# ══════════════════════════════════════════════════════════════════
function Get-ActiveUsers {
    Write-Log "Querying active users (Scope: $SearchScope) from:" INFO
    Write-Log "  $($Script:Config.TargetOU)" INFO

    $Props = @(
        # Identity
        'employeeID',
        # Name — GivenName & Surname are default properties, listed for clarity
        'GivenName', 'Surname',
        # Contact
        'EmailAddress',
        # Account status
        'Enabled', 'SmartcardLogonRequired', 'LastLogonDate',
        # Groups — required for membership resolution
        'MemberOf',
        # DistinguishedName required internally for nested group traversal
        'DistinguishedName'
    )

    try {
        $Users = Get-ADUser `
            -SearchBase  $Script:Config.TargetOU `
            -SearchScope $SearchScope `
            -Filter      { Enabled -eq $true } `
            -Properties  $Props `
            -ErrorAction Stop

        Write-Log "Found $($Users.Count) active user(s)." SUCCESS
        return $Users
    }
    catch {
        Write-Log "Failed to query Active Directory: $_" ERROR
        throw
    }
}


# ══════════════════════════════════════════════════════════════════
#  BUILD REPORT ROWS
# ══════════════════════════════════════════════════════════════════
function Build-Report {
    param ([object[]]$Users)

    $Report  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $Counter = 0
    $Total   = $Users.Count

    foreach ($User in $Users) {
        $Counter++
        $Pct = [math]::Round(($Counter / $Total) * 100, 1)
        Write-Progress -Activity 'Building Permission Report' `
                       -Status   "[$Counter / $Total]  $($User.SamAccountName)" `
                       -PercentComplete $Pct

        Write-Log "Processing ($Counter/$Total): $($User.SamAccountName)" INFO

        # ── Group memberships (direct + nested) ───────────────────────
        $AllGroups = Get-AllGroupMemberships -UserDN $User.DistinguishedName -MaxDepth $MaxNestDepth

        # Apply PermissionGroupsOnly filter if requested
        if ($PermissionGroupsOnly) {
            $AllGroups = @($AllGroups | Where-Object { $_.IsSecurityGroup })
        }

        # Determine the primary (highest-tier) role across all groups
        $PrimaryRole = 'Standard User'
        $PrimaryTier = 'Tier 5 — Standard'
        foreach ($G in ($AllGroups | Sort-Object NestDepth)) {
            if ($G.AssignedRole -ne 'Standard User') {
                $PrimaryRole = $G.AssignedRole
                $PrimaryTier = $G.PermissionTier
                break
            }
        }

        # Direct group names (for quick-scan column)
        $DirectGroupNames = @($AllGroups | Where-Object { $_.MembershipType -eq 'Direct' } |
                              Select-Object -ExpandProperty GroupName)
        $DirectGroupCount = $DirectGroupNames.Count
        $DirectGroupList  = $DirectGroupNames -join ' | '

        # ── Shared user-level fields (written once per group row) ─────
        $UserBase = [ordered]@{
            # Identity
            EmployeeID               = $User.employeeID

            # Name
            GivenName                = $User.GivenName
            Surname                  = $User.Surname

            # Contact
            EmailAddress             = $User.EmailAddress

            # Account status
            AccountStatus            = 'Active'
            Enabled                  = $User.Enabled
            SmartcardRequired        = $User.SmartcardLogonRequired
            LastLogonDate            = $User.LastLogonDate

            # Role summary (user level)
            PrimaryRole              = $PrimaryRole
            TotalGroupCount          = $AllGroups.Count
            DirectGroupCount         = $DirectGroupCount
            DirectGroupList          = $DirectGroupList
        }

        # ── Emit one row per group (or one "no groups" row) ───────────
        if ($AllGroups.Count -eq 0) {
            $Row = [ordered]@{}
            foreach ($K in $UserBase.Keys) { $Row[$K] = $UserBase[$K] }
            $Row['GroupName']        = ''
            $Row['GroupDescription'] = ''
            $Row['GroupManagedBy']   = ''
            $Row['AssignedRole']     = 'No Groups Assigned'
            $Report.Add([PSCustomObject]$Row)
        }
        else {
            foreach ($Group in ($AllGroups | Sort-Object NestDepth, GroupName)) {
                $Row = [ordered]@{}
                foreach ($K in $UserBase.Keys) { $Row[$K] = $UserBase[$K] }
                $Row['GroupName']        = $Group.GroupName
                $Row['GroupDescription'] = $Group.GroupDescription
                $Row['GroupManagedBy']   = $Group.GroupManagedBy
                $Row['AssignedRole']     = $Group.AssignedRole
                $Report.Add([PSCustomObject]$Row)
            }
        }
    }

    Write-Progress -Activity 'Building Permission Report' -Completed
    return $Report
}


# ══════════════════════════════════════════════════════════════════
#  CSV EXPORT
# ══════════════════════════════════════════════════════════════════
function Export-CsvReport {
    param ([System.Collections.Generic.List[PSCustomObject]]$Report)

    Write-Log "Exporting $($Report.Count) rows → $Script:ReportPath" INFO

    $Report | Export-Csv `
        -Path             $Script:ReportPath `
        -Delimiter        $Script:Config.Delimiter `
        -Encoding         $Script:Config.Encoding `
        -NoTypeInformation `
        -ErrorAction      Stop

    Write-Log "CSV report saved: $Script:ReportPath" SUCCESS
}


# ══════════════════════════════════════════════════════════════════
#  SUMMARY REPORT
# ══════════════════════════════════════════════════════════════════
function Write-Summary {
    param ([System.Collections.Generic.List[PSCustomObject]]$Report)

    # Unique users keyed on EmployeeID (falls back to GivenName+Surname pair)
    $UniqueUsers   = ($Report | Select-Object -Unique EmployeeID).Count
    $TotalRows     = $Report.Count
    $NoGroupUsers  = ($Report | Where-Object { $_.AssignedRole -eq 'No Groups Assigned' } |
                      Select-Object -Unique EmployeeID).Count
    $NeverLoggedIn = ($Report | Where-Object { -not $_.LastLogonDate } |
                      Select-Object -Unique EmployeeID).Count
    $SmartcardOn   = ($Report | Where-Object { $_.SmartcardRequired -eq $true } |
                      Select-Object -Unique EmployeeID).Count

    $RoleBreakdown = $Report | Group-Object AssignedRole |
                     Sort-Object Count -Descending | Select-Object Name, Count

    $TopGroups     = $Report | Where-Object { $_.GroupName -ne '' } |
                     Group-Object GroupName | Sort-Object Count -Descending |
                     Select-Object -First 15 Name, Count

    $Lines = @(
        '═══════════════════════════════════════════════════════════════════'
        '   ACTIVE DIRECTORY — USER PERMISSION REPORT SUMMARY'
        '═══════════════════════════════════════════════════════════════════'
        "  Generated         : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "  Domain            : $($Script:Config.Domain)"
        "  Source OU         : $($Script:Config.TargetOU)"
        "  Search Scope      : OneLevel (strictly this OU only — no child OUs)"
        "  Max Nest Depth    : $MaxNestDepth"
        "  Security Only     : $($PermissionGroupsOnly.IsPresent)"
        ''
        '─── USER COUNTS ────────────────────────────────────────────────────'
        "  Unique Active Users   : $UniqueUsers"
        "  Total Report Rows     : $TotalRows"
        "  Users with No Groups  : $NoGroupUsers"
        "  Never Logged In       : $NeverLoggedIn"
        "  Smartcard Required    : $SmartcardOn"
        ''
        '─── TOP 15 GROUPS BY MEMBERSHIP ────────────────────────────────────'
    )
    foreach ($G in $TopGroups) {
        $Lines += "  {0,-50} : {1}" -f $G.Name, $G.Count
    }

    $Lines += ''
    $Lines += '─── ASSIGNED ROLE BREAKDOWN ────────────────────────────────────────'
    foreach ($R in $RoleBreakdown) {
        $Lines += "  {0,-40} : {1}" -f $R.Name, $R.Count
    }

    $Lines += ''
    $Lines += '─── OUTPUT FILES ───────────────────────────────────────────────────'
    $Lines += "  CSV Report   : $Script:ReportPath"
    $Lines += "  Log File     : $Script:LogPath"
    $Lines += "  Summary File : $Script:SummaryPath"
    $Lines += '═══════════════════════════════════════════════════════════════════'

    $Lines | ForEach-Object { Write-Host $_ -ForegroundColor Magenta }
    $Lines | Out-File -FilePath $Script:SummaryPath -Encoding UTF8
    Write-Log "Summary written: $Script:SummaryPath" SUCCESS
}


# ══════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════
function Main {
    Write-Log '══════════════════════════════════════════════════' INFO
    Write-Log ' AD Active Users — Permission & Role Report v3.1  ' INFO
    Write-Log "  Domain    : $($Script:Config.Domain)"            INFO
    Write-Log "  Target OU : $($Script:Config.TargetOU)"          INFO
    Write-Log "  Output    : $OutputPath"                         INFO
    Write-Log '══════════════════════════════════════════════════' INFO

    Test-Prerequisites

    $Users  = Get-ActiveUsers
    $Report = Build-Report -Users $Users

    Export-CsvReport -Report $Report
    Write-Summary    -Report $Report

    Write-Log 'Script completed successfully.' SUCCESS
}

# ══════════════════════════════════════════════════════════════════
#  RUN
# ══════════════════════════════════════════════════════════════════
try {
    Main
}
catch {
    Write-Log "FATAL ERROR: $_"               ERROR
    Write-Log $_.ScriptStackTrace             ERROR
    exit 1
}
