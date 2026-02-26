# AGENTS.md

This file provides guidance to agents when working with code in this repository.

## Project Overview

This is a Visio Enterprise Audit Suite, a collection of PowerShell scripts designed to audit Visio installations in enterprise environments.

## Key Files

### visio-enterprise-audit.ps1
- **Purpose**: Main audit script that scans domain computers for Visio installations
- **Features**: Active Directory integration, parallel scanning, HTML/CSV reports
- **Key Functions**:
  - Get-DomainComputers: Queries AD for computers
  - Invoke-VisioScan: Parallel scanning using runspaces
  - ConvertTo-HtmlReport: Generates HTML reports
  - Export-ResultsToCSV: Exports results to CSV
  - Get-AuditSummary: Calculates audit statistics

### visio-helper-utils.ps1
- Contains utility functions used by the main audit script

### visio-usage-analytics.ps1
- Analyzes Visio usage patterns on the computers on the domain

### Office-Version-Detector.ps1
- Detects computers on the domain with installed Office versions

## Command Line Usage

Basic usage of the main audit script:
```powershell
.\visio-enterprise-audit.ps1 -OutputPath ".\Reports" -ThreadCount 20 -ComputerPrefix "GOT"
```

## Development Notes

- Scripts require PowerShell 5.1 or later
- Active Directory module is required for domain queries
- Scripts should be run with administrative privileges
- HTML reports use modern CSS with gradient backgrounds
- CSV reports are UTF-8 encoded with user-friendly column names
-mScan for errors and typescripts before reviewing code to complete
