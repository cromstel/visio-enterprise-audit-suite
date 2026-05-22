# Balanced scan (recommended start)
.\Run-PermissionScan.ps1 -Profile Balanced -SkipBuiltIn

# Ultra-safe scan during business hours, custom report folder
.\Run-PermissionScan.ps1 -Profile Safe -ReportFolder "D:\Reports" -SkipBuiltIn

# Schedule nightly scan at 1:30 AM
.\Run-PermissionScan.ps1 -Profile Safe -ScheduleTask -TaskTime "01:30"

# Analyse the output after scanning
.\Get-PermissionSummary.ps1 -InputCsv "C:\Reports\FolderPermissions_Balanced_20240915.csv"
