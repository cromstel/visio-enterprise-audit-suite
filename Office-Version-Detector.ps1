<#
.SYNOPSIS
    Detects Microsoft Office installations and identifies if Office 365 or Office 2019 is installed.

.DESCRIPTION
    This script performs version detection for Microsoft Office installations by checking
    registry keys for both Click-to-Run (C2R) and Windows Installer (MSI) installations.
    It specifically identifies Office 365 and Office 2019 installations while rejecting
    all other versions (Office 2016, 2013, 2010, etc.).

    The script supports:
    - Office 365/Microsoft 365 detection via Click-to-Run
    - Office 2019 detection via Click-to-Run and MSI
    - Both 32-bit and 64-bit system detection
    - Detailed logging to console and file

.PARAMETER LogFilePath
    Specifies the path for the log file. Default is ".\Office-Version-Detection.log"

.PARAMETER StrictErrorHandling
    Enables strict error handling mode. When enabled, the script will terminate on
    any non-critical errors instead of continuing with degraded functionality.

.PARAMETER VerboseLogging
    Enables verbose logging output to console. When enabled, additional debug
    information will be displayed during execution.

.EXAMPLE
    .\Office-Version-Detector.ps1

    Runs the script with default parameters and displays detection results.

.EXAMPLE
    .\Office-Version-Detector.ps1 -LogFilePath "C:\Logs\OfficeDetection.log" -VerboseLogging

    Runs the script with custom log file path and verbose logging enabled.

.EXAMPLE
    .\Office-Version-Detector.ps1 -StrictErrorHandling

    Runs the script with strict error handling enabled for production environments.

.OUTPUTS
    Console output indicating detected Office version and support status.
    Log file with detailed execution information.
    Exit codes: 0 (success - supported version), 1 (unsupported version or error)

.NOTES
    File Name      : Office-Version-Detector.ps1
    Prerequisite   : PowerShell 5.1 or later
    Author         : Visio Enterprise Audit Suite
    Version        : 1.0.0
    Last Modified  : 2026
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$LogFilePath = ".\Office-Version-Detection.log",
    
    [Parameter(Mandatory = $false)]
    [switch]$StrictErrorHandling = $false,
    
    [Parameter(Mandatory = $false)]
    [switch]$VerboseLogging = $false
)

#region Global Variables

# Script-wide variables for tracking state
[string]$script:DetectedVersion = $null
[string]$script:DetectedProductName = $null
[string]$script:DetectionMethod = $null
[bool]$script:DetectionSuccessful = $false

# Registry paths for Office detection
[string]$script:RegistryPathClickToRun = "SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
[string]$script:RegistryPathOffice16 = "SOFTWARE\Microsoft\Office\16.0\Configuration"
[string]$script:RegistryPathUninstall = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
[string]$script:RegistryPathWow6432 = "SOFTWARE\Wow6432Node\Microsoft\Office\ClickToRun\Configuration"

# Supported Office version identifiers
[string[]]$script:Office365Identifiers = @("O365", "Microsoft 365", "365")
[string[]]$script:Office2019Identifiers = @("2019")

#endregion

#region Helper Functions

<#
.SYNOPSIS
    Writes a timestamped log message to console and log file.

.DESCRIPTION
    Internal helper function that formats and outputs log messages with
    timestamp, log level, and function name for traceability.

.PARAMETER Message
    The log message to write.

.PARAMETER Level
    The log level (INFO, WARNING, ERROR, DEBUG).

.PARAMETER FunctionName
    The name of the function calling this logging helper.

.EXAMPLE
    Write-LogMessage -Message "Starting detection process" -Level INFO -FunctionName "Start-Detection"
#>
function Write-LogMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory = $false)]
        [string]$FunctionName = "Global"
    )
    
    try {
        # Create timestamp in consistent format
        [string]$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        # Format the log entry
        [string]$logEntry = "[$timestamp] [$Level] [$FunctionName] $Message"
        
        # Write to console with color coding based on level
        switch ($Level) {
            "ERROR" {
                Write-Host $logEntry -ForegroundColor Red
            }
            "WARNING" {
                Write-Host $logEntry -ForegroundColor Yellow
            }
            "DEBUG" {
                if ($VerboseLogging -or $script:VerboseLoggingMode) {
                    Write-Host $logEntry -ForegroundColor Cyan
                }
            }
            default {
                Write-Host $logEntry
            }
        }
        
        # Write to log file (suppress errors if file is locked)
        try {
            Add-Content -Path $LogFilePath -Value $logEntry -ErrorAction Stop
        }
        catch {
            # Silently continue if logging fails - don't let logging failure break execution
            Write-Warning "Failed to write to log file: $($_.Exception.Message)"
        }
    }
    catch {
        # Final fallback - ensure no exception escapes
        Write-Host "[LOG ERROR] $Message" -ForegroundColor Red
    }
}

<#
.SYNOPSIS
    Reads a value from the Windows registry with error handling.

.DESCRIPTION
    Centralized registry access function that handles all registry reads
    with consistent error handling and logging.

.PARAMETER Hive
    The registry hive (e.g., "HKLM", "HKCU").

.PARAMETER KeyPath
    The full path to the registry key.

.PARAMETER ValueName
    The name of the value to read.

.OUTPUTS
    The registry value if found, or $null if not found or error occurs.
#>
function Read-RegistryValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("HKLM", "HKCU", "HKCR", "HKU")]
        [string]$Hive,
        
        [Parameter(Mandatory = $true)]
        [string]$KeyPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ValueName
    )
    
    try {
        Write-LogMessage -Message "Reading registry value: $Hive\$KeyPath\$ValueName" -Level DEBUG -FunctionName "Read-RegistryValue"
        
        # Build full registry path using subexpression to handle registry drive colon
        [string]$fullPath = "$Hive`:\$KeyPath"
        
        # Check if key exists before reading
        if (-not (Test-Path -Path $fullPath -ErrorAction Stop)) {
            Write-LogMessage -Message "Registry key does not exist: $fullPath" -Level DEBUG -FunctionName "Read-RegistryValue"
            return $null
        }
        
        # Attempt to read the value
        [object]$value = Get-ItemProperty -Path $fullPath -Name $ValueName -ErrorAction Stop
        
        if ($null -eq $value) {
            Write-LogMessage -Message "Registry value is null: $Hive\$KeyPath\$ValueName" -Level DEBUG -FunctionName "Read-RegistryValue"
            return $null
        }
        
        [string]$result = $value.$ValueName
        Write-LogMessage -Message "Registry value retrieved: $result" -Level DEBUG -FunctionName "Read-RegistryValue"
        
        return $result
    }
    catch [System.Management.Automation.ItemNotFoundException] {
        Write-LogMessage -Message "Registry key not found: $Hive\$KeyPath" -Level DEBUG -FunctionName "Read-RegistryValue"
        return $null
    }
    catch [System.Security.SecurityException] {
        Write-LogMessage -Message "Registry access denied: $Hive\$KeyPath - $($_.Exception.Message)" -Level WARNING -FunctionName "Read-RegistryValue"
        return $null
    }
    catch {
        Write-LogMessage -Message "Registry read error: $Hive\$KeyPath - $($_.Exception.Message)" -Level WARNING -FunctionName "Read-RegistryValue"
        return $null
    }
}

<#
.SYNOPSIS
    Checks if a registry key exists with error handling.

.DESCRIPTION
    Verifies the existence of a registry key before attempting to read from it.

.PARAMETER Hive
    The registry hive.

.PARAMETER KeyPath
    The full path to the registry key.

.OUTPUTS
    $true if the key exists, $false otherwise.
#>
function Test-RegistryKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("HKLM", "HKCU", "HKCR", "HKU")]
        [string]$Hive,
        
        [Parameter(Mandatory = $true)]
        [string]$KeyPath
    )
    
    try {
        [string]$fullPath = "$Hive`:\$KeyPath"
        return Test-Path -Path $fullPath -ErrorAction Stop
    }
    catch {
        return $false
    }
}

<#
.SYNOPSIS
    Detects Office installation using Click-to-Run registry keys.

.DESCRIPTION
    Checks Office 365/2019 Click-to-Run installation paths in the registry
    and extracts version and product information.

.OUTPUTS
    Hash table with detected version and product name, or $null if not detected.
#>
function Test-ClickToRunInstallation {
    [CmdletBinding()]
    param()
    
    Write-LogMessage -Message "Checking Click-to-Run installation detection" -Level INFO -FunctionName "Test-ClickToRunInstallation"
    
    # Define registry paths to check for Click-to-Run
    [string[]]$clickToRunPaths = @(
        "$script:RegistryPathClickToRun",           # Standard 64-bit
        "$script:RegistryPathWow6432\$script:RegistryPathClickToRun"  # 32-bit on 64-bit
    )
    
    foreach ($path in $clickToRunPaths) {
        try {
            # Check for required Click-to-Run values
            [string]$version = Read-RegistryValue -Hive "HKLM" -KeyPath $path -ValueName "Version"
            [string]$productName = Read-RegistryValue -Hive "HKLM" -KeyPath $path -ValueName "ProductName"
            [string]$updateChannel = Read-RegistryValue -Hive "HKLM" -KeyPath $path -ValueName "UpdateChannel"
            
            # If version is found, this is a valid Click-to-Run installation
            if (-not [string]::IsNullOrEmpty($version)) {
                Write-LogMessage -Message "Found Click-to-Run installation at: $path" -Level DEBUG -FunctionName "Test-ClickToRunInstallation"
                Write-LogMessage -Message "Version: $version, ProductName: $productName, UpdateChannel: $updateChannel" -Level DEBUG -FunctionName "Test-ClickToRunInstallation"
                
                # Return the detected values
                @{
                    Version = $version
                    ProductName = $productName
                    UpdateChannel = $updateChannel
                    InstallationType = "Click-to-Run"
                    RegistryPath = $path
                }
                return
            }
        }
        catch {
            Write-LogMessage -Message "Error checking Click-to-Run path $path : $($_.Exception.Message)" -Level DEBUG -FunctionName "Test-ClickToRunInstallation"
            continue
        }
    }
    
    Write-LogMessage -Message "No Click-to-Run installation detected" -Level DEBUG -FunctionName "Test-ClickToRunInstallation"
    return $null
}

<#
.SYNOPSIS
    Detects Office installation using MSI registry keys.

.DESCRIPTION
    Checks Office MSI installation paths and the Windows Uninstall registry
    for Office installations.

.OUTPUTS
    Hash table with detected version and product name, or $null if not detected.
#>
function Test-MSIInstallation {
    [CmdletBinding()]
    param()
    
    Write-LogMessage -Message "Checking MSI installation detection" -Level INFO -FunctionName "Test-MSIInstallation"
    
    # Check Office 16.0 configuration key (used by Office 2019 MSI)
    try {
        [string]$version = Read-RegistryValue -Hive "HKLM" -KeyPath $script:RegistryPathOffice16 -ValueName "Version"
        [string]$productName = Read-RegistryValue -Hive "HKLM" -KeyPath $script:RegistryPathOffice16 -ValueName "ProductName"
        
        if (-not [string]::IsNullOrEmpty($version)) {
            Write-LogMessage -Message "Found MSI installation at: $script:RegistryPathOffice16" -Level DEBUG -FunctionName "Test-MSIInstallation"
            
            @{
                Version = $version
                ProductName = $productName
                InstallationType = "MSI"
                RegistryPath = $script:RegistryPathOffice16
            }
            return
        }
    }
    catch {
        Write-LogMessage -Message "Error checking MSI path $script:RegistryPathOffice16 : $($_.Exception.Message)" -Level DEBUG -FunctionName "Test-MSIInstallation"
    }
    
    # Check Windows Uninstall registry for Office entries
    try {
        [string]$uninstallPath = "$script:RegistryPathUninstall"
        [string]$uninstallPathWow = "Wow6432Node\$script:RegistryPathUninstall"
        
        # Check both 64-bit and 32-bit uninstall keys
        foreach ($basePath in @($uninstallPath, $uninstallPathWow)) {
            if (Test-RegistryKey -Hive "HKLM" -KeyPath $basePath) {
                # Get all uninstall entries
                [string[]]$subKeys = Get-ChildItem -Path "HKLM`:\$basePath" -ErrorAction SilentlyContinue | 
                    ForEach-Object { $_.PSChildName }
                
                foreach ($subKey in $subKeys) {
                    # Skip non-GUID entries and Office-related entries
                    if ($subKey -match '^{[A-F0-9-]{36}}$') {
                        try {
                            [string]$displayName = Read-RegistryValue -Hive "HKLM" -KeyPath "$basePath\$subKey" -ValueName "DisplayName"
                            [string]$displayVersion = Read-RegistryValue -Hive "HKLM" -KeyPath "$basePath\$subKey" -ValueName "DisplayVersion"
                            
                            if ($displayName -match "Office" -or $displayName -match "Microsoft 365") {
                                Write-LogMessage -Message "Found Office entry in Uninstall: $displayName version $displayVersion" -Level DEBUG -FunctionName "Test-MSIInstallation"
                                
                                @{
                                    Version = $displayVersion
                                    ProductName = $displayName
                                    InstallationType = "MSI-Uninstall"
                                    RegistryPath = "$basePath\$subKey"
                                }
                                return
                            }
                        }
                        catch {
                            continue
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-LogMessage -Message "Error checking Uninstall registry: $($_.Exception.Message)" -Level DEBUG -FunctionName "Test-MSIInstallation"
    }
    
    Write-LogMessage -Message "No MSI installation detected" -Level DEBUG -FunctionName "Test-MSIInstallation"
    return $null
}

<#
.SYNOPSIS
    Validates and categorizes detected Office version.

.DESCRIPTION
    Analyzes the detected version and product name to determine if it's
    a supported version (Office 365 or Office 2019) or unsupported.

.PARAMETER Version
    The detected version string.

.PARAMETER ProductName
    The detected product name string.

.OUTPUTS
    Hash table with validation results including supported status and category.
#>
function Test-OfficeVersionValidation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Version,
        
        [Parameter(Mandatory = $false)]
        [string]$ProductName = ""
    )
    
    Write-LogMessage -Message "Validating Office version: $Version, Product: $ProductName" -Level INFO -FunctionName "Test-OfficeVersionValidation"
    
    # Initialize result
    $result = @{
        IsSupported = $false
        OfficeType = "Unknown"
        Version = $Version
        ProductName = $ProductName
    }
    
    try {
        # Check if version starts with 16.0 (Office 2016/2019/365 family)
        if ($Version -notmatch '^16\.0\.') {
            Write-LogMessage -Message "Version $Version is not in Office 2016+ family (must start with 16.0)" -Level WARNING -FunctionName "Test-OfficeVersionValidation"
            return $result
        }
        
        # Check for Office 365 identifiers in product name or version string
        foreach ($identifier in $script:Office365Identifiers) {
            if ($ProductName -match $identifier -or $Version -match $identifier -or $ProductName -match "Microsoft 365") {
                $result.IsSupported = $true
                $result.OfficeType = "Office 365"
                Write-LogMessage -Message "Identified as Office 365" -Level INFO -FunctionName "Test-OfficeVersionValidation"
                return $result
            }
        }
        
        # Check for Office 2019 identifiers in product name
        foreach ($identifier in $script:Office2019Identifiers) {
            if ($ProductName -match $identifier) {
                $result.IsSupported = $true
                $result.OfficeType = "Office 2019"
                Write-LogMessage -Message "Identified as Office 2019" -Level INFO -FunctionName "Test-OfficeVersionValidation"
                return $result
            }
        }
        
        # If we reach here with 16.0 version but no specific identifier, it could be Office 2016 or other
        Write-LogMessage -Message "Version 16.0.x detected but no Office 365/2019 identifier found in: $ProductName" -Level WARNING -FunctionName "Test-OfficeVersionValidation"
        return $result
        
    }
    catch {
        Write-LogMessage -Message "Error validating Office version: $($_.Exception.Message)" -Level ERROR -FunctionName "Test-OfficeVersionValidation"
        return $result
    }
}

<#
.SYNOPSIS
    Parses version string into comparable format.

.DESCRIPTION
    Converts version string like "16.0.12345.12345" into a structured
    format for comparison and display.

.PARAMETER VersionString
    The version string to parse.

.OUTPUTS
    Array of version number integers, or $null if parsing fails.
#>
function ConvertFrom-OfficeVersionString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$VersionString
    )
    
    try {
        # Remove any extra text and split by dot
        [string]$cleanVersion = $VersionString.Trim()
        
        # Split into components
        [string[]]$parts = $cleanVersion.Split('.')
        
        if ($parts.Length -lt 2) {
            Write-LogMessage -Message "Invalid version format: $VersionString" -Level WARNING -FunctionName "ConvertFrom-OfficeVersionString"
            return $null
        }
        
        # Convert to integers
        [int[]]$versionNumbers = @()
        foreach ($part in $parts) {
            [int]$number = 0
            if ([int]::TryParse($part, [ref]$number)) {
                $versionNumbers += $number
            }
            else {
                # Handle non-numeric parts (like "0" in some version strings)
                $versionNumbers += 0
            }
        }
        
        return $versionNumbers
    }
    catch {
        Write-LogMessage -Message "Error parsing version string: $($_.Exception.Message)" -Level ERROR -FunctionName "ConvertFrom-OfficeVersionString"
        return $null
    }
}

<#
.SYNOPSIS
    Formats version array back to string for display.

.DESCRIPTION
    Converts version array back to dot-separated string format.

.PARAMETER VersionNumbers
    Array of version numbers.

.OUTPUTS
    Formatted version string.
#>
function ConvertTo-OfficeVersionString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int[]]$VersionNumbers
    )
    
    return $VersionNumbers -join '.'
}

#endregion

#region Main Detection Function

<#
.SYNOPSIS
    Main function that orchestrates Office version detection.

.DESCRIPTION
    Coordinates all detection methods and returns the final detection result
    with appropriate exit codes for the calling process.

.OUTPUTS
    Exit code: 0 for supported version, 1 for unsupported or error.
#>
function Start-OfficeDetection {
    [CmdletBinding()]
    param()
    
    Write-LogMessage -Message "========================================" -Level INFO -FunctionName "Start-OfficeDetection"
    Write-LogMessage -Message "Starting Office Version Detection" -Level INFO -FunctionName "Start-OfficeDetection"
    Write-LogMessage -Message "========================================" -Level INFO -FunctionName "Start-OfficeDetection"
    
    # Enable verbose logging if requested
    if ($VerboseLogging) {
        $script:VerboseLoggingMode = $true
        Write-LogMessage -Message "Verbose logging enabled" -Level INFO -FunctionName "Start-OfficeDetection"
    }
    
    try {
        # Initialize log file
        try {
            # Create log file header
            [string]$header = "========================================`nOffice Version Detection Log`nStarted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n========================================"
            Set-Content -Path $LogFilePath -Value $header -ErrorAction Stop
            Write-LogMessage -Message "Log file initialized: $LogFilePath" -Level INFO -FunctionName "Start-OfficeDetection"
        }
        catch {
            Write-Warning "Failed to initialize log file: $($_.Exception.Message)"
            Write-Warning "Continuing without file logging..."
        }
        
        # Reset detection state
        $script:DetectedVersion = $null
        $script:DetectedProductName = $null
        $script:DetectionMethod = $null
        $script:DetectionSuccessful = $false
        
        [object]$detectionResult = $null
        [string]$detectionSource = $null
        
        # Try Click-to-Run detection first (most common for Office 365/2019)
        Write-LogMessage -Message "Step 1: Checking Click-to-Run installations..." -Level INFO -FunctionName "Start-OfficeDetection"
        $detectionResult = Test-ClickToRunInstallation
        
        if ($null -ne $detectionResult) {
            $detectionSource = "Click-to-Run"
        }
        else {
            # Fall back to MSI detection
            Write-LogMessage -Message "Step 2: Checking MSI installations..." -Level INFO -FunctionName "Start-OfficeDetection"
            $detectionResult = Test-MSIInstallation
            
            if ($null -ne $detectionResult) {
                $detectionSource = "MSI"
            }
        }
        
        # Process detection result
        if ($null -ne $detectionResult) {
            $script:DetectedVersion = $detectionResult.Version
            $script:DetectedProductName = $detectionResult.ProductName
            $script:DetectionMethod = $detectionSource
            
            Write-LogMessage -Message "Detection successful via $detectionSource" -Level INFO -FunctionName "Start-OfficeDetection"
            Write-LogMessage -Message "Version: $script:DetectedVersion" -Level INFO -FunctionName "Start-OfficeDetection"
            Write-LogMessage -Message "Product: $script:DetectedProductName" -Level INFO -FunctionName "Start-OfficeDetection"
            
            # Validate the detected version
            $validationResult = Test-OfficeVersionValidation -Version $script:DetectedVersion -ProductName $script:DetectedProductName
            
            if ($validationResult.IsSupported) {
                # Supported Office version detected
                Write-LogMessage -Message "SUPPORTED: $($validationResult.OfficeType) detected - Version: $script:DetectedVersion" -Level INFO -FunctionName "Start-OfficeDetection"
                
                Write-Host ""
                Write-Host "========================================" -ForegroundColor Green
                Write-Host "SUPPORTED: $($validationResult.OfficeType) detected" -ForegroundColor Green
                Write-Host "Version: $script:DetectedVersion" -ForegroundColor Green
                Write-Host "Product: $script:DetectedProductName" -ForegroundColor Green
                Write-Host "Detection Method: $detectionSource" -ForegroundColor Green
                Write-Host "========================================" -ForegroundColor Green
                Write-Host ""
                
                $script:DetectionSuccessful = $true
                return 0  # Exit code 0: Success - supported version
            }
            else {
                # Unsupported Office version detected
                [string]$versionDisplay = ConvertTo-OfficeVersionString -VersionNumbers (ConvertFrom-OfficeVersionString -VersionString $script:DetectedVersion)
                
                Write-LogMessage -Message "UNSUPPORTED: Office $versionDisplay detected - This script only supports Office 365 and Office 2019" -Level WARNING -FunctionName "Start-OfficeDetection"
                
                Write-Host ""
                Write-Host "========================================" -ForegroundColor Yellow
                Write-Host "UNSUPPORTED: Office $versionDisplay detected" -ForegroundColor Yellow
                Write-Host "This script only supports Office 365 and Office 2019" -ForegroundColor Yellow
                Write-Host "Product: $script:DetectedProductName" -ForegroundColor Yellow
                Write-Host "========================================" -ForegroundColor Yellow
                Write-Host ""
                
                return 1  # Exit code 1: Unsupported version
            }
        }
        else {
            # No Office installation detected
            Write-LogMessage -Message "No Office installation detected on this system" -Level WARNING -FunctionName "Start-OfficeDetection"
            
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Red
            Write-Host "NO OFFICE DETECTED" -ForegroundColor Red
            Write-Host "No Microsoft Office installation was found." -ForegroundColor Red
            Write-Host "========================================" -ForegroundColor Red
            Write-Host ""
            
            return 1  # Exit code 1: Error - no installation found
        }
    }
    catch [System.Security.SecurityException] {
        Write-LogMessage -Message "ERROR: Registry access failed: $($_.Exception.Message)" -Level ERROR -FunctionName "Start-OfficeDetection"
        
        Write-Host ""
        Write-Host "ERROR: Registry access failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please run this script with administrator privileges." -ForegroundColor Red
        Write-Host ""
        
        if ($StrictErrorHandling) {
            throw
        }
        return 1
    }
    catch {
        Write-LogMessage -Message "ERROR: Version detection failed: $($_.Exception.Message)" -Level ERROR -FunctionName "Start-OfficeDetection"
        
        Write-Host ""
        Write-Host "ERROR: Version detection failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        
        if ($StrictErrorHandling) {
            throw
        }
        return 1
    }
    finally {
        Write-LogMessage -Message "Office detection process completed" -Level INFO -FunctionName "Start-OfficeDetection"
        Write-LogMessage -Message "========================================" -Level INFO -FunctionName "Start-OfficeDetection"
    }
}

#endregion

#region Script Execution

# Main script execution
try {
    # Set error action preference
    $ErrorActionPreference = if ($StrictErrorHandling) { "Stop" } else { "Continue" }
    
    # Execute detection
    [int]$exitCode = Start-OfficeDetection
    
    # Exit with appropriate code
    exit $exitCode
}
catch {
    Write-LogMessage -Message "FATAL ERROR: Script execution failed: $($_.Exception.Message)" -Level ERROR -FunctionName "Main"
    
    Write-Host ""
    Write-Host "FATAL ERROR: Script execution failed" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    
    exit 1
}

#endregion
