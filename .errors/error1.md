
flowchart TD
    A[Visio Detection Methods] --> B[WMI Win32_Product]
    A --> C[Registry Detection]
    A --> D[File System Detection]
    A --> E[Click-to-Run Detection]
    
    B --> B1[Office 365 Apps]
    B --> B2[Visio 2021 Standard]
    B --> B3[Visio 2021 Professional]
    B --> B4[Visio Professional 2019]
    
    C --> C1[Office16 Keys]
    C --> C2[Click-to-Run Keys]
    C --> C3[Standalone MSI Keys]
    
    D --> D1[Office16 Path]
    D --> D2[Office15 Path]
    D --> D3[Program Files Visio Path]


## Visio Enterprise Audit Script - Compatibility Enhancement Plan

### Executive Summary
Your current scripts only support Office 365/2019 with hardcoded `Office16` paths. To support all three Visio versions in your environment, the following changes are required:

---

### Current Compatibility Gap

The scripts at [`visio-enterprise-audit.ps1`](visio-enterprise-audit.ps1:78-85) contain:
```powershell
$VisioPaths = @(
    "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE"  # Only Office 365/2019
)
$RegistryPaths = @(
    "HKLM:\Software\Microsoft\Office\16.0\Common\InstallRoot"
)
```

This **excludes** Visio 2021 standalone editions and may miss some configurations.

---

### Version Detection Matrix

| Visio Version | Click-to-Run Path | Registry Key | Detection Method |
|---------------|-------------------|--------------|------------------|
| **Visio 2021 Standard (365 Apps 2408)** | `\root\Office16\VISIO.EXE` | `Office\16.0` | WMI + Registry |
| **Visio 2021 Professional (365 Apps 2408)** | `\root\Office16\VISIO.EXE` | `Office\16.0` | WMI + Registry |
| **Visio Professional 2019** | `\root\Office15\VISIO.EXE` | `Office\15.0` | WMI + Registry |
| **Office 365/2021** | `\root\Office16\VISIO.EXE` | `Office\16.0` | Current (works) |

---

### Recommended Script Updates

**1. Expand Installation Paths** ([`visio-enterprise-audit.ps1`](visio-enterprise-audit.ps1:78))
```powershell
$VisioPaths = @(
    "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE",   # Office 365/2021/2019
    "C:\Program Files\Microsoft Office\root\Office15\VISIO.EXE",   # Visio 2019
    "C:\Program Files\Microsoft Office\Office16\VISIO.EXE",        # Standalone
    "C:\Program Files\Microsoft Office\Office15\VISIO.EXE",        # Visio 2019 Standalone
    "C:\Program Files (x86)\Microsoft Office\root\Office16\VISIO.EXE"  # x86 systems
)
```

**2. Expand Registry Detection** ([`visio-enterprise-audit.ps1`](visio-enterprise-audit.ps1:83))
```powershell
$RegistryPaths = @(
    "HKLM:\Software\Microsoft\Office\16.0\Common\InstallRoot",     # Office 365/2021/2019
    "HKLM:\Software\Microsoft\Office\15.0\Common\InstallRoot",      # Visio 2019
    "HKLM:\Software\Microsoft\Visio\InstallRoot",                    # Standalone Visio
    "HKLM:\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot"  # x86
)
```

**3. Enhanced WMI Query**
```powershell
# Query all Office/Visio products with version details
$officeProducts = Get-CimInstance -CimSession $session `
    -ClassName Win32_Product `
    -Filter "Name LIKE '%Office%' OR Name LIKE '%Visio%'" `
    -ErrorAction SilentlyContinue

# Categorize by version
$visio2021Standard = $officeProducts | Where-Object { 
    $_.Name -match "Visio" -and $_.Name -match "2021" -and $_.Name -notmatch "Professional"
}
$visio2021Professional = $officeProducts | Where-Object { 
    $_.Name -match "Visio.*2021.*Professional|Visio.*Professional.*2021"
}
$visio2019Professional = $officeProducts | Where-Object { 
    $_.Name -match "Visio.*2019.*Professional|Visio.*Professional.*2019"
}
```

---

### Approval Checklist

| Item | Status | Impact |
|------|--------|--------|
| Update `$VisioPaths` array | Required | Medium |
| Update `$RegistryPaths` array | Required | Medium |
| Enhance WMI detection logic | Required | Medium |
| Add version categorization fields | Required | Low |
| Test against all three Visio versions | Required | Testing |
| Update HTML report to show edition type | Optional | Low |

---

### Risk Assessment: 🟢 LOW

- Changes only **add** detection capabilities
- No breaking changes to existing functionality
- Backward compatible with current Office 365/2019 detection
- All modifications are additive

---

### Testing Requirements

Before deployment, verify:
1. [ ] Visio 2021 Standard (2408 x64) is detected correctly
2. [ ] Visio 2021 Professional (2408 x64) is detected correctly
3. [ ] Visio Professional 2019 is detected correctly
4. [ ] Existing Office 365/2019 detection still works
5. [ ] CSV/HTML reports show correct version strings

---

**Recommendation:** Approve implementation. The changes are low-risk and essential for complete inventory accuracy across your mixed Visio version environment.