# Visio-Helper-Utils.ps1 - PowerShell 5.1 Compatibility Review

## Review Date: 2025-01-09

## Executive Summary

**Status:** ⚠️ **ENCODING ISSUES FOUND**

The file `Visio-Helper-Utils.ps1` has **encoding issues** that will cause display problems in PowerShell 5.1, but **no PowerShell 7+ syntax features** or **syntax errors** were found.

---

## Issues Found

### 1. Encoding Issues - Box-Drawing and Tree-Drawing Characters

**Severity:** ⚠️ Medium (Display issues, not functional)

**Description:**
The file contains Unicode box-drawing and tree-drawing characters that are UTF-8 encoded without a BOM (Byte Order Mark). PowerShell 5.1 does not automatically detect UTF-8 encoding without a BOM, causing these characters to display as garbled text.

**Affected Lines:**

| Line | Character Type | Original | Displayed in PS 5.1 |
|------|---------------|----------|---------------------|
| 22-24 | Box-drawing | ╔════════════════════════════════════════════════════════════════╗ | ����������������������������������������������������������������ͻ |
| 144-146 | Box-drawing | ╔════════════════════════════════════════════════════════════════╗ | ����������������������������������������������������������������ͻ |
| 211-213 | Box-drawing | ╔════════════════════════════════════════════════════════════════╗ | ����������������������������������������������������������������ͻ |
| 249-251 | Box-drawing | ╔════════════════════════════════════════════════════════════════╗ | ����������������������������������������������������������������ͻ |
| 163 | Tree-drawing | ├─ | �� |
| 164 | Tree-drawing | └─ | �� |
| 167 | Tree-drawing | └─ | �� |
| 261 | Dash | ─────────────────────────────────── | ������������������������������������ |

**Root Cause:**
The file is saved as UTF-8 without a BOM. When PowerShell 5.1 reads the file using the default system encoding (CP1252 on Windows), it misinterprets the UTF-8 byte sequences as individual CP1252 characters.

**Impact:**
- The decorative box headers will display as garbled text
- The tree structure in summary output will display incorrectly
- **Functionality is NOT affected** - only visual display

**Solution:**
Add a UTF-8 BOM to the file so PowerShell 5.1 can properly detect the encoding.

---

## Categories Checked

### ✅ 1. PowerShell 7+ Only Features

**Status:** None found

The following PowerShell 7+ features were checked and **NOT found**:
- ❌ Null-coalescing operators (`??`)
- ❌ Null-conditional indexing (`$null?.property`)
- ❌ Ternary operators (`? :`)
- ❌ Inline hash tables in pipelines
- ❌ Cmdlets/parameters introduced in PowerShell 7+

### ✅ 2. Syntax Errors

**Status:** None found

- ✅ All brackets, parentheses, and braces are properly matched
- ✅ No missing commas or semicolons
- ✅ All command syntax is correct
- ✅ All control structures are properly formed

**Syntax Check Result:** PASSED

### ⚠️ 3. Encoding Issues

**Status:** Issues found (see above)

- ⚠️ Unicode box-drawing characters (╔, ║, ╚, ═)
- ⚠️ Unicode tree-drawing characters (├, └)
- ⚠️ Unicode dash characters (─)
- ⚠️ File is UTF-8 without BOM

### ✅ 4. Variable References

**Status:** None found

- ✅ All variables are properly defined before use
- ✅ No typos in variable names
- ✅ Correct scope usage

### ✅ 5. String Interpolation

**Status:** None found

- ✅ All string interpolation is correct
- ✅ No escape character issues
- ✅ No quoting problems in strings

### ✅ 6. Compatibility Checks

**Status:** All cmdlets are PowerShell 5.1 compatible

**Cmdlets Used:**
- Clear-Host ✅
- Compare-Reports (custom function) ✅
- Export-Csv ✅
- Export-ToExcel (custom function) ✅
- Find-UnusedVisio (custom function) ✅
- Format-Table ✅
- Get-ChildItem ✅
- Get-Content ✅
- Group-Object ✅
- Import-Csv ✅
- Invoke-FullAudit (custom function) ✅
- Join-Path ✅
- New-CostAnalysis (custom function) ✅
- New-ScheduledAudit (custom function) ✅
- New-ScheduledTaskAction ✅
- New-ScheduledTaskTrigger ✅
- Pause ✅
- Read-Host ✅
- Register-ScheduledTask ✅
- Select-Object ✅
- Select-ReportByDepartment (custom function) ✅
- Send-EmailReport (custom function) ✅
- Send-MailMessage ✅
- Show-Menu (custom function) ✅
- Show-ReportSummary (custom function) ✅
- Sort-Object ✅
- Split-Path ✅
- Start-InteractiveMenu (custom function) ✅
- Start-Sleep ✅
- Test-Path ✅
- Where-Object ✅
- Write-Host ✅

All cmdlets are available in PowerShell 5.1.

---

## Recommendations

### Immediate Action Required

1. **Add UTF-8 BOM to the file**
   - This will allow PowerShell 5.1 to properly detect the encoding
   - The Unicode characters will display correctly
   - No code changes required

### Optional Improvements

2. **Replace Unicode characters with ASCII equivalents** (if BOM solution is not preferred)
   - Replace box-drawing characters with simple dashes and pipes
   - Replace tree-drawing characters with simple dashes
   - This would ensure compatibility across all systems regardless of encoding

---

## Corrected Code

### Option 1: Add UTF-8 BOM (Recommended)

**Action:** Save the file with UTF-8 with BOM encoding

**Command to add BOM:**
```powershell
$content = [System.IO.File]::ReadAllBytes('Visio-Helper-Utils.ps1')
$bom = [byte[]]@(0xEF, 0xBB, 0xBF)
$newContent = $bom + $content
[System.IO.File]::WriteAllBytes('Visio-Helper-Utils.ps1', $newContent)
```

### Option 2: Replace Unicode Characters with ASCII

**Lines 22-24 (Menu Header):**
```powershell
# Original:
Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║       VISIO ENTERPRISE AUDIT - HELPER UTILITIES                ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

# Corrected:
Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "       VISIO ENTERPRISE AUDIT - HELPER UTILITIES" -ForegroundColor Cyan
Write-Host "====================================================================" -ForegroundColor Cyan
```

**Lines 144-146 (Report Summary Header):**
```powershell
# Original:
Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║              LATEST AUDIT REPORT SUMMARY                       ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

# Corrected:
Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "              LATEST AUDIT REPORT SUMMARY" -ForegroundColor Cyan
Write-Host "====================================================================" -ForegroundColor Cyan
```

**Lines 211-213 (Changes Detected Header):**
```powershell
# Original:
Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                    CHANGES DETECTED                            ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

# Corrected:
Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "                    CHANGES DETECTED" -ForegroundColor Cyan
Write-Host "====================================================================" -ForegroundColor Cyan
```

**Lines 249-251 (Cost Analysis Header):**
```powershell
# Original:
Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                    VISIO LICENSE COST ANALYSIS                 ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

# Corrected:
Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "                    VISIO LICENSE COST ANALYSIS" -ForegroundColor Cyan
Write-Host "====================================================================" -ForegroundColor Cyan
```

**Lines 163-164, 167 (Tree Structure):**
```powershell
# Original:
Write-Host "  ├─ Online:                $($summary.Online)" -ForegroundColor Green
Write-Host "  └─ Offline:               $($summary.Offline)" -ForegroundColor Yellow
Write-Host "  └─ Office 365:            $($summary.Office365)" -ForegroundColor Cyan

# Corrected:
Write-Host "  - Online:                 $($summary.Online)" -ForegroundColor Green
Write-Host "  - Offline:                $($summary.Offline)" -ForegroundColor Yellow
Write-Host "  - Office 365:             $($summary.Office365)" -ForegroundColor Cyan
```

**Line 261 (Dash Line):**
```powershell
# Original:
Write-Host "  ────────────────────────────────────" -ForegroundColor Green

# Corrected:
Write-Host "  -------------------------------------" -ForegroundColor Green
```

---

## Summary

**Total Issues Found:** 1 (Encoding issue)
**Total Lines Affected:** 12
**PowerShell 7+ Features:** 0
**Syntax Errors:** 0
**Variable Reference Issues:** 0
**String Interpolation Issues:** 0
**Compatibility Issues:** 0 (all cmdlets are PS 5.1 compatible)

**Overall Compatibility Status:** ⚠️ **Compatible with encoding fix required**

The file is functionally compatible with PowerShell 5.1, but requires either:
1. Adding a UTF-8 BOM to the file (recommended), OR
2. Replacing Unicode characters with ASCII equivalents

Once the encoding issue is resolved, the file will be fully compatible with PowerShell 5.1.
