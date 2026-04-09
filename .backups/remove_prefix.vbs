' ============================================================
' Script Name : Remove Prefix from MOVEit Backup Files
' Author      : BNP IT Team
' Version     : 1.0
' Date        : 2026-04-09
'
' Description :
'   This script scans a specified folder and removes a defined
'   prefix from all filenames that start with that prefix.
'
' Use Case :
'   Useful for MOVEit backup files where automated processes
'   add unwanted prefixes to filenames.
'
' Requirements :
'   - Windows OS
'   - Windows Script Host (WSH) enabled
'
' Usage :
'   1. Update "folderPath" to target directory
'   2. Update "prefixToRemove"
'   3. Run via command line:
'        cscript remove_prefix.vbs
'
' Notes :
'   - Script renames files in-place
'   - Existing files with same target name may cause conflicts
'   - Errors are logged to console output
' ============================================================

Option Explicit

' === CONFIGURATION ===
Dim folderPath        ' Path to folder containing backup files
Dim prefixToRemove    ' Prefix string to remove from filenames

folderPath = "C:\Path\To\Your\Moveit\Backups"
prefixToRemove = "PREFIX_"   ' Example: "MOVEIT_"

' === OBJECTS ===
Dim fso, folder, file
Dim newName

Set fso = CreateObject("Scripting.FileSystemObject")

' === VALIDATE FOLDER ===
If Not fso.FolderExists(folderPath) Then
  WScript.Echo "ERROR: Folder not found -> " & folderPath
  WScript.Quit 1
End If

Set folder = fso.GetFolder(folderPath)

WScript.Echo "Starting prefix removal..."
WScript.Echo "Target folder: " & folderPath
WScript.Echo "Prefix to remove: " & prefixToRemove
WScript.Echo "--------------------------------------"

' === PROCESS FILES ===
For Each file In folder.Files

  ' Check if filename starts with the specified prefix (case-insensitive)
  If LCase(Left(file.Name, Len(prefixToRemove))) = LCase(prefixToRemove) Then
    
    ' Generate new filename by removing the prefix
    newName = Mid(file.Name, Len(prefixToRemove) + 1)

    ' Attempt rename with basic error handling
    On Error Resume Next
    file.Name = newName

    If Err.Number <> 0 Then
      WScript.Echo "FAILED : " & file.Name & " -> " & newName
      Err.Clear
    Else
      WScript.Echo "RENAMED: " & file.Name & " -> " & newName
    End If
    On Error GoTo 0

  End If

Next

WScript.Echo "--------------------------------------"
WScript.Echo "Process completed."