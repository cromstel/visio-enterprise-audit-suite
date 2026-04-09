' ============================================================
' Script Name : MOVEit Filename Cleanup (Silent)
' Description :
'   Removes everything before the defined marker in filenames.
'   Designed for unattended execution (Task Scheduler / MOVEit).
' ============================================================

Option Explicit

' === CONFIGURATION ===
Dim folderPath, marker
folderPath = "C:\Path\To\Your\Moveit\Backups"
marker = "P.00992001720."

Dim fso, folder, file
Dim pos, newName

Set fso = CreateObject("Scripting.FileSystemObject")

' Exit silently if folder does not exist
If Not fso.FolderExists(folderPath) Then WScript.Quit

Set folder = fso.GetFolder(folderPath)

For Each file In folder.Files

  ' Find marker in filename
  pos = InStr(1, file.Name, marker, vbTextCompare)

  If pos > 0 Then
    newName = Mid(file.Name, pos)

    ' Rename only if different
    If file.Name <> newName Then
      On Error Resume Next
      file.Name = newName
      On Error GoTo 0
    End If
  End If

Next
