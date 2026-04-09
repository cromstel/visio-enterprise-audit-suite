' ============================================================
' Script Name : Trim Filename to MOVEit Prefix Marker
' Description :
'   Keeps filenames starting from a defined marker (e.g. P.00992001720.)
'   and removes anything before it.
' ============================================================

Option Explicit

' === CONFIGURATION ===
Dim folderPath, marker
folderPath = "C:\Path\To\Your\Moveit\Backups"
marker = "P.00992001720."

Dim fso, folder, file
Dim pos, newName

Set fso = CreateObject("Scripting.FileSystemObject")

' Validate folder
If Not fso.FolderExists(folderPath) Then
  WScript.Echo "ERROR: Folder not found -> " & folderPath
  WScript.Quit 1
End If

Set folder = fso.GetFolder(folderPath)

WScript.Echo "Starting filename cleanup..."
WScript.Echo "Marker: " & marker
WScript.Echo "--------------------------------------"

For Each file In folder.Files

  ' Find position of marker inside filename
  pos = InStr(1, file.Name, marker, vbTextCompare)

  If pos > 0 Then
    ' Extract filename starting from marker
    newName = Mid(file.Name, pos)

    ' Only rename if needed
    If file.Name <> newName Then
      
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
  End If

Next

WScript.Echo "--------------------------------------"
WScript.Echo "Done."
