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
Dim startPos, endPos, newName, fullNewPath

Set fso = CreateObject("Scripting.FileSystemObject")

' Exit silently if folder not found
If Not fso.FolderExists(folderPath) Then WScript.Quit

Set folder = fso.GetFolder(folderPath)

For Each file In folder.Files

  ' Process only .dat files
  If LCase(fso.GetExtensionName(file.Name)) = "dat" Then

    ' Find marker
    startPos = InStr(1, file.Name, marker, vbTextCompare)

    If startPos > 0 Then
      
      ' Find ".dat" after marker
      endPos = InStr(startPos, file.Name, ".dat", vbTextCompare)

      If endPos > 0 Then
        
        ' Build clean filename
        newName = Mid(file.Name, startPos, (endPos - startPos) + 4)

        ' Rename only if different
        If file.Name <> newName Then
          
          fullNewPath = folder.Path & "\" & newName

          ' Prevent overwrite
          If Not fso.FileExists(fullNewPath) Then
            On Error Resume Next
            file.Name = newName
            On Error GoTo 0
          End If

        End If

      End If

    End If

  End If

Next